// createOpenSheet: composition root wiring core + commands + history (+ plugins).

import {
  CommandBus,
  createDefaultRegistry,
  type ApplyOperationsResult,
} from "@opensheet/commands";
import { toWorkbookSnapshot, Workbook, workbookFromSnapshot, Worksheet } from "@opensheet/core";
import { HistoryManager, type HistoryOptions } from "@opensheet/history";
import { createPluginHost, type PluginHost } from "@opensheet/plugin-api";
import {
  parseRange,
  SheetError,
  type CellAddress,
  type CellValue,
  type ChangeEvent,
  type ChangeListener,
  type FindOptions,
  type SupportedWorkbookSnapshot,
  type Unsubscribe,
  type WorkbookSnapshot,
} from "@opensheet/shared";
import type { ImportCSVResult, OpenSheetAPI, SheetInfo, WorkbookInfo } from "./api.js";
import { FormulaEngine, type FormulaEngineOptions } from "./formula-engine.js";
import { evaluateVisibleRows } from "./filter-engine.js";
import { findCells } from "./find-engine.js";

export interface OpenSheetOptions {
  history?: HistoryOptions;
  formula?: FormulaEngineOptions;
}

interface WorkbookEntry {
  workbook: Workbook;
  bus: CommandBus;
  history: HistoryManager;
  formulas: FormulaEngine;
}

const DEFAULT_ROWS = 1000;
const DEFAULT_COLUMNS = 26;

/**
 * Assemble a standalone OpenSheet instance. M0 wires workbook core, command
 * bus, transactions and history; formula engine (M3) and CSV (M5) plug into
 * the same seams (beforeCommit hooks / registry) without changing this API.
 */
export function createOpenSheet(options?: OpenSheetOptions): OpenSheetAPI {
  const entries = new Map<string, WorkbookEntry>();
  const pluginHost: PluginHost = createPluginHost();
  let currentId = "";

  const listeners = new Set<ChangeListener>();

  function registerEntry(workbook: Workbook): WorkbookEntry {
    const history = new HistoryManager(options?.history);
    const bus = new CommandBus(workbook, {
      history,
      registry: createDefaultRegistry(),
    });
    const formulas = new FormulaEngine(options?.formula);
    // M3: incremental recalculation folds into every commit. changedFormulas
    // = cells inside changed ranges whose CellData.formula presence/source
    // changed (formula.set, literal overwrite, undo/redo, structure rewrite)
    // — detected by diffing formula sources against the graph.
    bus.addBeforeCommitHook(({ workbook: wb, changes, derived }) => {
      const changedFormulas: Array<{ sheetId: string; row: number; col: number }> = [];
      // M4: filter changes never reach the engine — values are untouched, so
      // the dirty subgraph is empty by definition.
      const recalcChanges = changes.filter((change) => change.kind !== "filter");
      if (recalcChanges.length === 0) return;

      for (const change of recalcChanges) {
        const sheetView = wb.listSheetViews().find((s) => s.id === change.sheetId);
        if (sheetView === undefined) continue;

        if (change.kind === "rows" || change.kind === "columns" || change.kind === "reorder") {
          // Fix 2: structural command — rebuild the entire sheet graph sparsely
          // (iterates only non-empty CellStore entries, not all coordinates).
          // This corrects shifted coordinates, removes deleted-row nodes, and
          // handles out-of-bounds references atomically.
          // M4: "reorder" (sort/dedupe) moved formulas to new rows — same path.
          formulas.rebuildSheetGraph(wb, change.sheetId);
          // After rebuild, ALL formula cells on this sheet are "changed".
          for (const [row, col, data] of sheetView.cellEntries()) {
            if (data.formula !== undefined) {
              changedFormulas.push({ sheetId: change.sheetId, row, col });
            }
          }
        } else {
          // Fix 2: sparse union for ordinary cell changes —
          // 1. non-empty cells inside the changed range that carry a formula;
          // 2. stale graph nodes that fall inside the changed range.
          const seen = new Set<string>();
          const addChanged = (row: number, col: number) => {
            const key = `${row}:${col}`;
            if (seen.has(key)) return;
            seen.add(key);
            changedFormulas.push({ sheetId: change.sheetId, row, col });
          };

          // Side A: cells that NOW have a formula in this range.
          sheetView.forEachCellInRange(change.range, (row, col, data) => {
            const hasFormula = data.formula !== undefined;
            const registered = formulas.hasFormula(change.sheetId, row, col);
            if (
              hasFormula !== registered ||
              (hasFormula && data.formula !== formulas.formulaSourceOf(change.sheetId, row, col))
            ) {
              addChanged(row, col);
            }
          });

          // Side B: stale graph nodes inside this range (e.g. formula was deleted).
          for (const addr of formulas.graphFormulaCellsInRange(change.sheetId, change.range)) {
            const data = sheetView.getCell(addr.row, addr.col);
            if (data?.formula === undefined) {
              // Formula was removed — must de-register the old node.
              addChanged(addr.row, addr.col);
            }
          }
        }
      }

      formulas.recalculate(wb, recalcChanges, changedFormulas, derived);
    });
    const entry: WorkbookEntry = { workbook, bus, history, formulas };
    entries.set(workbook.id, entry);
    workbook.onChange((event) => {
      for (const listener of listeners) listener(event);
    });
    currentId = workbook.id;
    return entry;
  }

  function getEntry(workbookId?: string): WorkbookEntry {
    const id = workbookId ?? currentId;
    const entry = entries.get(id);
    if (entry === undefined) {
      throw new SheetError("E_WORKBOOK_NOT_FOUND", `Workbook not found: "${id}"`);
    }
    currentId = id;
    return entry;
  }

  function toWorkbookInfo(workbook: Workbook): WorkbookInfo {
    return { id: workbook.id, name: workbook.name, activeSheetId: workbook.activeSheetId };
  }

  /** DerivedWriter bridge for engine rebuilds outside a transaction (Snapshot load). */
  function makeDerivedBridge(entry: WorkbookEntry): { setComputedValue(sheetId: string, row: number, col: number, value: CellValue): void } {
    return {
      setComputedValue: (sheetId, row, col, value) => {
        const sheet = entry.workbook.getSheet(sheetId);
        if (row < 0 || row >= sheet.rowCount || col < 0 || col >= sheet.columnCount) return;
        const previous = sheet.getCell(row, col);
        const next = { ...(previous ?? {}), value };
        sheet.setCell(row, col, next);
        entry.workbook.emit({
          workbookId: entry.workbook.id,
          sheetId,
          changes: [{ range: { startRow: row, startCol: col, endRow: row, endCol: col }, kind: "cells" }],
          source: "derived",
          batch: false,
        });
      },
    };
  }

  function toSheetInfo(sheet: Worksheet): SheetInfo {
    return {
      id: sheet.id,
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
    };
  }

  const api: OpenSheetAPI = {
    createWorkbook({ id, name }) {
      const workbook = new Workbook({ id: id ?? crypto.randomUUID(), name });
      workbook.addSheet(
        new Worksheet({
          id: crypto.randomUUID(),
          name: "Sheet1",
          rowCount: DEFAULT_ROWS,
          columnCount: DEFAULT_COLUMNS,
        }),
      );
      registerEntry(workbook);
      pluginHost.emitWorkbookLoaded(workbook.id);
      return toWorkbookInfo(workbook);
    },

    loadWorkbook(snapshot: SupportedWorkbookSnapshot) {
      const workbook = workbookFromSnapshot(snapshot);
      const entry = registerEntry(workbook);
      // M3: rebuild the dependency graph from the loaded formulas and
      // recompute every cached value (do NOT trust the stored cache).
      const derived = makeDerivedBridge(entry);
      entry.formulas.rebuildAndRecalculateAll(workbook.asView(), derived);
      pluginHost.emitWorkbookLoaded(workbook.id);
      return toWorkbookInfo(workbook);
    },

    getWorkbookSnapshot(): WorkbookSnapshot {
      return toWorkbookSnapshot(getEntry().workbook);
    },

    listSheets(): SheetInfo[] {
      return getEntry().workbook.listSheets().map(toSheetInfo);
    },

    createSheet({ name, rows, columns }): SheetInfo {
      const entry = getEntry();
      const result = entry.bus.execute<{ sheetId: string; name: string; rowCount: number; columnCount: number }>(
        "sheet.create",
        { name, rows, columns },
        { source: "api" },
      );
      return { id: result.sheetId, name: result.name, rowCount: result.rowCount, columnCount: result.columnCount };
    },

    readRange({ sheetId, range }): CellValue[][] {
      const sheet = getEntry().workbook.getSheet(sheetId);
      const parsed = parseRange(range);
      const values: CellValue[][] = [];
      for (let row = parsed.startRow; row <= parsed.endRow; row++) {
        const rowValues: CellValue[] = [];
        for (let col = parsed.startCol; col <= parsed.endCol; col++) {
          rowValues.push(sheet.getCell(row, col)?.value ?? null);
        }
        values.push(rowValues);
      }
      return values;
    },

    async applyOperations(request): Promise<ApplyOperationsResult> {
      const entry = getEntry(request.workbookId);
      pluginHost.emitBeforeCommand({ commandId: "applyOperations", source: "api" });
      const result = entry.bus.applyOperations({
        sheetId: request.sheetId,
        operations: request.operations,
        atomic: request.atomic ?? false,
        source: "api",
      });
      pluginHost.emitAfterCommand({ commandId: "applyOperations", source: "api" });
      return result;
    },

    searchCells({ sheetId, query, mode }): CellAddress[] {
      return this.findCells({ sheetId, query, matchCase: true, wholeCell: mode === "exact", searchIn: "values", scope: "all", direction: "forward" });
    },

    findCells(options: { sheetId: string } & FindOptions): CellAddress[] {
      const sheet = getEntry().workbook.getSheet(options.sheetId);
      if (options.scope === "all" || sheet.filter === null) return findCells(sheet, options);
      const filter = sheet.filter;
      const visible = new Set(evaluateVisibleRows(sheet, filter));
      return findCells(sheet, options, row => row < filter.range.startRow || row > filter.range.endRow || visible.has(row));
    },

    importCSV(): Promise<ImportCSVResult> {
      return Promise.reject(
        new SheetError("E_NOT_IMPLEMENTED", "CSV import lands in M5 (import-export package)"),
      );
    },

    exportCSV(): Promise<Blob> {
      return Promise.reject(
        new SheetError("E_NOT_IMPLEMENTED", "CSV export lands in M5 (import-export package)"),
      );
    },

    undo() {
      const entry = getEntry();
      entry.history.undo(entry.bus);
    },

    redo() {
      const entry = getEntry();
      entry.history.redo(entry.bus);
    },

    onChange(listener: (event: ChangeEvent) => void): Unsubscribe {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },

    getWorksheetView(sheetId: string) {
      return getEntry().workbook.getSheetView(sheetId);
    },

    getFilterProjectionState(sheetId: string) {
      const sheet = getEntry().workbook.getSheet(sheetId);
      const filter = sheet.filter;
      return {
        filter,
        visibleRows: filter === null ? null : evaluateVisibleRows(sheet, filter),
      };
    },

    resolveStyle(styleId: string) {
      return getEntry().workbook.styles.get(styleId);
    },
  };

  return api;
}
