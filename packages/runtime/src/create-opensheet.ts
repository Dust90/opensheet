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
  type Unsubscribe,
  type WorkbookSnapshot,
} from "@opensheet/shared";
import type { ImportCSVResult, OpenSheetAPI, SheetInfo, WorkbookInfo } from "./api.js";

export interface OpenSheetOptions {
  history?: HistoryOptions;
}

interface WorkbookEntry {
  workbook: Workbook;
  bus: CommandBus;
  history: HistoryManager;
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
    const entry: WorkbookEntry = { workbook, bus, history };
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

    loadWorkbook(snapshot: WorkbookSnapshot) {
      const workbook = workbookFromSnapshot(snapshot);
      registerEntry(workbook);
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
      const sheet = getEntry().workbook.getSheet(sheetId);
      const matches: CellAddress[] = [];
      for (const [row, col, data] of sheet.cellEntries()) {
        const text = typeof data.value === "string" ? data.value : String(data.value ?? "");
        const hit = mode === "exact" ? text === query : text.includes(query);
        if (hit) matches.push({ row, col });
      }
      return matches;
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

    resolveStyle(styleId: string) {
      return getEntry().workbook.styles.get(styleId);
    },
  };

  return api;
}
