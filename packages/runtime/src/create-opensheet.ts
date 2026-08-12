// createOpenSheet: composition root wiring core + commands + history (+ plugins).

import {
  ApplyOperationsError,
  CommandBus,
  createDefaultRegistry,
  type ApplyOperationsResult,
} from "@injoysai/opensheet-commands";
import { toWorkbookSnapshot, Workbook, workbookFromSnapshot, Worksheet } from "@injoysai/opensheet-core";
import { HistoryManager, type HistoryOptions } from "@injoysai/opensheet-history";
import {
  createBrowserCSVWorker,
  CSVWorkerTaskHandler,
  stringifyCSV,
  validateCSVOptions,
  type CSVWorkerRequest,
  type CSVWorkerResponse,
  type CSVWorkerTransport,
} from "@injoysai/opensheet-import-export";
import {
  createPluginHost,
  type PluginCommandContext,
  type PluginFormulaFunction,
  type PluginHost,
  type PluginOperation,
} from "@injoysai/opensheet-plugin-api";
import {
  parseRange,
  CELL_ERROR_TYPES,
  SheetError,
  isCellError,
  isSheetError,
  validateFindOptions,
  type CellAddress,
  type CellValue,
  type ChangeEvent,
  type ChangeListener,
  type FindOptions,
  type SupportedWorkbookSnapshot,
  type Unsubscribe,
  type WorkbookSnapshot,
} from "@injoysai/opensheet-shared";
import type { FunctionImpl } from "@injoysai/opensheet-formula-engine";
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
  const reservedFormulaFunctionNames = new FormulaEngine().listFunctionNames();
  const pluginHost: PluginHost = createPluginHost({
    reservedCommandIds: createDefaultRegistry().list(),
    reservedFunctionNames: reservedFormulaFunctionNames,
  });
  let currentId = "";

  const listeners = new Set<ChangeListener>();

  function registerEntry(workbook: Workbook): WorkbookEntry {
    const history = new HistoryManager(options?.history);
    const bus = new CommandBus(workbook, {
      history,
      registry: createDefaultRegistry(),
    });
    const formulas = new FormulaEngine(options?.formula);
    installPluginFormulaFunctions(formulas);
    // M3: incremental recalculation folds into every commit. changedFormulas
    // = cells inside changed ranges whose CellData.formula presence/source
    // changed (formula.set, literal overwrite, undo/redo, structure rewrite)
    // — detected by diffing formula sources against the graph.
    bus.addBeforeCommitHook(({ workbook: wb, changes, derived }) => {
      const changedFormulas: Array<{ sheetId: string; row: number; col: number }> = [];
      // M4: filter changes never reach the engine — values are untouched, so
      // the dirty subgraph is empty by definition.
      // A sheet.create/sheet.import undo removes its sheet before replay hooks
      // run. Such a structure event is still useful to observers, but it has
      // no surviving worksheet for FormulaEngine to inspect or recalculate.
      const sheetViews = wb.listSheetViews();
      const recalcChanges = changes.filter(
        (change) => change.kind !== "filter" && sheetViews.some((sheet) => sheet.id === change.sheetId),
      );
      if (recalcChanges.length === 0) return;

      for (const change of recalcChanges) {
        const sheetView = sheetViews.find((s) => s.id === change.sheetId);
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
  function findCellsForSheet(sheetId: string, options: FindOptions): CellAddress[] {
    validateFindOptions(options);
    const sheet = getEntry().workbook.getSheet(sheetId);
    if (options.scope === "all" || sheet.filter === null) return findCells(sheet, options);
    const filter = sheet.filter;
    const visible = new Set(evaluateVisibleRows(sheet, filter));
    return findCells(sheet, options, row => row < filter.range.startRow || row > filter.range.endRow || visible.has(row));
  }
  function validateFindAnchor(from: CellAddress | undefined, rowCount: number, columnCount: number): void {
    if (from === undefined) return;
    if (!Number.isSafeInteger(from.row) || !Number.isSafeInteger(from.col) || from.row < 0 || from.col < 0) {
      throw new SheetError("E_VALIDATION", "findNext.from must contain non-negative safe integer row/col");
    }
    if (from.row >= rowCount || from.col >= columnCount) throw new SheetError("E_INVALID_RANGE", "findNext.from is outside worksheet bounds");
  }

  async function streamCSVRows(
    file: Blob,
    delimiter: string | undefined,
    onRows: (rows: readonly string[][]) => void,
  ): Promise<void> {
    const taskId = crypto.randomUUID();
    let transport: CSVWorkerTransport;
    try {
      transport = createBrowserCSVWorker() ?? createLocalCSVWorkerTransport();
    } catch (error) {
      throw new SheetError(
        "E_OP_FAILED",
        error instanceof Error && error.message.length > 0 ? error.message : "CSV Worker could not be created",
      );
    }
    let workerFailure: SheetError | undefined;
    let resolveCompletion: (() => void) | undefined;
    let rejectCompletion: ((error: unknown) => void) | undefined;
    const completion = new Promise<void>((resolve, reject) => {
      resolveCompletion = resolve;
      rejectCompletion = reject;
    });
    // The worker may fail while the main thread is awaiting the next Blob
    // chunk. Keep the rejection observed until the normal completion await.
    void completion.catch(() => undefined);
    const onMessage = (event: MessageEvent<CSVWorkerResponse>) => {
      const response = event.data;
      if (response.taskId !== taskId) return;
      if (response.type === "rows") {
        onRows(response.rows);
      } else if (response.type === "error") {
        workerFailure = new SheetError(response.code, response.message);
        rejectCompletion?.(workerFailure);
      } else {
        resolveCompletion?.();
      }
    };
    const failWorker = (message: string) => {
      if (workerFailure !== undefined) return;
      workerFailure = new SheetError("E_OP_FAILED", message);
      rejectCompletion?.(workerFailure);
    };
    const onWorkerError = (event: ErrorEvent) => {
      failWorker(event.message || "CSV Worker failed");
    };
    const onWorkerMessageError = () => {
      failWorker("CSV Worker message could not be deserialized");
    };
    transport.addEventListener("message", onMessage);
    transport.addEventListener("error", onWorkerError);
    transport.addEventListener("messageerror", onWorkerMessageError);
    const send = (request: CSVWorkerRequest) => transport.postMessage(request);
    const throwIfWorkerFailed = () => {
      if (workerFailure !== undefined) throw workerFailure;
    };
    const reader = file.stream().getReader();
    const decoder = new TextDecoder();
    try {
      send(delimiter === undefined ? { type: "start", taskId } : { type: "start", taskId, options: { delimiter } });
      throwIfWorkerFailed();
      while (true) {
        const next = await reader.read();
        throwIfWorkerFailed();
        if (next.done) break;
        const text = decoder.decode(next.value, { stream: true });
        // Bound each protocol message even if a custom Blob stream yields a
        // very large Uint8Array in one read.
        for (let offset = 0; offset < text.length; offset += 64 * 1024) {
          send({ type: "chunk", taskId, text: text.slice(offset, offset + 64 * 1024) });
          throwIfWorkerFailed();
        }
      }
      const finalText = decoder.decode();
      for (let offset = 0; offset < finalText.length; offset += 64 * 1024) {
        send({ type: "chunk", taskId, text: finalText.slice(offset, offset + 64 * 1024) });
        throwIfWorkerFailed();
      }
      send({ type: "finish", taskId });
      throwIfWorkerFailed();
      await completion;
    } catch (caught) {
      try {
        send({ type: "cancel", taskId });
      } catch {
        // A native Worker failure may make postMessage unavailable. terminate
        // below still releases the browser task and staging never escaped.
      }
      throw caught;
    } finally {
      reader.releaseLock();
      transport.removeEventListener?.("message", onMessage);
      transport.removeEventListener?.("error", onWorkerError);
      transport.removeEventListener?.("messageerror", onWorkerMessageError);
      transport.terminate?.();
    }
  }

  function createLocalCSVWorkerTransport(): CSVWorkerTransport {
    const tasks = new CSVWorkerTaskHandler();
    const listeners = new Set<(event: MessageEvent<CSVWorkerResponse>) => void>();
    const emit = (response: CSVWorkerResponse) => {
      const event = { data: response } as MessageEvent<CSVWorkerResponse>;
      for (const listener of listeners) listener(event);
    };
    const transport = {
      postMessage(request: CSVWorkerRequest) { tasks.handle(request, emit); },
      addEventListener(type: string, listener: unknown) {
        if (type === "message") listeners.add(listener as (event: MessageEvent<CSVWorkerResponse>) => void);
      },
      removeEventListener(type: string, listener: unknown) {
        if (type === "message") listeners.delete(listener as (event: MessageEvent<CSVWorkerResponse>) => void);
      },
      terminate() { listeners.clear(); },
    };
    return transport as CSVWorkerTransport;
  }

  function validateImportCSVOptions(value: unknown): asserts value is { file: Blob; delimiter?: string } {
    if (typeof value !== "object" || value === null || Array.isArray(value)) {
      throw new SheetError("E_VALIDATION", "importCSV options must be an object");
    }
    const options = value as Record<string, unknown>;
    for (const key of Object.keys(options)) {
      if (key !== "file" && key !== "delimiter") {
        throw new SheetError("E_VALIDATION", `importCSV options contains unknown field \"${key}\"`);
      }
    }
    if (typeof Blob === "undefined" || !(options.file instanceof Blob)) {
      throw new SheetError("E_VALIDATION", "importCSV.file must be a Blob");
    }
    validateCSVOptions(options.delimiter === undefined ? {} : { delimiter: options.delimiter });
  }

  function importedSheetName(file: Blob, workbook: Workbook): string {
    const candidate = file as Blob & { name?: unknown };
    const rawName = typeof candidate.name === "string" ? candidate.name.trim() : "";
    const basename = rawName.split(/[\\/]/).at(-1) ?? "";
    const base = (basename.replace(/\.csv$/i, "").trim() || "Imported CSV");
    const names = new Set(workbook.listSheets().map((sheet) => sheet.name));
    if (!names.has(base)) return base;
    let suffix = 2;
    while (names.has(`${base} (${suffix})`)) suffix += 1;
    return `${base} (${suffix})`;
  }

  function validateExportCSVOptions(value: unknown): asserts value is { sheetId: string } {
    if (typeof value !== "object" || value === null || Array.isArray(value)) {
      throw new SheetError("E_VALIDATION", "exportCSV options must be an object");
    }
    const options = value as Record<string, unknown>;
    if (Object.keys(options).some((key) => key !== "sheetId")) {
      throw new SheetError("E_VALIDATION", "exportCSV options contains an unknown field");
    }
    if (typeof options.sheetId !== "string" || options.sheetId.length === 0) {
      throw new SheetError("E_VALIDATION", "exportCSV.sheetId must be a non-empty string");
    }
  }

  function csvValueText(value: CellValue): string {
    if (value === null) return "";
    if (isCellError(value)) return value.type;
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    return String(value);
  }

  function executePluginCommand<TResult>(
    entry: WorkbookEntry,
    commandId: string,
    payload: unknown,
    options: { sheetId?: string; source?: "user" | "api" | "undo" | "redo" },
  ): TResult {
    const source = options.source ?? "api";
    pluginHost.emitBeforeCommand({ commandId, source });
    const result = entry.bus.execute<TResult>(commandId, payload, options);
    pluginHost.emitAfterCommand({ commandId, source });
    return result;
  }

  function pluginCommandContext(entry: WorkbookEntry, sheetId: string): PluginCommandContext {
    const sheet = entry.workbook.getSheet(sheetId);
    return {
      workbookId: entry.workbook.id,
      sheetId,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      getCell(row, col) {
        if (!Number.isSafeInteger(row) || !Number.isSafeInteger(col) || row < 0 || col < 0 || row >= sheet.rowCount || col >= sheet.columnCount) {
          throw new SheetError("E_INVALID_RANGE", "Plugin command cell address is outside worksheet bounds");
        }
        const value = sheet.getCell(row, col)?.value ?? null;
        return isCellError(value) ? { ...value } : value;
      },
    };
  }

  function pluginCommandOperations(value: unknown): PluginOperation[] {
    if (!Array.isArray(value)) {
      throw new SheetError("E_VALIDATION", "Plugin command execute() must return an operation array");
    }
    return value as PluginOperation[];
  }

  function installPluginFormulaFunctions(formulas: FormulaEngine): void {
    for (const contribution of pluginHost.listFunctionContributions()) {
      if (contribution.execute === undefined) continue;
      formulas.registerPluginFunction(
        contribution.name,
        pluginFormulaImplementation(contribution.name, contribution.minArgs, contribution.maxArgs, contribution.execute),
      );
    }
  }

  function pluginFormulaImplementation(
    name: string,
    minArgs: number,
    maxArgs: number,
    execute: PluginFormulaFunction,
  ): FunctionImpl {
    return (args) => {
      if (args.length < minArgs || args.length > maxArgs) {
        return { type: "#VALUE!", message: `${name} needs ${minArgs}-${maxArgs} arguments` };
      }
      try {
        return normalizePluginFormulaValue(execute(args));
      } catch (error) {
        if (isKnownCellError(error)) {
          return { ...error };
        }
        return { type: "#VALUE!", message: `Plugin function ${name} failed` };
      }
    };
  }

  function normalizePluginFormulaValue(value: unknown): CellValue {
    if (value === null || typeof value === "string" || typeof value === "boolean") return value;
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : { type: "#NUM!", message: "Plugin function returned a non-finite number" };
    }
    if (isKnownCellError(value)) {
      return value.message === undefined ? { type: value.type } : { type: value.type, message: value.message };
    }
    return { type: "#VALUE!", message: "Plugin function returned an invalid CellValue" };
  }

  function isKnownCellError(value: unknown): value is Extract<CellValue, object> {
    const candidate = value as { type?: unknown; message?: unknown } | null;
    return (
      typeof value === "object" &&
      value !== null &&
      typeof candidate?.type === "string" &&
      CELL_ERROR_TYPES.includes(candidate.type as (typeof CELL_ERROR_TYPES)[number]) &&
      (candidate.message === undefined || typeof candidate.message === "string")
    );
  }

  function recalculateAllFormulas(entry: WorkbookEntry): void {
    entry.formulas.rebuildAndRecalculateAll(entry.workbook.asView(), makeDerivedBridge(entry));
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
      const result = executePluginCommand<{ sheetId: string; name: string; rowCount: number; columnCount: number }>(
        entry,
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
      return findCellsForSheet(sheetId, { query, matchCase: true, wholeCell: mode === "exact", searchIn: "values", scope: "all", direction: "forward" });
    },

    findCells(options: { sheetId: string } & FindOptions): CellAddress[] {
      return findCellsForSheet(options.sheetId, options);
    },

    findNext(options: { sheetId: string; from?: CellAddress } & FindOptions): CellAddress | null {
      const sheet = getEntry().workbook.getSheet(options.sheetId);
      validateFindAnchor(options.from, sheet.rowCount, sheet.columnCount);
      const matches = findCellsForSheet(options.sheetId, options);
      if (matches.length === 0) return null;
      if (options.from === undefined) return matches[0]!;
      const compare = (a: CellAddress, b: CellAddress) => a.row - b.row || a.col - b.col;
      const next = options.direction === "forward"
        ? matches.find(match => compare(match, options.from!) > 0)
        : matches.find(match => compare(match, options.from!) < 0);
      return next ?? matches[0]!;
    },

    async importCSV(options): Promise<ImportCSVResult> {
      validateImportCSVOptions(options);
      const entry = getEntry();
      let rowCount = 0;
      let columnCount = 0;

      // First pass validates the complete stream and finds exact worksheet
      // dimensions without retaining parsed rows. Blob is replayable, so the
      // second pass can write directly into an isolated staging Worksheet.
      await streamCSVRows(options.file, options.delimiter, (rows) => {
        for (const row of rows) {
          rowCount += 1;
          columnCount = Math.max(columnCount, row.length);
        }
      });

      const sheet = new Worksheet({
        id: crypto.randomUUID(),
        name: importedSheetName(options.file, entry.workbook),
        // Empty CSV still creates an interactable empty worksheet. The result
        // reports parsed dimensions (0 × 0), while the model remains valid.
        rowCount: Math.max(1, rowCount),
        columnCount: Math.max(1, columnCount),
      });
      let row = 0;
      await streamCSVRows(options.file, options.delimiter, (rows) => {
        for (const values of rows) {
          for (let col = 0; col < values.length; col += 1) {
            sheet.setCell(row, col, { value: values[col]! });
          }
          row += 1;
        }
      });

      // The staging sheet only becomes observable through this single command.
      // If either parse/write pass rejects, no existing workbook state changed.
      const result = executePluginCommand<{ sheetId: string; rowCount: number; columnCount: number }>(
        entry,
        "sheet.import",
        { sheet },
        { source: "api" },
      );
      return { sheetId: result.sheetId, rowCount, columnCount };
    },

    async exportCSV(options): Promise<Blob> {
      validateExportCSVOptions(options);
      const sheet = getEntry().workbook.getSheet(options.sheetId);
      let bottomRow = -1;
      let rightColumn = -1;
      // CSV exports values only. A formula's source and style-only cells do
      // not expand this range; an explicit empty string does.
      for (const [row, col, data] of sheet.cellEntries()) {
        if (data.value === null) continue;
        bottomRow = Math.max(bottomRow, row);
        rightColumn = Math.max(rightColumn, col);
      }
      if (bottomRow < 0 || rightColumn < 0) {
        return new Blob([], { type: "text/csv;charset=utf-8" });
      }
      const rows: string[][] = [];
      for (let row = 0; row <= bottomRow; row += 1) {
        const values: string[] = [];
        for (let col = 0; col <= rightColumn; col += 1) {
          values.push(csvValueText(sheet.getCell(row, col)?.value ?? null));
        }
        rows.push(values);
      }
      return new Blob([stringifyCSV(rows)], { type: "text/csv;charset=utf-8" });
    },

    async usePlugin(plugin) {
      const beforeNames = new Set(pluginHost.listFunctionContributions().map((contribution) => contribution.name));
      await pluginHost.use(plugin);
      const added = pluginHost.listFunctionContributions().filter(
        (contribution) => contribution.execute !== undefined && !beforeNames.has(contribution.name),
      );
      if (added.length === 0) return;
      try {
        for (const entry of entries.values()) {
          for (const contribution of added) {
            entry.formulas.registerPluginFunction(
              contribution.name,
              pluginFormulaImplementation(contribution.name, contribution.minArgs, contribution.maxArgs, contribution.execute!),
            );
          }
          recalculateAllFormulas(entry);
        }
      } catch (error) {
        for (const entry of entries.values()) {
          for (const contribution of added) entry.formulas.unregisterPluginFunction(contribution.name);
        }
        await pluginHost.dispose(plugin.id);
        throw error;
      }
    },

    async disposePlugin(pluginId) {
      const before = pluginHost.listFunctionContributions().filter((contribution) => contribution.execute !== undefined);
      await pluginHost.dispose(pluginId);
      const retained = new Set(pluginHost.listFunctionContributions().map((contribution) => contribution.name));
      const removed = before.filter((contribution) => !retained.has(contribution.name));
      if (removed.length === 0) return;
      for (const entry of entries.values()) {
        for (const contribution of removed) entry.formulas.unregisterPluginFunction(contribution.name);
        recalculateAllFormulas(entry);
      }
    },

    getPluginContributions() {
      return {
        commands: pluginHost.listCommandContributions(),
        functions: pluginHost.listFunctionContributions(),
        menus: pluginHost.listMenuContributions(),
      };
    },

    async executePluginCommand(options) {
      if (typeof options !== "object" || options === null || Array.isArray(options)) {
        throw new SheetError("E_VALIDATION", "executePluginCommand options must be an object");
      }
      if (typeof options.workbookId !== "string" || typeof options.sheetId !== "string" || typeof options.commandId !== "string" || options.commandId.length === 0) {
        throw new SheetError("E_VALIDATION", "executePluginCommand requires workbookId, sheetId and commandId");
      }
      const entry = getEntry(options.workbookId);
      const contribution = pluginHost.listCommandContributions().find((command) => command.id === options.commandId);
      if (contribution === undefined) {
        throw new SheetError("E_UNKNOWN_COMMAND", `Plugin command not found: ${options.commandId}`);
      }
      if (contribution.execute === undefined) {
        throw new SheetError("E_NOT_IMPLEMENTED", `Plugin command has no executable handler: ${options.commandId}`);
      }
      pluginHost.emitBeforeCommand({ commandId: options.commandId, source: "api" });
      try {
        contribution.validate?.(options.payload);
        const operations = pluginCommandOperations(contribution.execute(pluginCommandContext(entry, options.sheetId), options.payload));
        const result = entry.bus.applyOperations({
          sheetId: options.sheetId,
          operations,
          atomic: true,
          source: "api",
        });
        pluginHost.emitAfterCommand({ commandId: options.commandId, source: "api" });
        return result;
      } catch (error) {
        if (isSheetError(error) || error instanceof ApplyOperationsError) throw error;
        throw new SheetError("E_OP_FAILED", error instanceof Error ? error.message : String(error));
      }
    },

    undo() {
      const entry = getEntry();
      pluginHost.emitBeforeCommand({ commandId: "history.undo", source: "undo" });
      entry.history.undo(entry.bus);
      pluginHost.emitAfterCommand({ commandId: "history.undo", source: "undo" });
    },

    redo() {
      const entry = getEntry();
      pluginHost.emitBeforeCommand({ commandId: "history.redo", source: "redo" });
      entry.history.redo(entry.bus);
      pluginHost.emitAfterCommand({ commandId: "history.redo", source: "redo" });
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
