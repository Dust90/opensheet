// Public SDK contract. This is the only surface external hosts should use.

import type {
  CellAddress,
  CellPrimitive,
  CellStyle,
  CellValue,
  ChangeEvent,
  FilterSpec,
  FindOptions,
  Unsubscribe,
  WorkbookSnapshot,
  SupportedWorkbookSnapshot,
} from "@opensheet/shared";
import type { WorksheetView } from "@opensheet/core";
import type {
  ApplyOperationsRequest,
  ApplyOperationsResult,
} from "@opensheet/commands";
import type {
  CommandContribution,
  FormulaFunctionContribution,
  MenuItemContribution,
  OpenSheetPlugin,
} from "@opensheet/plugin-api";

export interface WorkbookInfo {
  id: string;
  name: string;
  activeSheetId: string;
}

export interface SheetInfo {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
}

/** Distinguishes no active filter from an active filter with zero matching rows. */
export interface FilterProjectionState {
  filter: Readonly<FilterSpec> | null;
  visibleRows: Uint32Array | null;
}

export interface ImportCSVResult {
  sheetId: string;
  rowCount: number;
  columnCount: number;
}

export interface ImportCSVOptions {
  file: Blob;
  delimiter?: string;
}

export interface ExportCSVOptions {
  sheetId: string;
}

/** Readonly snapshots of metadata registered by installed plugins. */
export interface PluginContributions {
  commands: readonly CommandContribution[];
  functions: readonly FormulaFunctionContribution[];
  menus: readonly MenuItemContribution[];
}

export interface ExecutePluginCommandOptions {
  workbookId: string;
  sheetId: string;
  commandId: string;
  payload: unknown;
}

export interface OpenSheetAPI {
  createWorkbook(options: { id?: string; name: string }): WorkbookInfo;

  loadWorkbook(snapshot: SupportedWorkbookSnapshot): WorkbookInfo;

  getWorkbookSnapshot(): WorkbookSnapshot;

  listSheets(): SheetInfo[];

  createSheet(options: { name: string; rows?: number; columns?: number }): SheetInfo;

  readRange(options: { sheetId: string; range: string }): CellValue[][];

  applyOperations(request: ApplyOperationsRequest): Promise<ApplyOperationsResult>;

  searchCells(options: {
    sheetId: string;
    query: string;
    mode: "exact" | "contains";
  }): CellAddress[];

  findCells(options: { sheetId: string } & FindOptions): CellAddress[];

  findNext(options: { sheetId: string; from?: CellAddress } & FindOptions): CellAddress | null;

  /** Import CSV into a new worksheet at A1; existing worksheets are untouched. */
  importCSV(options: ImportCSVOptions): Promise<ImportCSVResult>;

  exportCSV(options: ExportCSVOptions): Promise<Blob>;

  /** Install a metadata/hook plugin into this Runtime instance. */
  usePlugin(plugin: OpenSheetPlugin): Promise<void>;

  /** Dispose a plugin and remove every contribution it registered. */
  disposePlugin(pluginId: string): Promise<void>;

  getPluginContributions(): PluginContributions;

  /** Execute one plugin command as a single atomic built-in operation batch. */
  executePluginCommand(options: ExecutePluginCommandOptions): Promise<ApplyOperationsResult>;

  undo(): void;

  redo(): void;

  /** Extension beyond the base spec: renderer/UI subscribe to merged change events. */
  onChange(listener: (event: ChangeEvent) => void): Unsubscribe;

  /** UI-facing readonly accessor: renderers consume WorksheetView only. */
  getWorksheetView(sheetId: string): WorksheetView;

  getFilterProjectionState(sheetId: string): FilterProjectionState;

  /** Resolve a style id to its (readonly) style; undefined if unknown. */
  resolveStyle(styleId: string): Readonly<CellStyle> | undefined;
}

export type { ApplyOperationsRequest, ApplyOperationsResult, CellPrimitive };
