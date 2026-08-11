// Public SDK contract. This is the only surface external hosts should use.

import type {
  CellAddress,
  CellPrimitive,
  CellStyle,
  CellValue,
  ChangeEvent,
  Unsubscribe,
  WorkbookSnapshot,
  SupportedWorkbookSnapshot,
} from "@opensheet/shared";
import type { WorksheetView } from "@opensheet/core";
import type {
  ApplyOperationsRequest,
  ApplyOperationsResult,
} from "@opensheet/commands";

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

export interface ImportCSVResult {
  sheetId: string;
  rows: number;
  columns: number;
  warnings: string[];
  truncated: boolean;
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

  importCSV(options: { file: File | Blob | string }): Promise<ImportCSVResult>;

  exportCSV(options: { sheetId: string }): Promise<Blob>;

  undo(): void;

  redo(): void;

  /** Extension beyond the base spec: renderer/UI subscribe to merged change events. */
  onChange(listener: (event: ChangeEvent) => void): Unsubscribe;

  /** UI-facing readonly accessor: renderers consume WorksheetView only. */
  getWorksheetView(sheetId: string): WorksheetView;

  /** Resolve a style id to its (readonly) style; undefined if unknown. */
  resolveStyle(styleId: string): Readonly<CellStyle> | undefined;
}

export type { ApplyOperationsRequest, ApplyOperationsResult, CellPrimitive };
