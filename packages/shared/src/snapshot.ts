// @opensheet/shared — persistence (snapshot) contracts

import type { CellData, CellStyle } from "./cell.js";

/**
 * JSON-serializable persistence model. Kept separate from the runtime model
 * (Map/CellStore-based) because:
 * - runtime maps carry non-serializable state (listeners, dependency graphs);
 * - snapshots must round-trip through JSON and support version migrations.
 */
export interface WorksheetSnapshot {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  /** Keyed by "row:col" (0-based) for unambiguous, order-stable serialization. */
  cells: Record<string, CellData>;
  rowHeights: Record<string, number>;
  columnWidths: Record<string, number>;
  frozenRows: number;
  frozenColumns: number;
}

export const WORKBOOK_SNAPSHOT_VERSION = 1;

export interface WorkbookSnapshot {
  id: string;
  name: string;
  activeSheetId: string;
  sheets: WorksheetSnapshot[];
  styles: Record<string, CellStyle>;
  version: number;
}
