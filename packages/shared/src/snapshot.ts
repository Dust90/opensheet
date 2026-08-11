// @opensheet/shared — persistence (snapshot) contracts

import type { CellData, CellStyle } from "./cell.js";
import type { FilterSpec } from "./data-operations.js";

/**
 * JSON-serializable persistence model. Kept separate from the runtime model
 * (Map/CellStore-based) because:
 * - runtime maps carry non-serializable state (listeners, dependency graphs);
 * - snapshots must round-trip through JSON and support version migrations.
 */
interface WorksheetSnapshotBase {
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

/** Legacy, strict V1 shape. V1 sheets intentionally have no filter field. */
export interface WorksheetSnapshotV1 extends WorksheetSnapshotBase {}

/** Current persistence shape. `filter` is required but nullable. */
export interface WorksheetSnapshotV2 extends WorksheetSnapshotBase {
  filter: FilterSpec | null;
}

export const WORKBOOK_SNAPSHOT_VERSION = 2;

interface WorkbookSnapshotBase<TSheet> {
  id: string;
  name: string;
  activeSheetId: string;
  sheets: TSheet[];
  styles: Record<string, CellStyle>;
}

export interface WorkbookSnapshotV1 extends WorkbookSnapshotBase<WorksheetSnapshotV1> {
  version: 1;
}

export interface WorkbookSnapshotV2 extends WorkbookSnapshotBase<WorksheetSnapshotV2> {
  version: 2;
}

/** All persisted formats accepted by the loader. */
export type SupportedWorkbookSnapshot = WorkbookSnapshotV1 | WorkbookSnapshotV2;

/** New writes always produce V2. */
export type WorkbookSnapshot = WorkbookSnapshotV2;
export type WorksheetSnapshot = WorksheetSnapshotV2;
