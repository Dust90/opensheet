// Read-only view contracts. Renderers (M1) and external read paths depend on
// these views only — mutation is possible exclusively through the Command Bus.

import type { CellData, Range } from "@opensheet/shared";

/**
 * Read-only view of a worksheet. Implementations may return internal objects
 * typed as Readonly — callers must treat them as frozen. (Dev builds may add
 * Object.freeze hardening later; the type boundary is the contract.)
 */
export interface WorksheetView {
  readonly id: string;
  readonly name: string;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly frozenRows: number;
  readonly frozenColumns: number;
  readonly cellCount: number;

  getCell(row: number, col: number): Readonly<CellData> | undefined;

  cellEntries(): IterableIterator<[row: number, col: number, data: Readonly<CellData>]>;

  forEachCellInRange(
    range: Range,
    callback: (row: number, col: number, data: Readonly<CellData>) => void,
  ): void;

  getRowHeight(row: number): number | undefined;

  getColumnWidth(col: number): number | undefined;
}
