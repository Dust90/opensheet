// Sparse worksheet. Mutations here are SILENT — change events are emitted by
// the command layer so transactions control exactly what observers see.

import {
  SheetError,
  validateFilterSpec,
  type CellData,
  type CellStore,
  type CellStoreFactory,
  type FilterSpec,
  type Range,
} from "@opensheet/shared";
import { chunkedCellStoreFactory } from "./cell-store/chunked-store.js";
import type { WorksheetView } from "./view.js";

export interface WorksheetInit {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  storeFactory?: CellStoreFactory;
}

export class Worksheet {
  readonly id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  frozenRows = 0;
  frozenColumns = 0;
  readonly rowHeights = new Map<number, number>();
  readonly columnWidths = new Map<number, number>();

  private _filter: FilterSpec | null = null;

  private cells: CellStore;
  private readonly storeFactory: CellStoreFactory;

  constructor(init: WorksheetInit) {
    this.id = init.id;
    this.name = init.name;
    this.rowCount = init.rowCount;
    this.columnCount = init.columnCount;
    // Default frozen by ADR-0005 (benchmark-driven, chunked won on all
    // speed axes at equal memory). Overridable per sheet.
    this.storeFactory = init.storeFactory ?? chunkedCellStoreFactory;
    this.cells = this.storeFactory.create();
  }

  get cellCount(): number {
    return this.cells.size;
  }

  /**
   * Return a detached FilterSpec so read consumers cannot mutate Worksheet
   * state outside the Command Bus. `Readonly` is the compile-time contract;
   * cloning is the runtime boundary.
   */
  get filter(): Readonly<FilterSpec> | null {
    return this._filter === null ? null : cloneFilterSpec(this._filter);
  }

  getFilter(): Readonly<FilterSpec> | null {
    return this.filter;
  }

  /** Silent mutation for command/journal use; commands own events/history. */
  setFilter(filter: FilterSpec | null): void {
    if (filter !== null) {
      validateFilterSpec(filter);
      if (filter.range.endRow >= this.rowCount || filter.range.endCol >= this.columnCount) {
        throw new SheetError(
          "E_INVALID_RANGE",
          `FilterSpec range exceeds worksheet bounds (${this.rowCount} rows × ${this.columnCount} columns)`,
        );
      }
    }
    this._filter = filter === null ? null : cloneFilterSpec(filter);
  }

  /**
   * Read access. The returned object is the INTERNAL cell, typed Readonly:
   * mutating it bypasses the Command Bus (no event, no history, no dirty
   * region) and is a contract violation. Renderers should use asView().
   */
  getCell(row: number, col: number): Readonly<CellData> | undefined {
    return this.cells.get(row, col);
  }

  /** Read-only view for renderers and other consumers outside the command path. */
  asView(): WorksheetView {
    return this;
  }

  getRowHeight(row: number): number | undefined {
    return this.rowHeights.get(row);
  }

  getColumnWidth(col: number): number | undefined {
    return this.columnWidths.get(col);
  }

  setCell(row: number, col: number, data: CellData): void {
    this.cells.set(row, col, data);
  }

  deleteCell(row: number, col: number): boolean {
    return this.cells.delete(row, col);
  }

  clearCells(): void {
    this.cells.clear();
  }

  cellEntries(): IterableIterator<[number, number, CellData]> {
    return this.cells.entries();
  }

  forEachCellInRange(
    range: Range,
    callback: (row: number, col: number, data: CellData) => void,
  ): void {
    this.cells.forEachInRange(range, callback);
  }

  /**
   * Insert `count` empty rows before row `at`. Cells below shift down.
   * O(non-empty cells) — rebuilds the store; acceptable at M0 scale and
   * isolated behind this method for later optimization.
   */
  insertRows(at: number, count: number): void {
    if (count <= 0) return;
    const next = this.storeFactory.create();
    for (const [row, col, data] of this.cells.entries()) {
      next.set(row >= at ? row + count : row, col, data);
    }
    this.cells = next;
    shiftSizeMap(this.rowHeights, at, count);
    this.rowCount += count;
  }

  deleteRows(at: number, count: number): void {
    if (count <= 0) return;
    const next = this.storeFactory.create();
    for (const [row, col, data] of this.cells.entries()) {
      if (row >= at && row < at + count) continue; // dropped
      next.set(row >= at + count ? row - count : row, col, data);
    }
    this.cells = next;
    shiftSizeMap(this.rowHeights, at, -count);
    this.rowCount = Math.max(0, this.rowCount - count);
  }

  insertColumns(at: number, count: number): void {
    if (count <= 0) return;
    const next = this.storeFactory.create();
    for (const [row, col, data] of this.cells.entries()) {
      next.set(row, col >= at ? col + count : col, data);
    }
    this.cells = next;
    shiftSizeMap(this.columnWidths, at, count);
    this.columnCount += count;
  }

  deleteColumns(at: number, count: number): void {
    if (count <= 0) return;
    const next = this.storeFactory.create();
    for (const [row, col, data] of this.cells.entries()) {
      if (col >= at && col < at + count) continue;
      next.set(row, col >= at + count ? col - count : col, data);
    }
    this.cells = next;
    shiftSizeMap(this.columnWidths, at, -count);
    this.columnCount = Math.max(0, this.columnCount - count);
  }
}

function cloneFilterSpec(spec: FilterSpec): FilterSpec {
  return {
    range: { ...spec.range },
    hasHeader: spec.hasHeader,
    conditions: spec.conditions.map((condition) => ({ ...condition })),
  };
}

function shiftSizeMap(map: Map<number, number>, at: number, delta: number): void {
  const next = new Map<number, number>();
  for (const [index, size] of map) {
    if (delta > 0) {
      next.set(index >= at ? index + delta : index, size);
    } else {
      const removed = -delta;
      if (index >= at && index < at + removed) continue;
      next.set(index >= at + removed ? index - removed : index, size);
    }
  }
  map.clear();
  for (const [index, size] of next) map.set(index, size);
}
