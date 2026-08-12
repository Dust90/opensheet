// @injoysai/opensheet-shared — sparse cell storage contract

import type { CellData } from "./cell.js";
import type { Range } from "./range.js";

/**
 * Sparse storage for cell data. Empty cells MUST NOT have entries.
 *
 * The keying strategy is an implementation detail behind this interface;
 * candidates (string key / numeric key / chunked) are benchmarked and the
 * default is frozen by ADR-0005.
 */
export interface CellStore {
  get(row: number, col: number): CellData | undefined;
  set(row: number, col: number, data: CellData): void;
  delete(row: number, col: number): boolean;
  has(row: number, col: number): boolean;
  readonly size: number;
  clear(): void;
  /** Iterate all non-empty cells. Order is implementation-defined. */
  entries(): IterableIterator<[row: number, col: number, data: CellData]>;
  /** Visit only non-empty cells inside the given range. */
  forEachInRange(range: Range, callback: (row: number, col: number, data: CellData) => void): void;
}

export interface CellStoreFactory {
  readonly name: string;
  create(): CellStore;
}
