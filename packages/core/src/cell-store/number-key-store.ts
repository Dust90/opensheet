// Candidate B: numeric key (row * KEY_STRIDE + col) in a single Map.

import type { CellData, CellStore, CellStoreFactory, Range } from "@opensheet/shared";
import { SheetError } from "@opensheet/shared";

/**
 * Stride between rows in the numeric key space. 2^20 = 1,048,576 columns of
 * headroom keeps keys well below Number.MAX_SAFE_INTEGER for 1M+ rows
 * (max key ~= 1.1e12).
 */
export const KEY_STRIDE = 1_048_576;

export class NumberKeyCellStore implements CellStore {
  private readonly map = new Map<number, CellData>();

  private key(row: number, col: number): number {
    if (col >= KEY_STRIDE) {
      throw new SheetError("E_INVALID_ADDRESS", `Column index exceeds numeric key space: ${col}`);
    }
    return row * KEY_STRIDE + col;
  }

  get(row: number, col: number): CellData | undefined {
    return this.map.get(row * KEY_STRIDE + col);
  }

  set(row: number, col: number, data: CellData): void {
    this.map.set(this.key(row, col), data);
  }

  delete(row: number, col: number): boolean {
    return this.map.delete(row * KEY_STRIDE + col);
  }

  has(row: number, col: number): boolean {
    return this.map.has(row * KEY_STRIDE + col);
  }

  get size(): number {
    return this.map.size;
  }

  clear(): void {
    this.map.clear();
  }

  *entries(): IterableIterator<[number, number, CellData]> {
    for (const [key, data] of this.map) {
      yield [Math.floor(key / KEY_STRIDE), key % KEY_STRIDE, data];
    }
  }

  forEachInRange(range: Range, callback: (row: number, col: number, data: CellData) => void): void {
    for (let row = range.startRow; row <= range.endRow; row++) {
      const rowBase = row * KEY_STRIDE;
      for (let col = range.startCol; col <= range.endCol; col++) {
        const data = this.map.get(rowBase + col);
        if (data !== undefined) callback(row, col, data);
      }
    }
  }
}

export const numberKeyCellStoreFactory: CellStoreFactory = {
  name: "number-key",
  create: () => new NumberKeyCellStore(),
};
