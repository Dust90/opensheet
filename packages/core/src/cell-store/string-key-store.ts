// Candidate A: string key "row:col" in a single Map.

import type { CellData, CellStore, CellStoreFactory, Range } from "@opensheet/shared";

export class StringKeyCellStore implements CellStore {
  private readonly map = new Map<string, CellData>();

  private key(row: number, col: number): string {
    return `${row}:${col}`;
  }

  get(row: number, col: number): CellData | undefined {
    return this.map.get(this.key(row, col));
  }

  set(row: number, col: number, data: CellData): void {
    this.map.set(this.key(row, col), data);
  }

  delete(row: number, col: number): boolean {
    return this.map.delete(this.key(row, col));
  }

  has(row: number, col: number): boolean {
    return this.map.has(this.key(row, col));
  }

  get size(): number {
    return this.map.size;
  }

  clear(): void {
    this.map.clear();
  }

  *entries(): IterableIterator<[number, number, CellData]> {
    for (const [key, data] of this.map) {
      const sep = key.indexOf(":");
      const row = Number(key.slice(0, sep));
      const col = Number(key.slice(sep + 1));
      yield [row, col, data];
    }
  }

  forEachInRange(range: Range, callback: (row: number, col: number, data: CellData) => void): void {
    for (const [key, data] of this.map) {
      const sep = key.indexOf(":");
      const row = Number(key.slice(0, sep));
      if (row < range.startRow || row > range.endRow) continue;
      const col = Number(key.slice(sep + 1));
      if (col < range.startCol || col > range.endCol) continue;
      callback(row, col, data);
    }
  }
}

export const stringKeyCellStoreFactory: CellStoreFactory = {
  name: "string-key",
  create: () => new StringKeyCellStore(),
};
