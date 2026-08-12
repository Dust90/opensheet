// Candidate C: chunked storage — Map<chunkKey, Map<innerKey, CellData>>.

import type { CellData, CellStore, CellStoreFactory, Range } from "@injoysai/opensheet-shared";

/** 128 rows × 128 cols per chunk. */
const CHUNK_SHIFT = 7;
const CHUNK_SIZE = 1 << CHUNK_SHIFT; // 128
const CHUNK_MASK = CHUNK_SIZE - 1;

export class ChunkedCellStore implements CellStore {
  private readonly chunks = new Map<number, Map<number, CellData>>();
  private cellCount = 0;

  private chunkKey(row: number, col: number): number {
    return (row >> CHUNK_SHIFT) * 1_048_576 + (col >> CHUNK_SHIFT);
  }

  private innerKey(row: number, col: number): number {
    return (row & CHUNK_MASK) * CHUNK_SIZE + (col & CHUNK_MASK);
  }

  get(row: number, col: number): CellData | undefined {
    return this.chunks.get(this.chunkKey(row, col))?.get(this.innerKey(row, col));
  }

  set(row: number, col: number, data: CellData): void {
    const ck = this.chunkKey(row, col);
    let chunk = this.chunks.get(ck);
    if (chunk === undefined) {
      chunk = new Map();
      this.chunks.set(ck, chunk);
    }
    const ik = this.innerKey(row, col);
    if (!chunk.has(ik)) this.cellCount++;
    chunk.set(ik, data);
  }

  delete(row: number, col: number): boolean {
    const ck = this.chunkKey(row, col);
    const chunk = this.chunks.get(ck);
    if (chunk === undefined) return false;
    const removed = chunk.delete(this.innerKey(row, col));
    if (removed) {
      this.cellCount--;
      if (chunk.size === 0) this.chunks.delete(ck);
    }
    return removed;
  }

  has(row: number, col: number): boolean {
    return this.chunks.get(this.chunkKey(row, col))?.has(this.innerKey(row, col)) ?? false;
  }

  get size(): number {
    return this.cellCount;
  }

  clear(): void {
    this.chunks.clear();
    this.cellCount = 0;
  }

  *entries(): IterableIterator<[number, number, CellData]> {
    for (const [ck, chunk] of this.chunks) {
      const chunkRow = Math.floor(ck / 1_048_576) << CHUNK_SHIFT;
      const chunkCol = (ck % 1_048_576) << CHUNK_SHIFT;
      for (const [ik, data] of chunk) {
        yield [
          chunkRow + Math.floor(ik / CHUNK_SIZE),
          chunkCol + (ik % CHUNK_SIZE),
          data,
        ];
      }
    }
  }

  forEachInRange(range: Range, callback: (row: number, col: number, data: CellData) => void): void {
    const startChunkRow = range.startRow >> CHUNK_SHIFT;
    const endChunkRow = range.endRow >> CHUNK_SHIFT;
    const startChunkCol = range.startCol >> CHUNK_SHIFT;
    const endChunkCol = range.endCol >> CHUNK_SHIFT;
    for (let cr = startChunkRow; cr <= endChunkRow; cr++) {
      for (let cc = startChunkCol; cc <= endChunkCol; cc++) {
        const chunk = this.chunks.get(cr * 1_048_576 + cc);
        if (chunk === undefined) continue;
        const baseRow = cr << CHUNK_SHIFT;
        const baseCol = cc << CHUNK_SHIFT;
        for (const [ik, data] of chunk) {
          const row = baseRow + Math.floor(ik / CHUNK_SIZE);
          if (row < range.startRow || row > range.endRow) continue;
          const col = baseCol + (ik % CHUNK_SIZE);
          if (col < range.startCol || col > range.endCol) continue;
          callback(row, col, data);
        }
      }
    }
  }
}

export const chunkedCellStoreFactory: CellStoreFactory = {
  name: "chunked",
  create: () => new ChunkedCellStore(),
};
