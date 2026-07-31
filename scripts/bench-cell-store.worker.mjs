// CellStore benchmark worker. Run by bench-cell-store.mjs in a dedicated process.
// Usage: node --expose-gc --max-old-space-size=6144 bench-cell-store.worker.mjs <string-key|number-key|chunked>
//
// Scenario: 100,000 rows x 20 cols = 2,000,000 non-empty cells.

import {
  stringKeyCellStoreFactory,
  numberKeyCellStoreFactory,
  chunkedCellStoreFactory,
} from "../packages/core/dist/index.js";

const ROWS = 100_000;
const COLS = 20;
const RANDOM_READS = 1_000_000;

const factories = {
  "string-key": stringKeyCellStoreFactory,
  "number-key": numberKeyCellStoreFactory,
  chunked: chunkedCellStoreFactory,
};

const impl = process.argv[2];
const factory = factories[impl];
if (factory === undefined) {
  console.error(`Unknown impl "${impl}". Expected one of: ${Object.keys(factories).join(", ")}`);
  process.exit(2);
}

const now = () => performance.now();
const mem = () => {
  global.gc();
  const m = process.memoryUsage();
  return { rssMB: +(m.rss / 1_048_576).toFixed(1), heapUsedMB: +(m.heapUsed / 1_048_576).toFixed(1) };
};

// Deterministic pseudo-random generator (same stream for every impl).
let seed = 42;
const rand = () => {
  seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
  return seed / 2_147_483_648;
};

const store = factory.create();

// --- write ---
let t0 = now();
for (let row = 0; row < ROWS; row++) {
  for (let col = 0; col < COLS; col++) {
    store.set(row, col, { value: row * COLS + col });
  }
}
const writeMs = +(now() - t0).toFixed(1);
const afterWrite = mem();

// --- random read ---
const reads = [];
for (let i = 0; i < RANDOM_READS; i++) {
  reads.push([Math.floor(rand() * ROWS), Math.floor(rand() * COLS)]);
}
t0 = now();
let hits = 0;
for (const [row, col] of reads) {
  if (store.get(row, col) !== undefined) hits++;
}
const readMs = +(now() - t0).toFixed(1);

// --- range iteration (full used range) ---
t0 = now();
let visited = 0;
store.forEachInRange({ startRow: 0, startCol: 0, endRow: ROWS - 1, endCol: COLS - 1 }, () => {
  visited++;
});
const rangeMs = +(now() - t0).toFixed(1);

// --- snapshot-like serialization (entries -> plain record) ---
t0 = now();
const record = {};
for (const [row, col, data] of store.entries()) {
  record[`${row}:${col}`] = data;
}
const serializeMs = +(now() - t0).toFixed(1);
const keys = Object.keys(record).length;

const result = {
  impl,
  cells: store.size,
  writeMs,
  randomReadMs: readMs,
  rangeIterateMs: rangeMs,
  serializeMs,
  ...afterWrite,
  sanity: { hits, visited, keys },
  node: process.version,
};
process.stdout.write(JSON.stringify(result));
