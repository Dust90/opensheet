// CellStore benchmark runner. NOT part of CI — run manually before freezing
// the storage design (ADR-0005). Each candidate runs in its own process so
// GC state never leaks between measurements.
//
// Usage: pnpm build && pnpm bench:cell-store

import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const worker = `${root}scripts/bench-cell-store.worker.mjs`;

if (!existsSync(`${root}packages/core/dist/index.js`)) {
  console.error("packages/core/dist not found. Run `pnpm build` first.");
  process.exit(2);
}

const impls = ["string-key", "number-key", "chunked"];
const results = [];

for (const impl of impls) {
  process.stderr.write(`[bench] running ${impl} in a dedicated process...\n`);
  const stdout = execFileSync(
    process.execPath,
    ["--expose-gc", "--max-old-space-size=6144", worker, impl],
    { cwd: root, maxBuffer: 64 * 1024 * 1024 },
  );
  const result = JSON.parse(stdout.toString());
  if (result.sanity.hits !== 1_000_000 || result.sanity.visited !== result.cells) {
    process.stderr.write(`[bench] SANITY FAILURE for ${impl}: ${JSON.stringify(result.sanity)}\n`);
    process.exit(1);
  }
  results.push(result);
}

const env = {
  date: new Date().toISOString(),
  node: process.version,
  platform: `${process.platform} ${process.arch}`,
  scenario: "100,000 rows x 20 cols = 2,000,000 non-empty cells",
};

console.log(`\nCellStore benchmark — ${env.scenario}`);
console.log(`${env.date} | node ${env.node} | ${env.platform}\n`);
console.log(
  "| impl | write ms | 1M random reads ms | full-range iterate ms | serialize ms | RSS MB | heapUsed MB |",
);
console.log("|---|---|---|---|---|---|---|");
for (const r of results) {
  console.log(
    `| ${r.impl} | ${r.writeMs} | ${r.randomReadMs} | ${r.rangeIterateMs} | ${r.serializeMs} | ${r.rssMB} | ${r.heapUsedMB} |`,
  );
}

mkdirSync(`${root}test-results`, { recursive: true });
writeFileSync(
  `${root}test-results/cell-store-benchmark.json`,
  JSON.stringify({ env, results }, null, 2),
);
console.log("\nRaw results written to test-results/cell-store-benchmark.json");
