// `pnpm bench:csv` — run the reproducible M5 CSV Worker benchmark.

import { execFileSync } from "node:child_process";
import { readFileSync, rmSync } from "node:fs";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const out = `${root}test-results/csv-perf.json`;
rmSync(out, { force: true });

execFileSync(
  "pnpm",
  [
    "exec",
    "playwright",
    "test",
    "-c",
    "playwright.perf.config.ts",
    "m5-csv.perf.spec.ts",
  ],
  { cwd: root, stdio: "inherit", env: { ...process.env, BENCH_OUT: out } },
);

const result = JSON.parse(readFileSync(out, "utf8"));
console.log("\n=== M5 CSV perf summary ===");
console.log(`Worker import : ${result.importMs} ms`);
console.log(`CSV export    : ${result.exportMs} ms`);
console.log(`CSV bytes     : ${result.csvBytes}`);
console.log(`raw JSON      : ${out}`);
