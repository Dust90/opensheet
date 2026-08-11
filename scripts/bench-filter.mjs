// `pnpm bench:filter` — run the reproducible M4.2 filter benchmark.

import { execFileSync } from "node:child_process";
import { readFileSync, rmSync } from "node:fs";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const out = `${root}test-results/filter-perf.json`;
rmSync(out, { force: true });

execFileSync(
  "pnpm",
  ["exec", "playwright", "test", "-c", "playwright.perf.config.ts", "m4-filter.perf.spec.ts"],
  { cwd: root, stdio: "inherit", env: { ...process.env, BENCH_OUT: out } },
);

const result = JSON.parse(readFileSync(out, "utf8"));
console.log("\n=== M4.2 filter perf summary ===");
console.log(`dataset      : ${result.dataset}`);
console.log(`load 2M      : ${result.loadMs} ms`);
console.log(`filter apply : ${result.filterApplyMs} ms`);
console.log(`filter clear : ${result.filterClearMs} ms`);
console.log(`paint ms     : p50=${result.gridPaintP50Ms} p95=${result.gridPaintP95Ms} (${result.paintSamples} samples)`);
console.log(`heap used    : ${result.heapUsedMB} MB`);
console.log(`raw JSON     : ${out}`);
