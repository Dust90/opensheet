// `pnpm bench:grid` — run the reproducible grid perf benchmark and print the
// summary. Requires a working Playwright Chromium install
// (`pnpm exec playwright install chromium`).

import { execFileSync } from "node:child_process";
import { readFileSync, rmSync } from "node:fs";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const out = `${root}test-results/grid-perf.json`;
try {
  rmSync(out, { force: true });
} catch {
  /* ignore */
}

execFileSync(
  "pnpm",
  ["exec", "playwright", "test", "-c", "playwright.perf.config.ts"],
  { cwd: root, stdio: "inherit", env: { ...process.env, BENCH_OUT: out } },
);

const result = JSON.parse(readFileSync(out, "utf8"));
console.log("\n=== Grid perf summary ===");
console.log(`dataset     : ${result.dataset}`);
console.log(`load 2M     : ${result.loadMs} ms`);
console.log(`heapUsed    : ${result.heapUsedMB} MB`);
console.log(`viewport    : ${JSON.stringify(result.viewport)}`);
console.log(`deviceScale : ${result.deviceScaleFactor}`);
console.log(
  `paint ms    : p50=${result.summary.p50} p90=${result.summary.p90} p95=${result.summary.p95} max=${result.summary.max}`,
);
console.log(
  `cells/frame : ${result.summary.paintedCellsMin}-${result.summary.paintedCellsMax}`,
);
console.log(`warmup      : ${result.warmupSamples.length} samples (excluded)`);
console.log(`measured    : ${result.measuredSamples.length} samples`);
console.log(`raw JSON    : ${out}`);
