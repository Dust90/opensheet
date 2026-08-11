// `pnpm bench:formula` — run the reproducible M3 lazy-range formula benchmark.

import { execFileSync } from "node:child_process";
import { readFileSync, rmSync } from "node:fs";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const out = `${root}test-results/formula-perf.json`;
rmSync(out, { force: true });

execFileSync(
  "pnpm",
  [
    "exec",
    "playwright",
    "test",
    "-c",
    "playwright.perf.config.ts",
    "m3-formula.perf.spec.ts",
  ],
  { cwd: root, stdio: "inherit", env: { ...process.env, BENCH_OUT: out } },
);

const result = JSON.parse(readFileSync(out, "utf8"));
console.log("\n=== M3 formula perf summary ===");
console.log(`lazy SUM 900k : ${result.formulaMs} ms`);
console.log(`result        : ${result.value}`);
console.log(`raw JSON      : ${out}`);
