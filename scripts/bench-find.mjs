// `pnpm bench:find` — run the reproducible M4.4 find benchmark.

import { execFileSync } from "node:child_process";
import { readFileSync, rmSync } from "node:fs";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const out = `${root}test-results/find-perf.json`;
rmSync(out, { force: true });

execFileSync(
  "pnpm",
  ["exec", "playwright", "test", "-c", "playwright.perf.config.ts", "m4-find.perf.spec.ts"],
  { cwd: root, stdio: "inherit", env: { ...process.env, BENCH_OUT: out } },
);

const result = JSON.parse(readFileSync(out, "utf8"));
console.log("\n=== M4.4 find perf summary ===");
console.log(`dense findCells : ${result.denseFindCellsMs} ms`);
console.log(`dense findNext  : ${result.denseFindNextMs} ms`);
console.log(`sparse findCells: ${result.sparseFindCellsMs} ms`);
console.log(`raw JSON        : ${out}`);
