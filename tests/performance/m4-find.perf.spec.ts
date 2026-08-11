import { expect, test } from "@playwright/test";
import { writeFileSync } from "node:fs";

const rows = 100_000;
const options = {
  query: "row-100000",
  matchCase: true,
  wholeCell: true,
  searchIn: "values" as const,
  scope: "all" as const,
  direction: "forward" as const,
};

test("find perf: dense and sparse 100k × 20 worksheets", async ({ page }) => {
  await page.goto("/");
  await page.getByRole("button", { name: "Load 100k×20" }).click();
  await page.waitForSelector("text=/Loaded 100,000×20/", { timeout: 300_000 });

  const dense = await page.evaluate((findOptions) => {
    const w = window as unknown as {
      __api: {
        findCells(value: object): { row: number; col: number }[];
        findNext(value: object): { row: number; col: number } | null;
      };
      __sheetId: string;
    };
    const start = performance.now();
    const matches = w.__api.findCells({ sheetId: w.__sheetId, ...findOptions });
    const findCellsMs = performance.now() - start;
    const nextStart = performance.now();
    const next = w.__api.findNext({ sheetId: w.__sheetId, ...findOptions, from: { row: 0, col: 0 } });
    const findNextMs = performance.now() - nextStart;
    return { findCellsMs, findNextMs, matches, next };
  }, options);
  expect(dense.matches).toEqual([{ row: rows - 1, col: 0 }]);
  expect(dense.next).toEqual({ row: rows - 1, col: 0 });

  const sparse = await page.evaluate(async () => {
    const w = window as unknown as {
      __api: {
        createSheet(value: { name: string; rows: number; columns: number }): { id: string };
        applyOperations(value: object): Promise<unknown>;
        findCells(value: object): { row: number; col: number }[];
      };
      __workbookId: string;
    };
    const sheet = w.__api.createSheet({ name: "Sparse Find", rows: 100_000, columns: 20 });
    await w.__api.applyOperations({
      workbookId: w.__workbookId,
      sheetId: sheet.id,
      atomic: true,
      operations: [{ type: "cell.set", range: "T100000", value: "sparse-tail" }],
    });
    const start = performance.now();
    const matches = w.__api.findCells({
      sheetId: sheet.id,
      query: "sparse-tail",
      matchCase: true,
      wholeCell: true,
      searchIn: "values",
      scope: "all",
      direction: "forward",
    });
    return { findCellsMs: performance.now() - start, matches };
  });
  expect(sparse.matches).toEqual([{ row: rows - 1, col: 19 }]);

  const result = {
    denseFindCellsMs: dense.findCellsMs,
    denseFindNextMs: dense.findNextMs,
    sparseFindCellsMs: sparse.findCellsMs,
  };
  console.log(JSON.stringify(result));
  if (process.env.BENCH_OUT !== undefined) writeFileSync(process.env.BENCH_OUT, JSON.stringify(result, null, 2));
  expect(result.denseFindCellsMs).toBeLessThanOrEqual(500);
  expect(result.denseFindNextMs).toBeLessThanOrEqual(500);
  expect(result.sparseFindCellsMs).toBeLessThanOrEqual(500);
});
