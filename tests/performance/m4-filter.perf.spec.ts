// Reproducible M4.2 filter benchmark.
// Run: pnpm bench:filter  → writes test-results/filter-perf.json (or $BENCH_OUT).
// This stays outside normal E2E because constructing the 100k × 20 dataset is
// intentionally expensive and its measurements are environment-dependent.

import { expect, test } from "@playwright/test";
import { writeFileSync } from "node:fs";

interface FrameStats {
  paintMs: number;
  full: boolean;
  paintedCells: number;
}

const ROWS = 100_000;
const WARMUP_FRAMES = 10;
const MEASURED_FRAMES = 120;

async function apply(page: import("@playwright/test").Page, operations: object[]): Promise<void> {
  await page.evaluate(async (ops) => {
    const w = window as unknown as {
      __api: { applyOperations(value: object): Promise<unknown> };
      __workbookId: string;
      __sheetId: string;
    };
    await w.__api.applyOperations({ workbookId: w.__workbookId, sheetId: w.__sheetId, atomic: true, operations: ops });
  }, operations);
}

async function visualRows(page: import("@playwright/test").Page): Promise<number> {
  return page.evaluate(() => (window as unknown as { __grid: { getRowProjection(): { visualRowCount: number } } }).__grid.getRowProjection().visualRowCount);
}

test("filter perf: 100k rows × 20 columns", async ({ page }) => {
  const pageErrors: string[] = [];
  page.on("pageerror", (error) => pageErrors.push(String(error)));
  await page.goto("/");
  await page.waitForSelector("[data-testid=sheet-grid] canvas");

  const loadStarted = performance.now();
  await page.getByRole("button", { name: "Load 100k×20" }).click();
  await page.waitForSelector("text=/Loaded 100,000×20/", { timeout: 300_000 });
  const loadMs = performance.now() - loadStarted;

  const spec = {
    range: { startRow: 0, startCol: 0, endRow: ROWS - 1, endCol: 19 },
    hasHeader: false,
    conditions: [{ columnOffset: 0, operator: "contains", value: "row-" }],
  };
  const started = await page.evaluate(() => performance.now());
  await apply(page, [{ type: "filter.apply", spec }]);
  await expect.poll(() => visualRows(page), { timeout: 30_000 }).toBe(ROWS);
  const applyMs = (await page.evaluate(() => performance.now())) - started;

  const clearStarted = await page.evaluate(() => performance.now());
  await apply(page, [{ type: "filter.clear" }]);
  await expect.poll(() => visualRows(page), { timeout: 30_000 }).toBe(ROWS);
  const clearMs = (await page.evaluate(() => performance.now())) - clearStarted;

  // Measure steady-state canvas painting separately from full redraws caused
  // by dataset load and projection replacement. The demo persists committed
  // workbook snapshots with a 400ms debounce; allow that unrelated storage
  // work to settle before starting the renderer sample window.
  await page.waitForTimeout(700);
  await page.evaluate(() => {
    (window as unknown as { __frameStats: FrameStats[] }).__frameStats = [];
  });
  const grid = page.locator("[data-testid=sheet-grid]");
  for (let i = 0; i < WARMUP_FRAMES + MEASURED_FRAMES; i += 1) {
    await grid.dispatchEvent("wheel", { deltaY: 200 });
    await page.waitForTimeout(40);
  }
  const frameStats = await page.evaluate(() => (window as unknown as { __frameStats?: FrameStats[] }).__frameStats ?? []);
  const measured = frameStats.slice(WARMUP_FRAMES);
  expect(measured).toHaveLength(MEASURED_FRAMES);
  const paints = measured.map((frame) => frame.paintMs).sort((a, b) => a - b);
  const p95 = paints.length === 0 ? 0 : paints[Math.min(paints.length - 1, Math.floor(paints.length * 0.95))]!;
  const heapUsedMB = await page.evaluate(() => performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : null);
  const result = {
    date: new Date().toISOString(),
    dataset: "100,000 rows x 20 cols = 2,000,000 cells",
    loadMs: Number(loadMs.toFixed(2)),
    filterApplyMs: Number(applyMs.toFixed(2)),
    filterClearMs: Number(clearMs.toFixed(2)),
    gridPaintP50Ms: Number((paints[Math.floor(paints.length * 0.5)] ?? 0).toFixed(2)),
    gridPaintP95Ms: Number(p95.toFixed(2)),
    paintSamples: paints.length,
    measuredSamples: measured,
    heapUsedMB,
    pageErrors,
  };
  writeFileSync(process.env.BENCH_OUT ?? "test-results/filter-perf.json", JSON.stringify(result, null, 2));

  expect(result.filterApplyMs).toBeLessThanOrEqual(500);
  expect(result.filterClearMs).toBeLessThanOrEqual(100);
  expect(result.gridPaintP95Ms).toBeLessThanOrEqual(25);
  expect(pageErrors).toEqual([]);
});
