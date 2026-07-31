// Reproducible grid performance benchmark.
// Run: pnpm bench:grid  →  writes test-results/grid-perf.json (or $BENCH_OUT).
// Not part of `pnpm test:e2e`; raw per-frame samples are kept so results can
// be re-derived and compared across machines.

import { expect, test } from "@playwright/test";
import { writeFileSync } from "node:fs";

interface FrameStats {
  paintMs: number;
  full: boolean;
  paintedCells: number;
}

const WARMUP_FRAMES = 10;
const MEASURED_FRAMES = 120;

test("grid perf: 2M-cell sheet scroll paint", async ({ page }) => {
  const pageErrors: string[] = [];
  page.on("pageerror", (error) => pageErrors.push(String(error)));

  await page.goto("/");
  await page.waitForSelector("[data-testid=sheet-grid] canvas");
  const viewport = page.viewportSize();

  const t0 = Date.now();
  await page.getByRole("button", { name: "Load 100k×20" }).click();
  await page.waitForSelector("text=/Loaded 100,000×20/", { timeout: 300_000 });
  const loadMs = Date.now() - t0;

  await page.evaluate(() => {
    (window as unknown as { __frameStats: FrameStats[] }).__frameStats = [];
  });

  for (let i = 0; i < WARMUP_FRAMES + MEASURED_FRAMES; i++) {
    await page.locator("[data-testid=sheet-grid]").dispatchEvent("wheel", { deltaY: 200 });
    await page.waitForTimeout(40);
  }

  const all = await page.evaluate(
    () => (window as unknown as { __frameStats: FrameStats[] }).__frameStats,
  );
  const warmup = all.slice(0, WARMUP_FRAMES);
  const measured = all.slice(WARMUP_FRAMES);
  // Guard against browser frame coalescing silently shrinking the sample set.
  expect(warmup).toHaveLength(WARMUP_FRAMES);
  expect(measured).toHaveLength(MEASURED_FRAMES);
  const paints = measured.map((f) => f.paintMs).sort((a, b) => a - b);
  const quantile = (q: number) =>
    paints[Math.min(paints.length - 1, Math.floor(paints.length * q))] ?? 0;

  const heap = await page.evaluate(() =>
    performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : null,
  );
  const deviceScaleFactor = await page.evaluate(() => devicePixelRatio);

  const summary = {
    samples: measured.length,
    p50: Number(quantile(0.5).toFixed(2)),
    p90: Number(quantile(0.9).toFixed(2)),
    p95: Number(quantile(0.95).toFixed(2)),
    max: Number((paints[paints.length - 1] ?? 0).toFixed(2)),
    paintedCellsMin: Math.min(...measured.map((f) => f.paintedCells)),
    paintedCellsMax: Math.max(...measured.map((f) => f.paintedCells)),
  };

  const out = {
    date: new Date().toISOString(),
    browser: "chromium (Playwright headless)",
    browserArgs: [],
    deviceScaleFactor,
    viewport,
    dataset: "100,000 rows x 20 cols = 2,000,000 cells",
    loadMs,
    heapUsedMB: heap,
    warmupSamples: warmup,
    measuredSamples: measured,
    summary,
    pageErrors,
  };

  const outPath = process.env.BENCH_OUT ?? "test-results/grid-perf.json";
  writeFileSync(outPath, JSON.stringify(out, null, 2));

  // Verdict in the JSON; the assertion here is a loose regression gate only
  // (headless software rasterization varies by runner).
  expect(summary.p95).toBeLessThan(50);
  expect(pageErrors).toEqual([]);
});
