import { expect, test, type Page } from "@playwright/test";

/**
 * M4.1.1 grid stability under projections (no filter UI yet — projections
 * are installed directly through the demo's E2E probe hooks).
 *
 * Demo sheet: 1000 rows × 26 cols, default row height 26.
 */

interface ProjectionSpec {
  rangeStart: number;
  rangeEnd: number;
  visible: number[];
}

async function installProjection(page: Page, spec: ProjectionSpec | null): Promise<void> {
  await page.evaluate((s) => {
    const w = window as unknown as {
      __grid: { setRowProjection(p: unknown): void };
      __FilteredRowProjection: new (
        rows: number,
        range: { startRow: number; endRow: number },
        visible: number[],
      ) => unknown;
    };
    w.__grid.setRowProjection(
      s === null ? null : new w.__FilteredRowProjection(1000, { startRow: s.rangeStart, endRow: s.rangeEnd }, s.visible),
    );
  }, spec);
}

async function gridProbe(page: Page): Promise<{ scrollY: number; maxScroll: number }> {
  return page.evaluate(() => {
    const g = (window as unknown as { __grid: { scrollY: number; scrollbarGeometry(): { vertical: { maxScroll: number } | null } } }).__grid;
    return { scrollY: g.scrollY, maxScroll: g.scrollbarGeometry().vertical?.maxScroll ?? 0 };
  });
}

test.describe("M4.1.1 projection stability", () => {
  test("deep scroll + drastically shrinking projection clamps scrollY and keeps content visible", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));

    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();

    // Deep scroll: Ctrl+End lands at the true max scroll (1000 rows).
    await grid.click();
    await page.keyboard.press("Control+End");
    await expect(page.getByText(/Active: Z1000/)).toBeVisible();
    const before = await gridProbe(page);
    expect(before.scrollY).toBeGreaterThan(20_000);

    // Hide rows 10..999 except three → visual axis shrinks to 13 rows.
    await installProjection(page, { rangeStart: 10, rangeEnd: 999, visible: [10, 11, 12] });
    await page.waitForTimeout(100);

    // scrollY is re-clamped against the shrunken axis (13 rows * 26px = 338,
    // far below the viewport) — the canvas can never strand on empty space.
    const after = await gridProbe(page);
    expect(after.scrollY).toBeLessThanOrEqual(after.maxScroll);

    // The active cell (Z1000, hidden) relocated to the nearest visible
    // physical row above (row 12) and the status bar reflects it.
    await expect(page.getByText(/Active: Z13/)).toBeVisible();
    expect(errors).toEqual([]);
  });

  test("projection with a foreign row count is rejected and the grid keeps working", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));

    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();

    const message = await page.evaluate(() => {
      const w = window as unknown as {
        __grid: { setRowProjection(p: unknown): void };
        __FilteredRowProjection: new (
          rows: number,
          range: { startRow: number; endRow: number },
          visible: number[],
        ) => unknown;
      };
      try {
        // Built for a 999-row sheet — one row short of the demo's 1000.
        w.__grid.setRowProjection(new w.__FilteredRowProjection(999, { startRow: 0, endRow: 10 }, [0, 1]));
        return null;
      } catch (error) {
        return String(error);
      }
    });
    expect(message).toMatch(/does not match/);

    // Grid is untouched: click still selects.
    const box = await grid.boundingBox();
    await page.mouse.click(box!.x + 48 + 150, box!.y + 26 + 65);
    await expect(page.getByText(/Active: B3/)).toBeVisible();
    expect(errors).toEqual([]);
  });

  test("stale projection cannot be reinstalled after a structural change", async ({ page }) => {
    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();

    // Install a valid projection, then insert a row (rowCount 1000 → 1001).
    await installProjection(page, { rangeStart: 5, rangeEnd: 14, visible: [5, 7, 9] });
    await page.evaluate(async () => {
      const w = window as unknown as {
        __api: { applyOperations(o: object): Promise<unknown> };
        __workbookId: string;
        __sheetId: string;
      };
      await w.__api.applyOperations({
        workbookId: w.__workbookId,
        sheetId: w.__sheetId,
        atomic: true,
        operations: [{ type: "row.insert", at: 0, count: 1 }],
      });
    });
    await page.waitForTimeout(80);

    // Reinstalling the projection built for 1000 rows must now fail.
    const message = await page.evaluate(() => {
      const w = window as unknown as {
        __grid: { setRowProjection(p: unknown): void };
        __FilteredRowProjection: new (
          rows: number,
          range: { startRow: number; endRow: number },
          visible: number[],
        ) => unknown;
      };
      try {
        w.__grid.setRowProjection(new w.__FilteredRowProjection(1000, { startRow: 5, endRow: 14 }, [5, 7, 9]));
        return null;
      } catch (error) {
        return String(error);
      }
    });
    expect(message).toMatch(/does not match/);
  });

  test("hidden active cell relocates AND scrolls into the viewport; F2 opens over it", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));

    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();

    // Active A61 (row 60). Filter rows 50..199 keeping only row 150 → row 60
    // is hidden; relocation target physical 150 maps to VISUAL row 50
    // (50 identity rows + index 0) ≈ 1300px below the fold: must auto-scroll.
    const startBox = await grid.boundingBox();
    await page.mouse.click(startBox!.x + 48 + 10, startBox!.y + 26 + 10);
    await expect(page.getByText(/Active: A1/)).toBeVisible();
    for (let i = 0; i < 60; i++) await page.keyboard.press("ArrowDown");
    await expect(page.getByText(/Active: A61/)).toBeVisible();
    await installProjection(page, { rangeStart: 50, rangeEnd: 199, visible: [150] });
    await page.waitForTimeout(100);

    await expect(page.getByText(/Active: A151/)).toBeVisible();
    const probe = await gridProbe(page);
    expect(probe.scrollY).toBeGreaterThan(0); // scrolled to reveal row 51

    // F2 opens the editor INSIDE the viewport (regression: frozen detection
    // via main.rowStart could crash or misposition the overlay here).
    await page.keyboard.press("F2");
    const editor = page.locator("[data-testid=sheet-grid] textarea");
    await expect(editor).toBeVisible();
    const editorBox = await editor.boundingBox();
    const gridBox = await grid.boundingBox();
    expect(editorBox!.y).toBeGreaterThanOrEqual(gridBox!.y);
    expect(editorBox!.y + editorBox!.height).toBeLessThanOrEqual(gridBox!.y + gridBox!.height + 1);
    await page.keyboard.press("Escape");
    expect(errors).toEqual([]);
  });

  test("zero visible rows: canvas empties safely, input is inert, clear restores", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));

    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();

    // Hide EVERY row: visualRowCount = 0.
    await installProjection(page, { rangeStart: 0, rangeEnd: 999, visible: [] });
    await page.waitForTimeout(100);

    // Keyboard navigation and Ctrl+End are no-ops — no -1 coordinates, no throws.
    await grid.click();
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Control+End");
    await page.keyboard.press("F2");
    await page.waitForTimeout(80);
    expect(errors).toEqual([]);

    // Clearing restores the identity projection and full interactivity.
    await installProjection(page, null);
    await page.waitForTimeout(80);
    const box = await grid.boundingBox();
    await page.mouse.click(box!.x + 48 + 150, box!.y + 26 + 65);
    await expect(page.getByText(/Active: B3/)).toBeVisible();
    expect(errors).toEqual([]);
  });
});
