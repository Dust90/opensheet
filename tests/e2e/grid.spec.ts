import { expect, test } from "@playwright/test";

test.describe("M1 grid", () => {
  test("renders canvases, selects cells via mouse and keyboard", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));

    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();
    // Dual canvases mounted.
    await expect(grid.locator("canvas")).toHaveCount(2);

    // Click a cell → status bar shows its address (click ~ B3 area).
    const box = await grid.boundingBox();
    if (box === null) throw new Error("no grid box");
    await page.mouse.click(box.x + 48 + 150, box.y + 26 + 65); // col 1, row 2
    await expect(page.getByText(/Active: B3/)).toBeVisible();

    // Keyboard navigation updates the active cell (container already focused
    // by the mousedown above).
    await page.keyboard.press("ArrowDown");
    await expect(page.getByText(/Active: B4/)).toBeVisible();
    await page.keyboard.press("ArrowRight");
    await expect(page.getByText(/Active: C4/)).toBeVisible();
    await page.keyboard.press("Tab");
    await expect(page.getByText(/Active: D4/)).toBeVisible();
    await page.keyboard.press("Enter");
    await expect(page.getByText(/Active: D5/)).toBeVisible();

    // Write sample data through the runtime and confirm no render errors.
    await page.getByRole("button", { name: "Write sample" }).click();
    await page.waitForTimeout(200);

    // Freeze panes toggle works via the command path.
    await page.getByRole("button", { name: "Freeze row 1 + col A" }).click();
    await expect(page.getByRole("button", { name: "Unfreeze" })).toBeVisible();

    expect(errors).toEqual([]);
  });

  test("frozen panes: coordinate semantics hold at scroll=0 and after scroll", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // Freeze row 1 + column A.
    await page.getByRole("button", { name: "Freeze row 1 + col A" }).click();
    await expect(page.getByRole("button", { name: "Unfreeze" })).toBeVisible();

    // Click the first MAIN cell (right below the frozen row, right of the
    // frozen col) → must be B2, not a duplicated A1.
    await page.mouse.click(box.x + 48 + 100 + 10, box.y + 26 + 26 + 10);
    await expect(page.getByText(/Active: B2/)).toBeVisible();

    // Click inside the frozen corner (A1 zone) → still selects A1.
    await page.mouse.click(box.x + 48 + 10, box.y + 26 + 10);
    await expect(page.getByText(/Active: A1/)).toBeVisible();

    // Scroll deep down via Ctrl+End → active cell must jump to the last
    // cell of the sheet and stay visible (scroll-to-cell respects freeze).
    await grid.click();
    await page.keyboard.press("Control+End");
    await expect(page.getByText(/Active: Z1000/)).toBeVisible();

    // The frozen corner still yields A1 after scrolling (A1 is frozen and
    // remains clickable; note the view intentionally stays scrolled since A1
    // is always visible — Excel-like semantics for frozen panes).
    await page.mouse.click(box.x + 48 + 10, box.y + 26 + 10);
    await expect(page.getByText(/Active: A1/)).toBeVisible();

    // Unfreeze → now Ctrl+Home can scroll the view back to the top-left.
    // (Toolbar buttons take focus; click the grid to refocus before keys.)
    await page.getByRole("button", { name: "Unfreeze" }).click();
    await grid.click();
    await page.keyboard.press("Control+Home");
    await expect(page.getByText(/Active: A1/)).toBeVisible();

    // Unfrozen, the point that was the first main cell (B2) is now the main
    // area's B2 still — structure change picked up by the renderer.
    await page.mouse.click(box.x + 48 + 100 + 10, box.y + 26 + 26 + 10);
    await expect(page.getByText(/Active: B2/)).toBeVisible();

    // Undo the unfreeze → frozen geometry is restored. Scroll is already at
    // 0 (set by Ctrl+Home while unfrozen), so the same point is B2 again.
    await page.getByRole("button", { name: "Undo" }).click();
    await page.mouse.click(box.x + 48 + 100 + 10, box.y + 26 + 26 + 10);
    await expect(page.getByText(/Active: B2/)).toBeVisible();

    expect(errors).toEqual([]);
  });

  test("M1.10: Ctrl+End under freeze reaches the true last scroll position", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));

    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();

    await page.getByRole("button", { name: "Freeze row 1 + col A" }).click();
    await expect(page.getByRole("button", { name: "Unfreeze" })).toBeVisible();

    // Ctrl+End on the default 1000x26 sheet (row height 26 → total 26000).
    await grid.click();
    await page.keyboard.press("Control+End");
    await expect(page.getByText(/Active: Z1000/)).toBeVisible();

    // The renderer's scroll position must equal the scrollbar's true max.
    // Before M1.10 clampScroll double-subtracted the frozen size, leaving the
    // last frozen-row height (26px) unreachable (maxScroll - 26).
    const probe = await page.evaluate(() => {
      const g = (window as unknown as { __grid: { scrollY: number; scrollbarGeometry(): { vertical: { maxScroll: number } | null } } }).__grid;
      return { scrollY: g.scrollY, maxScroll: g.scrollbarGeometry().vertical!.maxScroll };
    });
    expect(probe.scrollY).toBeCloseTo(probe.maxScroll, 5);

    expect(errors).toEqual([]);
  });

  test("M1.10: same-frame cell dirty + selection move repaints header highlight", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));

    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();

    // Step 1: click C5 → header highlight paints on the C column (own frame).
    // clientX/Y are viewport-relative, so offset by the canvas rect.
    await page.evaluate(() => {
      const overlay = document.querySelectorAll<HTMLCanvasElement>("[data-testid=sheet-grid] canvas")[1];
      const rect = overlay.getBoundingClientRect();
      overlay.dispatchEvent(
        new MouseEvent("mousedown", { clientX: rect.left + 48 + 2 * 100 + 50, clientY: rect.top + 26 + 4 * 26 + 13, button: 0, bubbles: true }),
      );
    });
    await page.waitForTimeout(80);
    await expect(page.getByText(/Active: C5/)).toBeVisible();

    // Step 2: in ONE synchronous evaluate — move the selection to D5 (headerDirty
    // + schedules a frame) then immediately write data (dirty rects + schedules
    // the same frame). Both land in the same rAF: the old else-if chain painted
    // only the dirty rects and dropped the header repaint.
    await page.evaluate(async () => {
      const overlay = document.querySelectorAll<HTMLCanvasElement>("[data-testid=sheet-grid] canvas")[1];
      const rect = overlay.getBoundingClientRect();
      overlay.dispatchEvent(
        new MouseEvent("mousedown", { clientX: rect.left + 48 + 3 * 100 + 50, clientY: rect.top + 26 + 4 * 26 + 13, button: 0, bubbles: true }),
      );
      const api = (window as unknown as { __api: { applyOperations(o: object): Promise<unknown> } }).__api;
      await api.applyOperations({
        workbookId: (window as unknown as { __workbookId: string }).__workbookId,
        sheetId: (window as unknown as { __sheetId: string }).__sheetId,
        atomic: true,
        operations: [{ type: "range.write", range: "A5:B5", values: [["x", 1]] }],
      });
    });
    await page.waitForTimeout(80);
    await expect(page.getByText(/Active: D5/)).toBeVisible();

    // D column header pixel must now be the highlight color (211,227,253),
    // not the plain background (247,248,250) it had while C was selected.
    const sample = await page.evaluate(() => {
      const content = document.querySelectorAll<HTMLCanvasElement>("[data-testid=sheet-grid] canvas")[0];
      const dpr = window.devicePixelRatio || 1;
      const px = content.getContext("2d")!.getImageData(Math.round(398 * dpr), Math.round(13 * dpr), 1, 1).data;
      return { r: px[0], g: px[1], b: px[2] };
    });
    expect(Math.abs(sample.r - 211) + Math.abs(sample.g - 227) + Math.abs(sample.b - 253)).toBeLessThan(30);

    expect(errors).toEqual([]);
  });
});
