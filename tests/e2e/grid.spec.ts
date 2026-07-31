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
});
