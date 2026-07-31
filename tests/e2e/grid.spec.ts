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
});
