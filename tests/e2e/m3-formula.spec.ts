// M3 E2E: formula entry, recalculation, undo/redo, editor display semantics.

import { expect, test } from "@playwright/test";

const GRID = "[data-testid=sheet-grid]";

async function cellValue(page: import("@playwright/test").Page, row: number, col: number): Promise<unknown> {
  return page.evaluate(
    ([r, c]) => {
      const name = (n: number): string => {
        let out = "";
        let v = n + 1;
        while (v > 0) {
          out = String.fromCharCode(65 + ((v - 1) % 26)) + out;
          v = Math.floor((v - 1) / 26);
        }
        return out;
      };
      const api = (window as unknown as { __api: { readRange(o: { sheetId: string; range: string }): unknown[][] } }).__api;
      const sheetId = (window as unknown as { __sheetId: string }).__sheetId;
      return api.readRange({ sheetId, range: `${name(c)}${r + 1}` })[0]![0]!;
    },
    [row, col] as const,
  );
}

test.describe("M3 formulas", () => {
  test("typing =A1+1 into B1 computes and shows the result; recalc on A1 edit", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // Seed A1 = 2 via the editor, then put =A1*3 into B1.
    await page.mouse.dblclick(box.x + 48 + 50, box.y + 26 + 13);
    let editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("2");
    await page.keyboard.press("Enter");
    await expect(page.getByText(/Active: A2/)).toBeVisible();

    await page.mouse.dblclick(box.x + 48 + 150, box.y + 26 + 13); // B1
    editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("=A1*3");
    await page.keyboard.press("Enter");
    await expect(page.getByText(/Active: B2/)).toBeVisible();
    expect(await cellValue(page, 0, 1)).toBe(6); // recalculated in the same commit

    // F2 on B1 shows the FORMULA source, not the computed value.
    await page.mouse.click(box.x + 48 + 150, box.y + 26 + 13); // B1
    await page.keyboard.press("F2");
    editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await expect(editor).toHaveValue("=A1*3");
    await page.keyboard.press("Escape");

    // Editing A1 recomputes B1.
    await page.mouse.dblclick(box.x + 48 + 50, box.y + 26 + 13); // A1
    editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("5");
    await page.keyboard.press("Enter");
    expect(await cellValue(page, 0, 1)).toBe(15);

    // One undo reverts the A1 edit AND the recomputed B1.
    await page.getByRole("button", { name: "Undo" }).click();
    await page.waitForTimeout(100);
    expect(await cellValue(page, 0, 0)).toBe(2);
    expect(await cellValue(page, 0, 1)).toBe(6);
    expect(errors).toEqual([]);
  });

  test("syntax-error formula is rejected; the cell keeps its old value", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    await page.mouse.dblclick(box.x + 48 + 50, box.y + 26 + 13); // A1
    let editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("keep");
    await page.keyboard.press("Enter");

    await page.mouse.dblclick(box.x + 48 + 50, box.y + 26 + 13);
    editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("=SUM(");
    await page.keyboard.press("Enter");
    // Command rejected → cell unchanged.
    expect(await cellValue(page, 0, 0)).toBe("keep");
    expect(errors).toEqual([]);
  });
});
