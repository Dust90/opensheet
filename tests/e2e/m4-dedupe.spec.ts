import { expect, test, type Page } from "@playwright/test";

async function apply(page: Page, operations: object[]): Promise<void> {
  await page.evaluate(async (ops) => {
    const w = window as unknown as { __api: { applyOperations(value: object): Promise<unknown> }; __workbookId: string; __sheetId: string };
    await w.__api.applyOperations({ workbookId: w.__workbookId, sheetId: w.__sheetId, atomic: true, operations: ops });
  }, operations);
}

async function values(page: Page, count: number): Promise<unknown[]> {
  return page.evaluate((rows) => {
    const w = window as unknown as { __api: { getWorksheetView(id: string): { getCell(row: number, col: number): { value: unknown } | undefined } }; __sheetId: string };
    const sheet = w.__api.getWorksheetView(w.__sheetId);
    return Array.from({ length: rows }, (_, row) => sheet.getCell(row, 0)?.value);
  }, count);
}

const dedupe = { type: "range.dedupe", spec: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 }, hasHeader: false, keyColumnOffsets: [0], keep: "first" } };

test("M4.5: dedupe compacts physical rows, clears its tail, and preserves hit testing", async ({ page }) => {
  await page.goto("/");
  const grid = page.locator("[data-testid=sheet-grid]");
  await expect(grid).toBeVisible();
  await apply(page, [{ type: "range.write", range: "A1:A4", values: [["x"], ["x"], ["y"], ["z"]] }, dedupe]);
  await expect.poll(() => values(page, 4)).toEqual(["x", "y", "z", undefined]);
  const box = await grid.boundingBox(); if (box === null) throw new Error("missing grid");
  await page.mouse.click(box.x + 60, box.y + 26 + 26 + 10);
  await expect(page.getByText(/Active: A2/)).toBeVisible();
});

test("M4.5: dedupe undo and redo synchronize the grid", async ({ page }) => {
  await page.goto("/"); await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
  await apply(page, [{ type: "range.write", range: "A1:A4", values: [["x"], ["x"], ["y"], ["z"]] }]);
  await apply(page, [dedupe]); await expect.poll(() => values(page, 4)).toEqual(["x", "y", "z", undefined]);
  await page.evaluate(() => (window as unknown as { __api: { undo(): void } }).__api.undo());
  await expect.poll(() => values(page, 4)).toEqual(["x", "x", "y", "z"]);
  await page.evaluate(() => (window as unknown as { __api: { redo(): void } }).__api.redo());
  await expect.poll(() => values(page, 4)).toEqual(["x", "y", "z", undefined]);
});

test("M4.5: dedupe leaves the header fixed while compacting its body", async ({ page }) => {
  await page.goto("/"); await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
  await apply(page, [
    { type: "range.write", range: "A1:A4", values: [["Name"], ["x"], ["x"], ["y"]] },
    { type: "range.dedupe", spec: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 }, hasHeader: true, keyColumnOffsets: [0], keep: "first" } },
  ]);
  await expect.poll(() => values(page, 4)).toEqual(["Name", "x", "y", undefined]);
});
