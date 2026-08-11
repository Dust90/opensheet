import { expect, test } from "@playwright/test";

test("M4.3: sort updates the grid's physical row order", async ({ page }) => {
  await page.goto("/");
  const grid = page.locator("[data-testid=sheet-grid]");
  await expect(grid).toBeVisible();
  await page.evaluate(async () => {
    const w = window as unknown as { __api: { applyOperations(v: object): Promise<unknown> }; __workbookId: string; __sheetId: string };
    await w.__api.applyOperations({ workbookId: w.__workbookId, sheetId: w.__sheetId, atomic: true, operations: [
      { type: "range.write", range: "A1:A2", values: [[2], [1]] },
      { type: "range.sort", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 }, hasHeader: false, keys: [{ columnOffset: 0, direction: "asc" }] } },
    ] });
  });
  await expect.poll(() => page.evaluate(() => (window as unknown as { __api: { getWorksheetView(id: string): { getCell(r: number, c: number): { value: unknown } | undefined } }; __sheetId: string }).__api.getWorksheetView((window as unknown as { __sheetId: string }).__sheetId).getCell(0, 0)?.value)).toBe(1);
  const box = await grid.boundingBox(); if (box === null) throw new Error("missing grid box");
  await page.mouse.click(box.x + 48 + 10, box.y + 26 + 10);
  await expect(page.getByText(/Active: A1/)).toBeVisible();
});
