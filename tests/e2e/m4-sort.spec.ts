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

test("M4.3: sort undo and redo restore the first physical row", async ({ page }) => {
  await page.goto("/");
  await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
  await page.evaluate(async () => {
    const w = window as unknown as { __api: { applyOperations(v: object): Promise<unknown>; undo(): void; redo(): void }; __workbookId: string; __sheetId: string };
    await w.__api.applyOperations({ workbookId: w.__workbookId, sheetId: w.__sheetId, atomic: true, operations: [{ type: "range.write", range: "A1:A2", values: [[2], [1]] }, { type: "range.sort", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 }, hasHeader: false, keys: [{ columnOffset: 0, direction: "asc" }] } }] });
    w.__api.undo(); w.__api.redo();
  });
  await expect.poll(() => page.evaluate(() => (window as unknown as { __api: { getWorksheetView(id: string): { getCell(r: number, c: number): { value: unknown } | undefined } }; __sheetId: string }).__api.getWorksheetView((window as unknown as { __sheetId: string }).__sheetId).getCell(0, 0)?.value)).toBe(1);
});

test("M4.3: header stays fixed while body sorts", async ({ page }) => {
  await page.goto("/"); await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
  await page.evaluate(async () => { const w = window as unknown as { __api: { applyOperations(v: object): Promise<unknown> }; __workbookId: string; __sheetId: string }; await w.__api.applyOperations({ workbookId:w.__workbookId,sheetId:w.__sheetId,atomic:true,operations:[{type:"range.write",range:"A1:A4",values:[["Name"],["Charlie"],["Alice"],["Bob"]]},{type:"range.sort",spec:{range:{startRow:0,startCol:0,endRow:3,endCol:0},hasHeader:true,keys:[{columnOffset:0,direction:"asc"}]}}]}); });
  await expect.poll(() => page.evaluate(() => { const w=window as unknown as {__api:{getWorksheetView(id:string):{getCell(r:number,c:number):{value:unknown}|undefined}};__sheetId:string}; return [0,1,2,3].map(r=>w.__api.getWorksheetView(w.__sheetId).getCell(r,0)?.value); })).toEqual(["Name","Alice","Bob","Charlie"]);
});
