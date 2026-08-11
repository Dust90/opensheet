import { expect, test } from "@playwright/test";

const rows = 100_000;
test("dedupe perf: 100k × 20 with 50k duplicate rows", async ({ page }) => {
  await page.goto("/?historyMaxBytes=536870912"); await page.getByRole("button", { name: "Load 100k×20" }).click();
  await page.waitForSelector("text=/Loaded 100,000×20/", { timeout: 300_000 });
  const result = await page.evaluate(async () => {
    const w = window as unknown as { __api: { applyOperations(value: object): Promise<unknown>; undo(): void; redo(): void; getWorksheetView(id: string): { getCell(row: number, col: number): { value: unknown } | undefined } }; __workbookId: string; __sheetId: string };
    const write = Array.from({ length: 100_000 }, (_, row) => [`key-${row % 50_000}`]);
    await w.__api.applyOperations({ workbookId: w.__workbookId, sheetId: w.__sheetId, atomic: true, operations: [{ type: "range.write", range: "A1:A100000", values: write }] });
    const now = performance.now();
    await w.__api.applyOperations({ workbookId: w.__workbookId, sheetId: w.__sheetId, atomic: true, operations: [{ type: "range.dedupe", spec: { range: { startRow: 0, startCol: 0, endRow: 99_999, endCol: 19 }, hasHeader: false, keyColumnOffsets: [0], keep: "first" } }] });
    const dedupeMs = performance.now() - now;
    const undoStart = performance.now(); w.__api.undo(); const undoMs = performance.now() - undoStart;
    const undoValue = w.__api.getWorksheetView(w.__sheetId).getCell(50_000, 0)?.value;
    const redoStart = performance.now(); w.__api.redo(); const redoMs = performance.now() - redoStart;
    return { dedupeMs, undoMs, redoMs, first: w.__api.getWorksheetView(w.__sheetId).getCell(0, 0)?.value, undoValue, tail: w.__api.getWorksheetView(w.__sheetId).getCell(50_000, 0) };
  });
  expect(result.first).toBe("key-0"); expect(result.undoValue).toBe("key-0"); expect(result.tail).toBeUndefined();
  console.log(JSON.stringify(result));
  expect(result.dedupeMs).toBeLessThanOrEqual(1500); expect(result.undoMs).toBeLessThanOrEqual(1500); expect(result.redoMs).toBeLessThanOrEqual(1500);
});
