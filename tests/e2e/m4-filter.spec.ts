import { expect, test, type Page } from "@playwright/test";

async function apply(page: Page, operations: object[]): Promise<void> {
  await page.evaluate(async (ops) => {
    const w = window as unknown as {
      __api: { applyOperations(value: object): Promise<unknown> };
      __workbookId: string;
      __sheetId: string;
    };
    await w.__api.applyOperations({ workbookId: w.__workbookId, sheetId: w.__sheetId, atomic: true, operations: ops });
  }, operations);
  // Host projection updates are coalesced in a microtask after the event.
  await page.waitForTimeout(25);
}

async function projection(page: Page): Promise<{ rows: number; mapped: number[] }> {
  return page.evaluate(() => {
    const g = (window as unknown as { __grid: { getRowProjection(): { visualRowCount: number; visualToPhysical(row: number): number } } }).__grid;
    const p = g.getRowProjection();
    return { rows: p.visualRowCount, mapped: Array.from({ length: Math.min(5, p.visualRowCount) }, (_, row) => p.visualToPhysical(row)) };
  });
}

test.describe("M4.2-F runtime filter projection", () => {
  test("apply, clear, undo and redo automatically rebuild the grid projection", async ({ page }) => {
    await page.goto("/");
    await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
    await apply(page, [
      { type: "range.write", range: "A1:A4", values: [["Header"], ["east"], ["west"], ["east"]] },
      { type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 }, hasHeader: true, conditions: [{ columnOffset: 0, operator: "equals", value: "east" }] } },
    ]);
    const filtered = await projection(page);
    expect(filtered.rows).toBe(999);
    expect(filtered.mapped.slice(0, 3)).toEqual([0, 1, 3]);

    await apply(page, [{ type: "filter.clear" }]);
    expect((await projection(page)).rows).toBe(1000);
    await page.evaluate(() => (window as unknown as { __api: { undo(): void } }).__api.undo());
    await page.waitForTimeout(25);
    expect((await projection(page)).mapped.slice(0, 3)).toEqual([0, 1, 3]);
    await page.evaluate(() => (window as unknown as { __api: { redo(): void } }).__api.redo());
    await page.waitForTimeout(25);
    expect((await projection(page)).rows).toBe(1000);
  });

  test("derived formula changes re-evaluate filtered visibility", async ({ page }) => {
    await page.goto("/");
    await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
    await apply(page, [
      { type: "cell.set", range: "A1", value: 10 },
      { type: "formula.set", range: "B1", formula: "=A1*2" },
      { type: "filter.apply", spec: { range: { startRow: 0, startCol: 1, endRow: 0, endCol: 1 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "greaterThan", value: 15 }] } },
    ]);
    expect((await projection(page)).rows).toBe(1000);
    await apply(page, [{ type: "cell.set", range: "A1", value: 2 }]);
    expect((await projection(page)).rows).toBe(999);
    await page.evaluate(() => (window as unknown as { __api: { undo(): void } }).__api.undo());
    await page.waitForTimeout(25);
    expect((await projection(page)).rows).toBe(1000);
  });
});
