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
}

async function projection(page: Page): Promise<{ rows: number; mapped: number[] }> {
  return page.evaluate(() => {
    const g = (window as unknown as { __grid: { getRowProjection(): { visualRowCount: number; visualToPhysical(row: number): number } } }).__grid;
    const p = g.getRowProjection();
    return { rows: p.visualRowCount, mapped: Array.from({ length: Math.min(5, p.visualRowCount) }, (_, row) => p.visualToPhysical(row)) };
  });
}

async function expectProjectionRows(page: Page, rows: number): Promise<void> {
  await expect.poll(async () => (await projection(page)).rows).toBe(rows);
}

test.describe("M4.2-F runtime filter projection", () => {
  test("apply, clear, undo and redo automatically rebuild the grid projection", async ({ page }) => {
    await page.goto("/");
    await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
    await apply(page, [
      { type: "range.write", range: "A1:A4", values: [["Header"], ["east"], ["west"], ["east"]] },
      { type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 }, hasHeader: true, conditions: [{ columnOffset: 0, operator: "equals", value: "east" }] } },
    ]);
    await expectProjectionRows(page, 999);
    const filtered = await projection(page);
    expect(filtered.mapped.slice(0, 3)).toEqual([0, 1, 3]);

    await apply(page, [{ type: "filter.clear" }]);
    await expectProjectionRows(page, 1000);
    await page.evaluate(() => (window as unknown as { __api: { undo(): void } }).__api.undo());
    await expect.poll(async () => (await projection(page)).mapped.slice(0, 3)).toEqual([0, 1, 3]);
    await page.evaluate(() => (window as unknown as { __api: { redo(): void } }).__api.redo());
    await expect.poll(async () => (await projection(page)).rows).toBe(1000);
  });

  test("derived formula changes re-evaluate filtered visibility", async ({ page }) => {
    await page.goto("/");
    await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
    await apply(page, [
      { type: "cell.set", range: "A1", value: 10 },
      { type: "formula.set", range: "B1", formula: "=A1*2" },
      { type: "filter.apply", spec: { range: { startRow: 0, startCol: 1, endRow: 0, endCol: 1 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "greaterThan", value: 15 }] } },
    ]);
    await expectProjectionRows(page, 1000);
    await apply(page, [{ type: "cell.set", range: "A1", value: 2 }]);
    await expectProjectionRows(page, 999);
    await page.evaluate(() => (window as unknown as { __api: { undo(): void } }).__api.undo());
    await expect.poll(async () => (await projection(page)).rows).toBe(1000);
  });

  test("editing a predicate cell refreshes visibility and undo restores it", async ({ page }) => {
    await page.goto("/");
    await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
    await apply(page, [
      { type: "range.write", range: "A1:A2", values: [["east"], ["west"]] },
      { type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "equals", value: "east" }] } },
    ]);
    await expectProjectionRows(page, 999);

    await apply(page, [{ type: "cell.set", range: "A1", value: "west" }]);
    await expectProjectionRows(page, 998);
    await page.evaluate(() => (window as unknown as { __api: { undo(): void } }).__api.undo());
    await expectProjectionRows(page, 999);
  });

  test("structure clears the filter and undo restores the filtered projection", async ({ page }) => {
    await page.goto("/");
    await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
    await apply(page, [
      { type: "range.write", range: "A1:A4", values: [["Header"], ["east"], ["west"], ["east"]] },
      { type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 }, hasHeader: true, conditions: [{ columnOffset: 0, operator: "equals", value: "east" }] } },
    ]);
    await expectProjectionRows(page, 999);

    await apply(page, [{ type: "row.insert", at: 0, count: 1 }]);
    await expectProjectionRows(page, 1001);
    await page.evaluate(() => (window as unknown as { __api: { undo(): void } }).__api.undo());
    await expectProjectionRows(page, 999);
    await expect.poll(async () => (await projection(page)).mapped.slice(0, 3)).toEqual([0, 1, 3]);
  });

  test("restores a persisted V2 filter before the grid first becomes interactive", async ({ page }) => {
    await page.goto("/");
    await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
    await apply(page, [
      { type: "range.write", range: "A1:A2", values: [["east"], ["west"]] },
      { type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "equals", value: "east" }] } },
    ]);
    await expectProjectionRows(page, 999);

    await page.reload();
    await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
    await expect(page.getByText(/Restored from storage/)).toBeVisible();
    await expectProjectionRows(page, 999);
  });

  test("a zero-match filter keeps the empty grid safe and clear restores interaction", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();
    await apply(page, [{ type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 999, endCol: 25 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "equals", value: "missing" }] } }]);
    await expectProjectionRows(page, 0);

    const box = await grid.boundingBox();
    if (box === null) throw new Error("grid did not have a bounding box");
    await page.mouse.click(box.x + 80, box.y + 50);
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("F2");
    await apply(page, [{ type: "filter.clear" }]);
    await expectProjectionRows(page, 1000);
    expect(errors).toEqual([]);
  });

  test("filtering with frozen rows preserves visual-to-physical hit testing", async ({ page }) => {
    await page.goto("/");
    const grid = page.locator("[data-testid=sheet-grid]");
    await expect(grid).toBeVisible();
    await apply(page, [
      { type: "range.write", range: "A1:A4", values: [["Header"], ["east"], ["west"], ["east"]] },
      { type: "sheet.freeze", frozenRows: 2, frozenColumns: 1 },
      { type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 }, hasHeader: true, conditions: [{ columnOffset: 0, operator: "equals", value: "east" }] } },
    ]);
    await expectProjectionRows(page, 999);
    await expect.poll(async () => (await projection(page)).mapped.slice(0, 3)).toEqual([0, 1, 3]);

    const box = await grid.boundingBox();
    if (box === null) throw new Error("grid did not have a bounding box");
    await page.mouse.click(box.x + 48 + 100 + 10, box.y + 26 + 52 + 10);
    await expect(page.getByText(/Active: B4/)).toBeVisible();
  });
});
