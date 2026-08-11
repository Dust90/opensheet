import { expect, test, type Page } from "@playwright/test";

async function apply(page: Page, operations: object[]): Promise<void> {
  await page.evaluate(async (ops) => {
    const w = window as unknown as {
      __api: { applyOperations(value: object): Promise<unknown> };
      __workbookId: string;
      __sheetId: string;
    };
    await w.__api.applyOperations({
      workbookId: w.__workbookId,
      sheetId: w.__sheetId,
      atomic: true,
      operations: ops,
    });
  }, operations);
}

async function openFind(page: Page, query: string): Promise<void> {
  // Chromium reserves Ctrl/Cmd+F under automation; the visible trigger opens
  // the same panel, while production users also have the keyboard shortcut.
  await page.getByRole("button", { name: "Find", exact: true }).click();
  const input = page.getByRole("textbox", { name: "Find" });
  await expect(input).toBeFocused();
  await input.fill(query);
}

test("M4.4: Find panel selects and reveals the next physical match", async ({ page }) => {
  await page.goto("/");
  await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
  await apply(page, [{ type: "range.write", range: "A1:A3", values: [["alpha"], ["needle"], ["omega"]] }]);

  await openFind(page, "needle");
  await page.keyboard.press("Enter");
  await expect(page.getByText(/Active: A2/)).toBeVisible();
  await expect(page.getByTestId("find-summary")).toHaveText("1 result");
});

test("M4.4: Enter and Shift+Enter navigate matches with wrap", async ({ page }) => {
  await page.goto("/");
  await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
  await apply(page, [{ type: "range.write", range: "A1:A3", values: [["needle"], ["needle"], ["other"]] }]);

  await openFind(page, "needle");
  await page.keyboard.press("Enter");
  await expect(page.getByText(/Active: A2/)).toBeVisible();
  await page.keyboard.press("Enter");
  await expect(page.getByText(/Active: A1/)).toBeVisible();
  await page.keyboard.press("Shift+Enter");
  await expect(page.getByText(/Active: A2/)).toBeVisible();
});

test("M4.4: visible Find skips matches hidden by an active filter", async ({ page }) => {
  await page.goto("/");
  await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();
  await apply(page, [
    { type: "range.write", range: "A1:B3", values: [["other", "hide"], ["needle", "show"], ["needle", "hide"]] },
    { type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, hasHeader: false, conditions: [{ columnOffset: 1, operator: "equals", value: "show" }] } },
  ]);

  await openFind(page, "needle");
  await page.keyboard.press("Enter");
  await expect(page.getByText(/Active: A2/)).toBeVisible();
  await expect(page.getByTestId("find-summary")).toHaveText("1 result");
});
