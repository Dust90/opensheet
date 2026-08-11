import { expect, test } from "@playwright/test";

test("M5: browser CSV Worker imports a new sheet, exports values, and supports undo/redo", async ({
  page,
}) => {
  await page.goto("/");
  await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();

  const state = await page.evaluate(async () => {
    const w = window as unknown as {
      __api: {
        importCSV(options: {
          file: Blob;
        }): Promise<{ sheetId: string; rowCount: number; columnCount: number }>;
        exportCSV(options: { sheetId: string }): Promise<Blob>;
        readRange(options: { sheetId: string; range: string }): unknown[][];
        listSheets(): Array<{ id: string }>;
        undo(): void;
        redo(): void;
      };
    };
    const imported = await w.__api.importCSV({
      file: new File(["Name,Amount\r\nAda,10\r\n"], "sales-2026.csv", {
        type: "text/csv",
      }),
    });
    const rows = w.__api.readRange({
      sheetId: imported.sheetId,
      range: "A1:B2",
    });
    const csv = await w.__api.exportCSV({ sheetId: imported.sheetId });
    const beforeUndo = w.__api.listSheets().length;
    w.__api.undo();
    const afterUndo = w.__api.listSheets().length;
    w.__api.redo();
    const afterRedo = w.__api.listSheets().length;
    return {
      imported,
      rows,
      csv: await csv.text(),
      beforeUndo,
      afterUndo,
      afterRedo,
    };
  });

  expect(state.imported).toMatchObject({ rowCount: 2, columnCount: 2 });
  expect(state.rows).toEqual([
    ["Name", "Amount"],
    ["Ada", "10"],
  ]);
  expect(state.csv).toBe("Name,Amount\r\nAda,10");
  expect(state.afterUndo).toBe(state.beforeUndo - 1);
  expect(state.afterRedo).toBe(state.beforeUndo);
});

test("M5: browser plugin formula functions evaluate and unload through Runtime", async ({
  page,
}) => {
  await page.goto("/");
  await expect(page.locator("[data-testid=sheet-grid]")).toBeVisible();

  const values = await page.evaluate(async () => {
    const w = window as unknown as {
      __api: {
        usePlugin(plugin: {
          id: string;
          setup(context: {
            functions: { registerFunction(value: unknown): void };
          }): void;
        }): Promise<void>;
        disposePlugin(id: string): Promise<void>;
        applyOperations(request: object): Promise<unknown>;
        readRange(options: { sheetId: string; range: string }): unknown[][];
      };
      __workbookId: string;
      __sheetId: string;
    };
    await w.__api.usePlugin({
      id: "double-formula",
      setup(context) {
        context.functions.registerFunction({
          name: "PLUGIN_DOUBLE",
          minArgs: 1,
          maxArgs: 1,
          execute(args: unknown[]) {
            return typeof args[0] === "number"
              ? args[0] * 2
              : { type: "#VALUE!", message: "number required" };
          },
        });
      },
    });
    await w.__api.applyOperations({
      workbookId: w.__workbookId,
      sheetId: w.__sheetId,
      atomic: true,
      operations: [
        { type: "cell.set", range: "A1", value: 7 },
        { type: "formula.set", range: "B1", formula: "=PLUGIN_DOUBLE(A1)" },
      ],
    });
    const computed = w.__api.readRange({
      sheetId: w.__sheetId,
      range: "B1",
    })[0]![0];
    await w.__api.disposePlugin("double-formula");
    const unloaded = w.__api.readRange({
      sheetId: w.__sheetId,
      range: "B1",
    })[0]![0];
    return { computed, unloaded };
  });

  expect(values.computed).toBe(14);
  expect(values.unloaded).toMatchObject({ type: "#NAME?" });
});
