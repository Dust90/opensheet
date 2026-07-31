// M2 E2E: editing, IME, clipboard, structure/style commands, snapshot.

import { expect, test } from "@playwright/test";

const GRID = "[data-testid=sheet-grid]";

/** Read a cell value through the exposed runtime API. */
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
      return api.readRange({ sheetId, range: `${name(c)}${r + 1}` })[0][0];
    },
    [row, col] as const,
  );
}

function colName(col: number): string {
  let name = "";
  let n = col + 1;
  while (n > 0) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

test.beforeEach(async ({ context }) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"], {
    origin: "http://127.0.0.1:5173",
  });
});

test.describe("M2 editing", () => {
  test("dblclick edits and commits; Enter moves the selection down", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // dblclick B2 (col 1, row 1).
    await page.mouse.dblclick(box.x + 48 + 150, box.y + 26 + 26 + 13);
    const editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();

    await editor.fill("Hello");
    await page.keyboard.press("Enter");
    await expect(editor).not.toBeVisible();
    await expect(page.getByText(/Active: B3/)).toBeVisible();
    expect(await cellValue(page, 1, 1)).toBe("Hello");
    expect(errors).toEqual([]);
  });

  test("F2 edits the original value; Escape cancels without new history", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // Seed A1 = 42 (one history entry).
    await page.mouse.click(box.x + 48 + 50, box.y + 26 + 13);
    await page.keyboard.press("F2");
    const editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("42");
    await page.keyboard.press("Enter");
    await expect(page.getByText(/Active: A2/)).toBeVisible();
    expect(await cellValue(page, 0, 0)).toBe(42);

    // F2 again (back on A1) → editor shows the stored value.
    await page.mouse.click(box.x + 48 + 50, box.y + 26 + 13); // A1
    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await expect(editor).toHaveValue("42");

    // Escape cancels: no write, no selection move, no history entry.
    await page.keyboard.press("Escape");
    await expect(editor).not.toBeVisible();
    await expect(page.getByText(/Active: A1/)).toBeVisible();
    expect(await cellValue(page, 0, 0)).toBe(42);

    // One undo now reverts the ORIGINAL seed write — proving Escape added none.
    await page.getByRole("button", { name: "Undo" }).click();
    await page.waitForTimeout(100);
    expect(await cellValue(page, 0, 0)).toBeNull();
    expect(errors).toEqual([]);
  });

  test("Tab commits and moves right; typed number becomes numeric", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    await page.mouse.dblclick(box.x + 48 + 2 * 100 + 50, box.y + 26 + 3 * 26 + 13); // C4
    const editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("123.5");
    await page.keyboard.press("Tab");
    await expect(editor).not.toBeVisible();
    await expect(page.getByText(/Active: D4/)).toBeVisible();
    expect(await cellValue(page, 3, 2)).toBe(123.5); // typed value inferred numeric
    expect(errors).toEqual([]);
  });

  test("IME: Enter during composition never commits", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    await page.mouse.dblclick(box.x + 48 + 50, box.y + 26 + 13); // A1
    const editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();

    // Begin composition; pressing Enter must NOT commit nor move.
    await page.evaluate(() => {
      const el = document.querySelector("[data-testid=cell-editor]") as HTMLTextAreaElement;
      el.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }));
      el.value = "nihao";
      el.dispatchEvent(new KeyboardEvent("keydown", { key: "Process", bubbles: true }));
      el.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    });
    await page.waitForTimeout(100);
    await expect(editor).toBeVisible(); // still editing
    await expect(page.getByText(/Active: A1/)).toBeVisible(); // selection did not move
    expect(await cellValue(page, 0, 0)).toBeNull(); // nothing committed

    // compositionend → Enter now commits.
    await page.evaluate(() => {
      const el = document.querySelector("[data-testid=cell-editor]") as HTMLTextAreaElement;
      el.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true }));
      el.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    });
    await expect(editor).not.toBeVisible();
    await expect(page.getByText(/Active: A2/)).toBeVisible();
    expect(await cellValue(page, 0, 0)).toBe("nihao");
    expect(errors).toEqual([]);
  });
});

test.describe("M2 clipboard", () => {
  test("copy selection → paste TSV as ONE atomic write → single Undo reverts all", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // Seed a 2x2 block at A1:B2, then copy it via Ctrl+C.
    await page.evaluate(async () => {
      const api = (window as unknown as { __api: { applyOperations(o: object): Promise<unknown> } }).__api;
      const wb = (window as unknown as { __workbookId: string }).__workbookId;
      const sheet = (window as unknown as { __sheetId: string }).__sheetId;
      await api.applyOperations({
        workbookId: wb,
        sheetId: sheet,
        atomic: true,
        operations: [
          { type: "range.write", range: "A1:B2", values: [["alpha", 1], ["beta", 2]] },
        ],
      });
    });
    await page.waitForTimeout(100);

    // Select A1:B2 (shift+click extends the grid selection), then Ctrl+C.
    await page.mouse.click(box.x + 48 + 50, box.y + 26 + 13); // A1
    await page.keyboard.down("Shift");
    await page.mouse.click(box.x + 48 + 150, box.y + 26 + 26 + 13); // B2
    await page.keyboard.up("Shift");
    await page.keyboard.press("Control+c");

    // The system clipboard now holds TSV — verify the demo wrote it.
    const tsv = await page.evaluate(() => navigator.clipboard.readText());
    expect(tsv).toContain("alpha\t1");

    // Move anchor to D4 and paste.
    await page.mouse.click((await grid.boundingBox())!.x + 48 + 3 * 100 + 50, (await grid.boundingBox())!.y + 26 + 3 * 26 + 13);
    await page.keyboard.press("Control+v");
    await page.waitForTimeout(200);
    expect(await cellValue(page, 3, 3)).toBe("alpha");
    expect(await cellValue(page, 3, 4)).toBe(1);
    expect(await cellValue(page, 4, 3)).toBe("beta");
    expect(await cellValue(page, 4, 4)).toBe(2);

    // ONE undo reverts the whole paste (one atomic transaction).
    await page.getByRole("button", { name: "Undo" }).click();
    await page.waitForTimeout(100);
    expect(await cellValue(page, 3, 3)).toBeNull();
    expect(await cellValue(page, 4, 4)).toBeNull();
    expect(errors).toEqual([]);
  });
});

test.describe("M2 structure & style", () => {
  test("insert/delete rows and columns undo/redo fully restores data", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // Seed A1:B3.
    await page.evaluate(async () => {
      const api = (window as unknown as { __api: { applyOperations(o: object): Promise<unknown> } }).__api;
      const wb = (window as unknown as { __workbookId: string }).__workbookId;
      const sheet = (window as unknown as { __sheetId: string }).__sheetId;
      await api.applyOperations({
        workbookId: wb,
        sheetId: sheet,
        atomic: true,
        operations: [{ type: "range.write", range: "A1:B3", values: [["a", 1], ["b", 2], ["c", 3]] }],
      });
    });
    await page.waitForTimeout(100);

    // Select B2 (index row 1) then insert a row above it.
    await page.mouse.click(box.x + 48 + 150, box.y + 26 + 26 + 13);
    await page.getByRole("button", { name: "⤵ Row+" }).click();
    await page.waitForTimeout(150);
    expect(await cellValue(page, 0, 0)).toBe("a");
    expect(await cellValue(page, 2, 0)).toBe("b"); // shifted down one
    expect(await cellValue(page, 3, 0)).toBe("c");

    // Undo insert → rows return.
    await page.getByRole("button", { name: "Undo" }).click();
    await page.waitForTimeout(150);
    expect(await cellValue(page, 0, 0)).toBe("a");
    expect(await cellValue(page, 1, 0)).toBe("b");
    expect(await cellValue(page, 2, 0)).toBe("c");

    // Insert a column before B (active col was B).
    await page.getByRole("button", { name: "⤴ Col+" }).click();
    await page.waitForTimeout(150);
    expect(await cellValue(page, 0, 0)).toBe("a");
    expect(await cellValue(page, 0, 2)).toBe(1); // shifted right
    await page.getByRole("button", { name: "Undo" }).click();
    await page.waitForTimeout(150);
    expect(await cellValue(page, 0, 1)).toBe(1);

    // Delete row 1 (index 0) — 'a' disappears, rest shift up.
    await page.mouse.click(box.x + 48 + 50, box.y + 26 + 13); // A1
    await page.getByRole("button", { name: "Row−" }).click();
    await page.waitForTimeout(150);
    expect(await cellValue(page, 0, 0)).toBe("b");
    await page.getByRole("button", { name: "Undo" }).click();
    await page.waitForTimeout(150);
    expect(await cellValue(page, 0, 0)).toBe("a");
    expect(await cellValue(page, 1, 0)).toBe("b");

    // Redo re-applies the deletion.
    await page.getByRole("button", { name: "Redo" }).click();
    await page.waitForTimeout(150);
    expect(await cellValue(page, 0, 0)).toBe("b");
    expect(errors).toEqual([]);
  });

  test("style changes survive undo/redo", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // Select A1 and apply bold via toolbar.
    await page.mouse.click(box.x + 48 + 50, box.y + 26 + 13);
    await page.getByRole("button", { name: "B", exact: true }).click();
    await page.waitForTimeout(150);
    const styleId = await page.evaluate(() => {
      const api = (window as unknown as { __api: { getWorksheetView(id: string): { getCell(r: number, c: number): { styleId?: string } | undefined } } }).__api;
      const sheet = (window as unknown as { __sheetId: string }).__sheetId;
      return api.getWorksheetView(sheet).getCell(0, 0)?.styleId ?? null;
    });
    expect(styleId).not.toBeNull();
    const bold = await page.evaluate(
      ([id]) => {
        const api = (window as unknown as { __api: { resolveStyle(id: string): { bold?: boolean } | undefined } }).__api;
        return api.resolveStyle(id)?.bold ?? false;
      },
      [styleId] as const,
    );
    expect(bold).toBe(true);

    // Undo → styleId reverts to undefined (cell unset).
    await page.getByRole("button", { name: "Undo" }).click();
    await page.waitForTimeout(150);
    const afterUndo = await page.evaluate(() => {
      const api = (window as unknown as { __api: { getWorksheetView(id: string): { getCell(r: number, c: number): { styleId?: string } | undefined } } }).__api;
      const sheet = (window as unknown as { __sheetId: string }).__sheetId;
      return api.getWorksheetView(sheet).getCell(0, 0)?.styleId ?? null;
    });
    expect(afterUndo).toBeNull();

    // Redo → bold is back.
    await page.getByRole("button", { name: "Redo" }).click();
    await page.waitForTimeout(150);
    const afterRedo = await page.evaluate(
      ([id]) => {
        const api = (window as unknown as { __api: { resolveStyle(id: string): { bold?: boolean } | undefined } }).__api;
        return api.resolveStyle(id)?.bold ?? false;
      },
      [styleId] as const,
    );
    expect(afterRedo).toBe(true);
    expect(errors).toEqual([]);
  });
});

test.describe("M2 persistence", () => {
  test("committed data survives reload via debounced snapshot", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // Edit via the real editor path (which also drives autoSave).
    await page.mouse.dblclick(box.x + 48 + 50, box.y + 26 + 13);
    const editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("persisted");
    await page.keyboard.press("Enter");
    await expect(page.getByText(/Active: A2/)).toBeVisible();
    // Wait out the 400ms debounce + write.
    await page.waitForTimeout(700);

    await page.reload();
    await expect(grid).toBeVisible();
    await expect(page.getByText(/Restored from storage/)).toBeVisible();
    expect(await cellValue(page, 0, 0)).toBe("persisted");
    expect(errors).toEqual([]);
  });

  test("corrupt snapshot is rejected and does not overwrite the current workbook", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    // Seed corrupt storage BEFORE the app boots.
    await page.addInitScript(() => {
      window.localStorage.setItem("opensheet:workbook", "{ definitely not json");
    });
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    await expect(page.getByText(/New workbook/)).toBeVisible();

    // Workbook is blank (fresh), and the corrupt payload was NOT clobbered.
    expect(await cellValue(page, 0, 0)).toBeNull();
    const stored = await page.evaluate(() => window.localStorage.getItem("opensheet:workbook"));
    expect(stored).toBe("{ definitely not json");
    expect(errors).toEqual([]);
  });
});

test.describe("M2 renderer integration", () => {
  test("editor commit + selection move in one frame repaints cell AND headers", async ({ page }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(String(error)));
    await page.goto("/");
    const grid = page.locator(GRID);
    await expect(grid).toBeVisible();
    const box = (await grid.boundingBox())!;

    // Real M2 path: dblclick D5, type, Enter → commit + move to D6.
    await page.mouse.dblclick(box.x + 48 + 3 * 100 + 50, box.y + 26 + 4 * 26 + 13);
    const editor = page.locator("[data-testid=cell-editor]");
    await expect(editor).toBeVisible();
    await editor.fill("sync");
    await page.keyboard.press("Enter");
    await expect(page.getByText(/Active: D6/)).toBeVisible();
    expect(await cellValue(page, 4, 3)).toBe("sync");

    // D column header must be highlighted (content canvas pixel).
    await page.waitForTimeout(80);
    const sample = await page.evaluate(() => {
      const content = document.querySelectorAll<HTMLCanvasElement>(`${"body"} [data-testid=sheet-grid] canvas`)[0];
      const dpr = window.devicePixelRatio || 1;
      const px = content.getContext("2d")!.getImageData(Math.round(398 * dpr), Math.round(13 * dpr), 1, 1).data;
      return { r: px[0], g: px[1], b: px[2] };
    });
    // Header highlight #d3e3fd = rgb(211,227,253).
    expect(Math.abs(sample.r - 211) + Math.abs(sample.g - 227) + Math.abs(sample.b - 253)).toBeLessThan(30);
    expect(errors).toEqual([]);
  });
});
