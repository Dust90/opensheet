import { describe, expect, it } from "vitest";
import { SheetError } from "@opensheet/shared";
import { createOpenSheet } from "../create-opensheet.js";

function namedCSV(text: string, name: string): Blob {
  return Object.assign(new Blob([text], { type: "text/csv" }), { name });
}

describe("OpenSheetAPI.importCSV", () => {
  it("imports a Blob into a new named worksheet and restores it with one undo/redo", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      atomic: true,
      operations: [{ type: "cell.set", range: "A1", value: "existing" }],
    });

    const imported = await api.importCSV({ file: namedCSV("Name,Amount\r\nAda,10\r\n", "sales-2026.csv") });
    expect(imported).toEqual({ sheetId: imported.sheetId, rowCount: 2, columnCount: 2 });
    expect(api.listSheets().map((sheet) => sheet.name)).toEqual(["Sheet1", "sales-2026"]);
    expect(api.readRange({ sheetId: imported.sheetId, range: "A1:B2" })).toEqual([["Name", "Amount"], ["Ada", "10"]]);
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "A1" })).toEqual([["existing"]]);

    api.undo();
    expect(api.listSheets().map((sheet) => sheet.name)).toEqual(["Sheet1"]);
    api.redo();
    expect(api.readRange({ sheetId: imported.sheetId, range: "A1:B2" })).toEqual([["Name", "Amount"], ["Ada", "10"]]);
  });

  it("does not create a worksheet or emit an event when streaming CSV parsing fails", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    const events: unknown[] = [];
    api.onChange((event) => events.push(event));

    await expect(api.importCSV({ file: namedCSV('a,"unterminated', "bad.csv") })).rejects.toBeInstanceOf(SheetError);
    expect(api.listSheets()).toHaveLength(1);
    expect(events).toEqual([]);
  });

  it("rejects native Worker failures without exposing a partial imported sheet", async () => {
    const api = createOpenSheet();
    api.createWorkbook({ name: "Book" });
    const originalWorker = globalThis.Worker;
    let terminated = false;
    class FailingWorker {
      private readonly listeners = new Set<(event: ErrorEvent) => void>();
      constructor(..._args: unknown[]) {}
      postMessage(): void {
        for (const listener of this.listeners) listener({ message: "worker bootstrap failed" } as ErrorEvent);
      }
      addEventListener(type: string, listener: unknown): void {
        if (type === "error") this.listeners.add(listener as (event: ErrorEvent) => void);
      }
      removeEventListener(type: string, listener: unknown): void {
        if (type === "error") this.listeners.delete(listener as (event: ErrorEvent) => void);
      }
      terminate(): void { terminated = true; }
    }
    Object.defineProperty(globalThis, "Worker", { configurable: true, value: FailingWorker });
    try {
      await expect(api.importCSV({ file: namedCSV("a,b", "broken-worker.csv") })).rejects.toMatchObject({
        code: "E_OP_FAILED",
        message: "worker bootstrap failed",
      });
      expect(api.listSheets()).toHaveLength(1);
      expect(terminated).toBe(true);
    } finally {
      Object.defineProperty(globalThis, "Worker", { configurable: true, value: originalWorker });
    }
  });

  it("derives a collision-free fallback name and preserves raw empty fields", async () => {
    const api = createOpenSheet();
    api.createWorkbook({ name: "Book" });
    const first = await api.importCSV({ file: namedCSV("a,\r\n", "data.csv") });
    const duplicateName = await api.importCSV({ file: namedCSV("b", "data.csv") });
    const second = await api.importCSV({ file: new Blob(["x"]) });
    expect(api.listSheets().map((sheet) => sheet.name)).toEqual(["Sheet1", "data", "data (2)", "Imported CSV"]);
    expect(api.readRange({ sheetId: first.sheetId, range: "A1:B1" })).toEqual([["a", ""]]);
    expect(api.readRange({ sheetId: duplicateName.sheetId, range: "A1" })).toEqual([["b"]]);
    expect(second.rowCount).toBe(1);
    expect(second.columnCount).toBe(1);
  });
});

describe("OpenSheetAPI.exportCSV", () => {
  it("exports A1-based used range, retaining internal blanks and explicit empty strings", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      atomic: true,
      operations: [
        { type: "cell.set", range: "A1", value: "a" },
        { type: "cell.set", range: "C1", value: "" },
        { type: "cell.set", range: "A3", value: "c" },
        { type: "range.style", range: "Z100", style: { bold: true } },
      ],
    });

    const csv = await api.exportCSV({ sheetId: workbook.activeSheetId });
    expect(await csv.text()).toBe("a,,\r\n,,\r\nc,,");
  });

  it("exports formula computed values and a zero-byte Blob for an empty sheet", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      atomic: true,
      operations: [
        { type: "cell.set", range: "A1", value: 20 },
        { type: "cell.set", range: "A2", value: 22 },
        { type: "formula.set", range: "B1", formula: "=SUM(A1:A2)" },
        { type: "formula.set", range: "C2", formula: "=1/0" },
      ],
    });
    const csv = await api.exportCSV({ sheetId: workbook.activeSheetId });
    expect(await csv.text()).toBe("20,42,\r\n22,,#DIV/0!");

    const empty = api.createSheet({ name: "Empty" });
    const emptyCSV = await api.exportCSV({ sheetId: empty.id });
    expect(emptyCSV.size).toBe(0);
    expect(await emptyCSV.text()).toBe("");
  });
});
