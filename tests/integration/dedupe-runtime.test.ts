import { describe, expect, it } from "vitest";
import { createOpenSheet } from "@opensheet/runtime";

describe("range.dedupe runtime formulas", () => {
  it("uses fresh formula values from preceding atomic operations as dedupe keys", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Dedupe" });
    const id = workbook.activeSheetId;
    const apply = (operations: Parameters<typeof api.applyOperations>[0]["operations"]) =>
      api.applyOperations({ workbookId: workbook.id, sheetId: id, atomic: true, operations });
    await apply([
      { type: "cell.set", range: "A1", value: "first" }, { type: "formula.set", range: "B1", formula: "=C1" }, { type: "cell.set", range: "C1", value: 1 },
      { type: "cell.set", range: "A2", value: "second" }, { type: "formula.set", range: "B2", formula: "=C2" }, { type: "cell.set", range: "C2", value: 2 },
    ]);
    await apply([
      { type: "cell.set", range: "C1", value: 2 },
      { type: "range.dedupe", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 2 }, hasHeader: false, keyColumnOffsets: [1], keep: "first" } },
    ]);
    const sheet = api.getWorksheetView(id);
    expect(sheet.getCell(0, 0)?.value).toBe("first");
    expect(sheet.getCell(1, 0)).toBeUndefined();
  });

  it("rebuilds moved and external formula dependencies across undo and redo", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Dedupe" });
    const id = workbook.activeSheetId;
    const apply = (operations: Parameters<typeof api.applyOperations>[0]["operations"]) =>
      api.applyOperations({ workbookId: workbook.id, sheetId: id, atomic: true, operations });
    await apply([
      { type: "cell.set", range: "A1", value: "x" }, { type: "formula.set", range: "B1", formula: "=C1*10" }, { type: "cell.set", range: "C1", value: 2 },
      { type: "cell.set", range: "A2", value: "x" }, { type: "formula.set", range: "B2", formula: "=C2*10" }, { type: "cell.set", range: "C2", value: 9 },
      { type: "cell.set", range: "A3", value: "y" }, { type: "formula.set", range: "B3", formula: "=C3*10" }, { type: "cell.set", range: "C3", value: 1 },
      { type: "formula.set", range: "E1", formula: "=A2" },
    ]);
    await apply([{ type: "range.dedupe", spec: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 }, hasHeader: false, keyColumnOffsets: [0], keep: "first" } }]);
    let sheet = api.getWorksheetView(id);
    expect(sheet.getCell(1, 0)?.value).toBe("y");
    expect(sheet.getCell(1, 1)).toMatchObject({ formula: "=C2*10", value: 10 });
    expect(sheet.getCell(0, 4)).toMatchObject({ formula: "=A2", value: "y" });
    api.undo();
    sheet = api.getWorksheetView(id);
    expect(sheet.getCell(2, 1)).toMatchObject({ formula: "=C3*10", value: 10 });
    api.redo();
    await apply([{ type: "cell.set", range: "A2", value: "z" }, { type: "cell.set", range: "C2", value: 4 }]);
    sheet = api.getWorksheetView(id);
    expect(sheet.getCell(1, 1)?.value).toBe(40);
    expect(sheet.getCell(0, 4)?.value).toBe("z");
  });

  it("recalculates translated #REF! formulas and restores their source on undo", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Dedupe" });
    const id = workbook.activeSheetId;
    const apply = (operations: Parameters<typeof api.applyOperations>[0]["operations"]) =>
      api.applyOperations({ workbookId: workbook.id, sheetId: id, atomic: true, operations });
    await apply([
      { type: "cell.set", range: "A1", value: "x" },
      { type: "cell.set", range: "A2", value: "x" },
      { type: "cell.set", range: "A3", value: "y" },
      { type: "formula.set", range: "B3", formula: "=A1" },
    ]);
    await apply([{ type: "range.dedupe", spec: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, hasHeader: false, keyColumnOffsets: [0], keep: "first" } }]);
    expect(api.getWorksheetView(id).getCell(1, 1)).toMatchObject({ formula: "=#REF!", value: { type: "#REF!" } });
    api.undo();
    expect(api.getWorksheetView(id).getCell(2, 1)).toMatchObject({ formula: "=A1", value: "x" });
  });
});
