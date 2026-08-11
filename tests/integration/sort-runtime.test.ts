import { describe, expect, it } from "vitest";
import { createOpenSheet } from "@opensheet/runtime";

describe("range.sort runtime formulas", () => {
  it("sorts by fresh computed values from preceding atomic operations", async () => {
    const api = createOpenSheet(); const wb = api.createWorkbook({ name: "Sort" }); const id = wb.activeSheetId;
    await api.applyOperations({ workbookId: wb.id, sheetId: id, atomic: true, operations: [
      { type: "cell.set", range: "A1", value: 1 }, { type: "formula.set", range: "B1", formula: "=A1" },
      { type: "cell.set", range: "A2", value: 2 }, { type: "formula.set", range: "B2", formula: "=A2" },
    ] });
    await api.applyOperations({ workbookId: wb.id, sheetId: id, atomic: true, operations: [
      { type: "cell.set", range: "A1", value: 3 },
      { type: "range.sort", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, hasHeader: false, keys: [{ columnOffset: 1, direction: "asc" }] } },
    ] });
    const sheet = api.getWorksheetView(id);
    expect(sheet.getCell(0, 0)?.value).toBe(2); expect(sheet.getCell(0, 1)?.value).toBe(2);
    expect(sheet.getCell(1, 0)?.value).toBe(3); expect(sheet.getCell(1, 1)?.value).toBe(3);
  });

  it("rebuilds moved and external formula dependencies across undo and redo", async () => {
    const api = createOpenSheet(); const wb = api.createWorkbook({ name: "Sort" }); const id = wb.activeSheetId;
    const apply = (operations: Parameters<typeof api.applyOperations>[0]["operations"]) => api.applyOperations({ workbookId: wb.id, sheetId: id, atomic: true, operations });
    await apply([
      { type: "cell.set", range: "A1", value: 2 }, { type: "formula.set", range: "B1", formula: "=A1*10" },
      { type: "cell.set", range: "A2", value: 1 }, { type: "formula.set", range: "B2", formula: "=A2*10" },
      { type: "formula.set", range: "E1", formula: "=A1" },
    ]);
    await apply([{ type: "range.sort", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, hasHeader: false, keys: [{ columnOffset: 0, direction: "asc" }] } }]);
    let sheet = api.getWorksheetView(id);
    expect(sheet.getCell(0, 1)).toMatchObject({ formula: "=A1*10", value: 10 });
    expect(sheet.getCell(1, 1)).toMatchObject({ formula: "=A2*10", value: 20 });
    expect(sheet.getCell(0, 4)).toMatchObject({ formula: "=A1", value: 1 });
    api.undo(); sheet = api.getWorksheetView(id);
    expect(sheet.getCell(0, 1)).toMatchObject({ formula: "=A1*10", value: 20 });
    api.redo(); sheet = api.getWorksheetView(id);
    expect(sheet.getCell(0, 0)?.value).toBe(1); expect(sheet.getCell(0, 4)?.value).toBe(1);
    await apply([{ type: "cell.set", range: "A1", value: 4 }]);
    sheet = api.getWorksheetView(id); expect(sheet.getCell(0, 1)?.value).toBe(40); expect(sheet.getCell(0, 4)?.value).toBe(4);
  });

  it("recalculates translated #REF! formulas and restores them on undo", async () => {
    const api = createOpenSheet(); const wb = api.createWorkbook({ name: "Sort" }); const id = wb.activeSheetId;
    await api.applyOperations({ workbookId: wb.id, sheetId: id, atomic: true, operations: [
      { type: "cell.set", range: "A1", value: 2 }, { type: "cell.set", range: "A2", value: 1 },
      { type: "formula.set", range: "B2", formula: "=A1" },
    ] });
    await api.applyOperations({ workbookId: wb.id, sheetId: id, atomic: true, operations: [{ type: "range.sort", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, hasHeader: false, keys: [{ columnOffset: 0, direction: "asc" }] } }] });
    expect(api.getWorksheetView(id).getCell(0, 1)).toMatchObject({ formula: "=#REF!", value: { type: "#REF!" } });
    api.undo(); expect(api.getWorksheetView(id).getCell(1, 1)).toMatchObject({ formula: "=A1", value: 2 });
  });

  it("rolls back an intermediate derived flush when a later operation fails", async () => {
    const api = createOpenSheet(); const wb = api.createWorkbook({ name: "Sort" }); const id = wb.activeSheetId;
    await api.applyOperations({ workbookId: wb.id, sheetId: id, atomic: true, operations: [
      { type: "cell.set", range: "A1", value: 1 }, { type: "formula.set", range: "B1", formula: "=A1" },
      { type: "cell.set", range: "A2", value: 2 }, { type: "formula.set", range: "B2", formula: "=A2" },
    ] });
    const before = JSON.stringify(api.getWorkbookSnapshot()); const events: unknown[] = []; const stop = api.onChange(e => events.push(e));
    await expect(api.applyOperations({ workbookId: wb.id, sheetId: id, atomic: true, operations: [
      { type: "cell.set", range: "A1", value: 3 },
      { type: "range.sort", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, hasHeader: false, keys: [{ columnOffset: 1, direction: "asc" }] } },
      { type: "cell.set", range: "ZZ100", value: 1 },
    ] })).rejects.toThrow();
    stop(); expect(JSON.stringify(api.getWorkbookSnapshot())).toBe(before); expect(events).toEqual([]);
  });
});
