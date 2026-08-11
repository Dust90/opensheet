import { describe, expect, it } from "vitest";
import { CommandBus, createDefaultRegistry } from "@opensheet/commands";
import { toWorkbookSnapshot, Workbook, Worksheet } from "@opensheet/core";
import { HistoryManager } from "@opensheet/history";

function setup() {
  const workbook = new Workbook({ id: "wb", name: "Book" });
  const sheet = new Worksheet({ id: "s", name: "Sheet", rowCount: 10, columnCount: 5 });
  workbook.addSheet(sheet);
  const history = new HistoryManager();
  const bus = new CommandBus(workbook, { history, registry: createDefaultRegistry() });
  const events: import("@opensheet/shared").ChangeEvent[] = [];
  workbook.onChange(e => events.push(e));
  return { workbook, sheet, bus, history, events };
}
const spec = { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, hasHeader: false, keys: [{ columnOffset: 0, direction: "asc" as const }] };

describe("range.sort", () => {
  it("moves full CellData, emits reorder, and undo/redo is exact", () => {
    const { workbook, sheet, bus, history, events } = setup();
    sheet.setCell(0, 0, { value: 2, styleId: "s2", numberFormat: "0.0" });
    sheet.setCell(0, 1, { value: 20, formula: "=A1*10" });
    sheet.setCell(1, 0, { value: 1, styleId: "s1", numberFormat: "0" });
    sheet.setCell(1, 1, { value: 10, formula: "=A2*10" });
    sheet.setCell(2, 0, { value: 3 }); sheet.rowHeights.set(0, 44);
    const before = JSON.stringify(toWorkbookSnapshot(workbook));
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "range.sort", spec }] });
    const sorted = JSON.stringify(toWorkbookSnapshot(workbook));
    expect(sheet.getCell(0, 0)).toMatchObject({ value: 1, styleId: "s1", numberFormat: "0" });
    expect(sheet.getCell(0, 1)?.formula).toBe("=A1*10");
    expect(sheet.getCell(1, 1)?.formula).toBe("=A2*10");
    expect(sheet.rowHeights.get(0)).toBe(44);
    expect(events[0]!.changes[0]!.kind).toBe("reorder");
    history.undo(bus); expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(before);
    history.redo(bus); expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(sorted);
  });

  it("rejects row-span filter conflicts and identity sorts are invisible", () => {
    const { sheet, bus, history, events } = setup();
    sheet.setCell(0, 0, { value: 1 }); sheet.setCell(1, 0, { value: 2 }); sheet.setCell(2, 0, { value: 3 });
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "range.sort", spec }] });
    expect(history.undoDepth).toBe(0); expect(events).toHaveLength(0);
    sheet.setFilter({ range: { startRow: 1, startCol: 4, endRow: 5, endCol: 4 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "isBlank" }] });
    expect(() => bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "range.sort", spec }] })).toThrow(/intersects/);
  });

  it("rolls a sort back when a later atomic operation fails", () => {
    const { sheet, bus, history, events } = setup();
    sheet.setCell(0, 0, { value: 2 }); sheet.setCell(1, 0, { value: 1 });
    expect(() => bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "range.sort", spec }, { type: "cell.set", range: "ZZ100", value: 1 }] })).toThrow();
    expect(sheet.getCell(0, 0)?.value).toBe(2);
    expect(history.undoDepth).toBe(0); expect(events).toHaveLength(0);
  });

  it("detaches its journal from later caller SortSpec mutations", () => {
    const { workbook, sheet, bus, history } = setup();
    sheet.setCell(0, 0, { value: 2 }); sheet.setCell(1, 0, { value: 1 });
    const callerSpec = { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 }, hasHeader: false, keys: [{ columnOffset: 0, direction: "asc" as const }] };
    const before = JSON.stringify(toWorkbookSnapshot(workbook));
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "range.sort", spec: callerSpec }] });
    const sorted = JSON.stringify(toWorkbookSnapshot(workbook));
    callerSpec.range.startRow = 7; callerSpec.range.endRow = 9; callerSpec.keys[0]!.columnOffset = 0;
    history.undo(bus); expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(before);
    history.redo(bus); expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(sorted);
  });
});
