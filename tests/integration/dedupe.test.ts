import { describe, expect, it } from "vitest";
import { CommandBus, createDefaultRegistry } from "@injoysai/opensheet-commands";
import { toWorkbookSnapshot, Workbook, Worksheet } from "@injoysai/opensheet-core";
import { HistoryManager } from "@injoysai/opensheet-history";

function setup() {
  const workbook = new Workbook({ id: "wb", name: "Book" });
  const sheet = new Worksheet({ id: "s", name: "Sheet", rowCount: 10, columnCount: 5 });
  workbook.addSheet(sheet);
  const history = new HistoryManager();
  const bus = new CommandBus(workbook, { history, registry: createDefaultRegistry() });
  const events: import("@injoysai/opensheet-shared").ChangeEvent[] = [];
  workbook.onChange((event) => events.push(event));
  return { workbook, sheet, bus, history, events };
}

const spec = {
  range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 },
  hasHeader: false,
  keyColumnOffsets: [0],
  keep: "first" as const,
};

describe("range.dedupe", () => {
  it("stably compacts full CellData, clears the tail, and undo/redo is exact", () => {
    const { workbook, sheet, bus, history, events } = setup();
    sheet.setCell(0, 0, { value: "x", styleId: "first", numberFormat: "@" });
    sheet.setCell(0, 1, { value: 1, formula: "=A1" });
    sheet.setCell(1, 0, { value: "x", styleId: "duplicate" });
    sheet.setCell(1, 1, { value: 2, formula: "=A2*2" });
    sheet.setCell(2, 0, { value: "y", styleId: "kept", numberFormat: "0.00" });
    sheet.setCell(2, 1, { value: 3, formula: "=A3" });
    sheet.setCell(3, 0, { value: "z" });
    sheet.rowHeights.set(1, 44); sheet.frozenRows = 1;
    const before = JSON.stringify(toWorkbookSnapshot(workbook));

    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "range.dedupe", spec }] });
    const compacted = JSON.stringify(toWorkbookSnapshot(workbook));
    expect(sheet.getCell(0, 0)).toMatchObject({ value: "x", styleId: "first", numberFormat: "@" });
    expect(sheet.getCell(1, 0)).toMatchObject({ value: "y", styleId: "kept", numberFormat: "0.00" });
    expect(sheet.getCell(1, 1)?.formula).toBe("=A2");
    expect(sheet.getCell(2, 0)).toMatchObject({ value: "z" });
    expect(sheet.getCell(3, 0)).toBeUndefined();
    expect(sheet.rowCount).toBe(10); expect(sheet.rowHeights.get(1)).toBe(44); expect(sheet.frozenRows).toBe(1);
    expect(events[0]!.changes[0]!.kind).toBe("reorder");
    history.undo(bus); expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(before);
    history.redo(bus); expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(compacted);
  });

  it("rejects filter row-span conflicts and leaves no-op dedupe invisible", () => {
    const { sheet, bus, history, events } = setup();
    sheet.setCell(0, 0, { value: "x" }); sheet.setCell(1, 0, { value: "y" });
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "range.dedupe", spec: { ...spec, range: { ...spec.range, endRow: 1 } } }] });
    expect(history.undoDepth).toBe(0); expect(events).toHaveLength(0);
    sheet.setFilter({ range: { startRow: 2, startCol: 4, endRow: 5, endCol: 4 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "isBlank" }] });
    expect(() => bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "range.dedupe", spec }] })).toThrow(/intersects/);
  });

  it("uses freshly flushed derived values before building the dedupe plan", () => {
    const { sheet, bus } = setup();
    sheet.setCell(0, 1, { value: 1 }); sheet.setCell(1, 1, { value: 2 });
    bus.addBeforeCommitHook(({ derived }) => derived.setComputedValue(sheet.id, 0, 1, 2));
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [
      { type: "cell.set", range: "C1", value: "trigger" },
      { type: "range.dedupe", spec: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, hasHeader: false, keyColumnOffsets: [1], keep: "first" } },
    ] });
    expect(sheet.getCell(1, 1)).toBeUndefined();
  });

  it("rolls compaction back when a later atomic operation fails", () => {
    const { workbook, sheet, bus, history, events } = setup();
    sheet.setCell(0, 0, { value: "x" }); sheet.setCell(1, 0, { value: "x" }); sheet.setCell(2, 0, { value: "y" });
    const before = JSON.stringify(toWorkbookSnapshot(workbook));
    expect(() => bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [
      { type: "range.dedupe", spec }, { type: "cell.set", range: "ZZ100", value: 1 },
    ] })).toThrow();
    expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(before);
    expect(history.undoDepth).toBe(0); expect(events).toHaveLength(0);
  });
});
