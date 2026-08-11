import { describe, expect, it } from "vitest";
import { ApplyOperationsError, CommandBus, createDefaultRegistry } from "@opensheet/commands";
import { Workbook, Worksheet } from "@opensheet/core";
import { HistoryManager } from "@opensheet/history";
import type { ChangeEvent, FilterSpec } from "@opensheet/shared";

const FILTER_A: FilterSpec = {
  range: { startRow: 1, startCol: 0, endRow: 20, endCol: 2 },
  hasHeader: true,
  conditions: [{ columnOffset: 0, operator: "equals", value: "east" }],
};
const FILTER_B: FilterSpec = {
  range: { startRow: 4, startCol: 1, endRow: 30, endCol: 3 },
  hasHeader: false,
  conditions: [{ columnOffset: 2, operator: "greaterThan", value: 10 }],
};

function setup() {
  const workbook = new Workbook({ id: "wb", name: "Book" });
  const sheet = new Worksheet({ id: "s1", name: "Sheet", rowCount: 100, columnCount: 10 });
  workbook.addSheet(sheet);
  const history = new HistoryManager();
  const bus = new CommandBus(workbook, { history, registry: createDefaultRegistry() });
  const events: ChangeEvent[] = [];
  workbook.onChange((event) => events.push(event));
  return { sheet, history, bus, events };
}

describe("filter commands", () => {
  it("applies, replaces, and undo/redoes complete FilterSpecs", () => {
    const { sheet, history, bus, events } = setup();
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.apply", spec: FILTER_A }] });
    expect(sheet.filter).toEqual(FILTER_A);
    expect(events).toHaveLength(1);
    expect(events[0]!.changes).toEqual([{ range: FILTER_A.range, kind: "filter" }]);
    expect(history.undoDepth).toBe(1);

    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.apply", spec: FILTER_B }] });
    expect(sheet.filter).toEqual(FILTER_B);
    history.undo(bus);
    expect(sheet.filter).toEqual(FILTER_A);
    history.undo(bus);
    expect(sheet.filter).toBeNull();
    history.redo(bus);
    expect(sheet.filter).toEqual(FILTER_A);
    history.redo(bus);
    expect(sheet.filter).toEqual(FILTER_B);
  });

  it("clears with one inverse entry, but an already-clear sheet is a true no-op", () => {
    const { sheet, history, bus, events } = setup();
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.clear" }] });
    expect(events).toHaveLength(0);
    expect(history.undoDepth).toBe(0);

    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.apply", spec: FILTER_A }] });
    events.length = 0;
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.clear" }] });
    expect(sheet.filter).toBeNull();
    expect(events).toHaveLength(1);
    expect(events[0]!.changes[0]!.kind).toBe("filter");
    expect(history.undoDepth).toBe(2);
    history.undo(bus);
    expect(sheet.filter).toEqual(FILTER_A);
  });

  it("rejects an invalid spec atomically, without event or history", () => {
    const { sheet, history, bus, events } = setup();
    expect(() => bus.applyOperations({
      sheetId: sheet.id,
      atomic: true,
      operations: [{ type: "filter.apply", spec: { ...FILTER_A, range: { ...FILTER_A.range, endRow: 100 } } }],
    })).toThrow(ApplyOperationsError);
    expect(sheet.filter).toBeNull();
    expect(events).toHaveLength(0);
    expect(history.undoDepth).toBe(0);
  });
});
