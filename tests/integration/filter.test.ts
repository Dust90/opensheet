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

  it("reports the union of non-overlapping old and new filter ranges on apply, undo, and redo", () => {
    const { sheet, history, bus, events } = setup();
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.apply", spec: FILTER_A }] });
    events.length = 0;
    const expectedRange = { startRow: 1, startCol: 0, endRow: 30, endCol: 3 };
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.apply", spec: FILTER_B }] });
    expect(events[0]!.changes).toEqual([{ range: expectedRange, kind: "filter" }]);
    history.undo(bus);
    expect(events[1]!.changes).toEqual([{ range: expectedRange, kind: "filter" }]);
    history.redo(bus);
    expect(events[2]!.changes).toEqual([{ range: expectedRange, kind: "filter" }]);
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

  it.each([
    { type: "row.insert", at: 0, count: 1, kind: "rows" },
    { type: "row.delete", at: 50, count: 50, kind: "rows" },
    { type: "column.insert", at: 0, count: 1, kind: "columns" },
    { type: "column.delete", at: 5, count: 1, kind: "columns" },
  ] as const)("$type clears Filter in the same history entry and restores it on undo", (operation) => {
    const { sheet, history, bus, events } = setup();
    const edgeFilter: FilterSpec = {
      range: { startRow: 0, startCol: 0, endRow: 99, endCol: 9 },
      hasHeader: true,
      conditions: [{ columnOffset: 0, operator: "notBlank" }],
    };
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.apply", spec: edgeFilter }] });
    events.length = 0;
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [operation] });
    expect(sheet.filter).toBeNull();
    expect(history.undoDepth).toBe(2); // Apply + one combined structure/filter entry.
    expect(events).toHaveLength(1);
    expect(events[0]!.changes.map((change) => change.kind).sort()).toEqual(["filter", operation.kind].sort());

    history.undo(bus);
    expect(sheet.filter).toEqual(edgeFilter);
    history.redo(bus);
    expect(sheet.filter).toBeNull();
  });

  it("leaves ordinary structure events and history unchanged when no Filter is active", () => {
    const { sheet, history, bus, events } = setup();
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "row.insert", at: 0 }] });
    expect(history.undoDepth).toBe(1);
    expect(events[0]!.changes.map((change) => change.kind)).toEqual(["rows"]);
  });

  it("rolls back both structure and Filter when a later atomic operation fails", () => {
    const { sheet, history, bus, events } = setup();
    const edgeFilter: FilterSpec = {
      range: { startRow: 0, startCol: 0, endRow: 99, endCol: 9 },
      hasHeader: false,
      conditions: [{ columnOffset: 0, operator: "notBlank" }],
    };
    bus.applyOperations({ sheetId: sheet.id, atomic: true, operations: [{ type: "filter.apply", spec: edgeFilter }] });
    events.length = 0;
    const historyBefore = history.undoDepth;
    expect(() => bus.applyOperations({
      sheetId: sheet.id,
      atomic: true,
      operations: [
        { type: "row.delete", at: 50, count: 50 },
        { type: "row.delete", at: 99, count: 2 }, // invalid after the first operation
      ],
    })).toThrow();
    expect(sheet.rowCount).toBe(100);
    expect(sheet.filter).toEqual(edgeFilter);
    expect(events).toHaveLength(0);
    expect(history.undoDepth).toBe(historyBefore);
  });
});
