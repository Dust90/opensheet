import type { WorksheetView } from "@opensheet/core";
import type { CellValue, FilterCondition, FilterSpec } from "@opensheet/shared";
import { describe, expect, it } from "vitest";
import { evaluateVisibleRows, rowMatchesFilter } from "../filter-engine.js";
import { createOpenSheet } from "../create-opensheet.js";

const RANGE = { startRow: 10, startCol: 0, endRow: 20, endCol: 2 };

function filter(condition: FilterCondition | readonly FilterCondition[], hasHeader = false): FilterSpec {
  return {
    range: RANGE,
    hasHeader,
    conditions: Array.isArray(condition) ? condition : [condition],
  };
}

function sheet(cells: Record<string, CellValue> = {}, formulas: Record<string, string> = {}): WorksheetView {
  return {
    id: "sheet",
    name: "Filter test",
    rowCount: 25,
    columnCount: 3,
    frozenRows: 0,
    frozenColumns: 0,
    cellCount: Object.keys(cells).length,
    getCell(row, col) {
      const key = `${row}:${col}`;
      const value = cells[key];
      return value === undefined ? undefined : { value, formula: formulas[key] };
    },
  } as WorksheetView;
}

describe("filter engine", () => {
  it("exposes a distinguishable runtime projection state", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Projection" });
    await api.applyOperations({
      workbookId: workbook.id, sheetId: workbook.activeSheetId, atomic: true,
      operations: [
        { type: "range.write", range: "A1:A3", values: [["Header"], ["east"], ["west"]] },
        { type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 0 }, hasHeader: true, conditions: [{ columnOffset: 0, operator: "equals", value: "east" }] } },
      ],
    });
    expect(api.getFilterProjectionState(workbook.activeSheetId)).toMatchObject({ filter: expect.any(Object) });
    expect([...api.getFilterProjectionState(workbook.activeSheetId).visibleRows!]).toEqual([0, 1]);
    await api.applyOperations({ workbookId: workbook.id, sheetId: workbook.activeSheetId, atomic: true, operations: [{ type: "filter.clear" }] });
    expect(api.getFilterProjectionState(workbook.activeSheetId)).toEqual({ filter: null, visibleRows: null });
    await api.applyOperations({
      workbookId: workbook.id, sheetId: workbook.activeSheetId, atomic: true,
      operations: [{ type: "filter.apply", spec: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 0 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "equals", value: "missing" }] } }],
    });
    const empty = api.getFilterProjectionState(workbook.activeSheetId);
    expect(empty.filter).not.toBeNull();
    expect(empty.visibleRows).toBeInstanceOf(Uint32Array);
    expect(empty.visibleRows).toHaveLength(0);
  });

  it("keeps only filter-range rows and leaves range-external rows to the projection", () => {
    const view = sheet({ "9:0": "outside-before", "10:0": "match", "11:0": "miss", "21:0": "outside-after" });
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "equals", value: "match" }))]).toEqual([10]);
    expect(rowMatchesFilter(view, 9, filter({ columnOffset: 0, operator: "equals", value: "outside-before" }))).toBe(false);
  });

  it("uses strict primitive equality, with optional case-insensitive string equality", () => {
    const view = sheet({ "10:0": 1, "11:0": "1", "12:0": "Alpha", "13:0": true });
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "equals", value: 1 }))]).toEqual([10]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "equals", value: "alpha" }))]).toEqual([12]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "equals", value: "alpha", caseSensitive: true }))]).toEqual([]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "notEquals", value: 1 }))]).toEqual([11, 12, 13, 14, 15, 16, 17, 18, 19, 20]);
  });

  it("matches contains against display text and honors case sensitivity", () => {
    const view = sheet({ "10:0": 12345, "11:0": true, "12:0": "AlphaBeta", "13:0": null });
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "contains", value: "23" }))]).toEqual([10]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "contains", value: "rue" }))]).toEqual([11]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "contains", value: "alpha" }))]).toEqual([12]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "contains", value: "alpha", caseSensitive: true }))]).toEqual([]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "contains", value: null }))]).toEqual([]);
  });

  it("compares only finite numbers and finite numeric strings", () => {
    const view = sheet({ "10:0": 12, "11:0": "12", "12:0": "1e309", "13:0": "abc", "14:0": true });
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "greaterThan", value: 3 }))]).toEqual([10, 11]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "lessThan", value: "13" }))]).toEqual([10, 11]);
  });

  it("defines blank as null only and treats an absent cell as blank", () => {
    const view = sheet({ "10:0": null, "11:0": "", "12:0": false });
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "isBlank" }))]).toEqual([10, 13, 14, 15, 16, 17, 18, 19, 20]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "notBlank" }))]).toEqual([11, 12]);
  });

  it("combines conditions with AND across columns", () => {
    const view = sheet({ "10:0": "east", "10:1": 12, "11:0": "east", "11:1": 2, "12:0": "west", "12:1": 12 });
    expect([...evaluateVisibleRows(view, filter([
      { columnOffset: 0, operator: "equals", value: "east" },
      { columnOffset: 1, operator: "greaterThan", value: 5 },
    ]))]).toEqual([10]);
  });

  it("uses cached formula results; errors are not blank but never match comparisons", () => {
    const view = sheet(
      { "10:0": 20, "11:0": { type: "#DIV/0!", message: "division" } },
      { "10:0": "=A1*2" },
    );
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "greaterThan", value: 15 }))]).toEqual([10]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "notEquals", value: 20 }))]).toEqual([12, 13, 14, 15, 16, 17, 18, 19, 20]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "isBlank" }))]).toEqual([12, 13, 14, 15, 16, 17, 18, 19, 20]);
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "notBlank" }))]).toEqual([10, 11]);
  });

  it("keeps the header visible even when it does not match, including an all-hidden data result", () => {
    const view = sheet({ "10:0": "Header", "11:0": "no", "12:0": "no" });
    expect([...evaluateVisibleRows(view, filter({ columnOffset: 0, operator: "equals", value: "yes" }, true))]).toEqual([10]);
  });

  it("returns every data row when every condition matches", () => {
    const cells: Record<string, CellValue> = {};
    for (let row = 10; row <= 20; row++) cells[`${row}:0`] = "visible";
    expect([...evaluateVisibleRows(sheet(cells), filter({ columnOffset: 0, operator: "equals", value: "visible" }))]).toEqual(
      [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
    );
  });
});
