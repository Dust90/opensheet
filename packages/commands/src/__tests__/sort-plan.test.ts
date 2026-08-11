import { Worksheet } from "@opensheet/core";
import { describe, expect, it } from "vitest";
import { buildSortPlan, conflictsWithFilter, rowSpansIntersect } from "../sort-plan.js";

function sheet(values: unknown[][]): Worksheet {
  const result = new Worksheet({ id: "sheet", name: "Sheet", rowCount: values.length, columnCount: values[0]?.length ?? 1 });
  values.forEach((row, r) => row.forEach((value, c) => {
    if (value !== undefined) result.setCell(r, c, { value: value as never });
  }));
  return result;
}

function plan(values: unknown[][], keys: { columnOffset: number; direction: "asc" | "desc" }[], hasHeader = false) {
  const source = sheet(values);
  return buildSortPlan(source, { range: { startRow: 0, startCol: 0, endRow: values.length - 1, endCol: values[0]!.length - 1 }, hasHeader, keys });
}

describe("buildSortPlan", () => {
  it("sorts multiple cached keys stably and provides both permutations", () => {
    const result = plan([[2, "b"], [1, "z"], [2, "a"], [2, "a"]], [{ columnOffset: 0, direction: "asc" }, { columnOffset: 1, direction: "asc" }]);
    expect([...result.destinationToSource]).toEqual([1, 2, 3, 0]);
    expect([...result.sourceToDestination]).toEqual([3, 0, 1, 2]);
    expect(result.movedRows).toBe(4);
    for (let source = 0; source < result.bodyRowCount; source += 1) {
      expect(result.destinationToSource[result.sourceToDestination[source]!]!).toBe(source);
    }
    for (let destination = 0; destination < result.bodyRowCount; destination += 1) {
      expect(result.sourceToDestination[result.destinationToSource[destination]!]!).toBe(destination);
    }
  });

  it("keeps a header fixed and treats an identity plan as a true no-op", () => {
    const result = plan([["Header"], [1], [2]], [{ columnOffset: 0, direction: "asc" }], true);
    expect(result.bodyStartRow).toBe(1);
    expect(result.bodyRowCount).toBe(2);
    expect([...result.destinationToSource]).toEqual([0, 1]);
    expect(result.movedRows).toBe(0);
  });

  it("keeps errors and blanks last for both directions", () => {
    const values = [[2], [null], [{ type: "#REF!" }], [1], ["a"], [true]];
    expect([...plan(values, [{ columnOffset: 0, direction: "asc" }]).destinationToSource]).toEqual([3, 0, 4, 5, 2, 1]);
    expect([...plan(values, [{ columnOffset: 0, direction: "desc" }]).destinationToSource]).toEqual([5, 4, 0, 3, 2, 1]);
  });

  it("uses row-span rather than rectangle overlap for filter conflicts", () => {
    expect(rowSpansIntersect({ startRow: 5, endRow: 8 }, { startRow: 8, endRow: 10 })).toBe(true);
    expect(conflictsWithFilter(
      { startRow: 5, endRow: 8, startCol: 10, endCol: 12 },
      { range: { startRow: 8, endRow: 10, startCol: 0, endCol: 3 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "isBlank" }] },
    )).toBe(true);
    expect(conflictsWithFilter({ startRow: 0, endRow: 4, startCol: 10, endCol: 12 }, null)).toBe(false);
    expect(conflictsWithFilter(
      { startRow: 0, endRow: 4, startCol: 10, endCol: 12 },
      { range: { startRow: 5, endRow: 10, startCol: 0, endCol: 3 }, hasHeader: false, conditions: [{ columnOffset: 0, operator: "isBlank" }] },
    )).toBe(false);
  });

  it("has an empty body when a header-only range is sorted", () => {
    const result = plan([["Header"]], [{ columnOffset: 0, direction: "asc" }], true);
    expect(result.bodyStartRow).toBe(1);
    expect(result.bodyRowCount).toBe(0);
    expect(result.movedRows).toBe(0);
  });
});
