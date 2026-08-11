import { Worksheet } from "@opensheet/core";
import { describe, expect, it } from "vitest";
import { buildDedupePlan, encodeDedupeValue } from "../dedupe-plan.js";

function sheet(values: unknown[][]): Worksheet {
  const result = new Worksheet({ id: "sheet", name: "Sheet", rowCount: values.length, columnCount: values[0]?.length ?? 1 });
  values.forEach((row, r) => row.forEach((value, c) => {
    if (value !== undefined) result.setCell(r, c, { value: value as never });
  }));
  return result;
}

function plan(values: unknown[][], keyColumnOffsets: number[] = [], hasHeader = false) {
  return buildDedupePlan(sheet(values), {
    range: { startRow: 0, startCol: 0, endRow: values.length - 1, endCol: values[0]!.length - 1 },
    hasHeader,
    keyColumnOffsets,
    keep: "first",
  });
}

describe("buildDedupePlan", () => {
  it("keeps the first matching row in stable source order", () => {
    const result = plan([["a"], ["b"], ["a"], ["b"], ["c"]]);
    expect([...result.keptSourceOffsets]).toEqual([0, 1, 4]);
    expect([...result.removedSourceOffsets]).toEqual([2, 3]);
    expect(result.keptRowCount).toBe(3);
    expect(result.removedRows).toBe(2);
  });

  it("uses typed key encoding and compares CellError by type plus message", () => {
    const error = { type: "#REF!" as const, message: "bad reference" };
    const result = plan([[1], ["1"], [true], [null], [""], [error], [error], [{ type: "#REF!", message: "other" }]]);
    expect([...result.keptSourceOffsets]).toEqual([0, 1, 2, 3, 4, 5, 7]);
    expect([...result.removedSourceOffsets]).toEqual([6]);
    expect(new Set([encodeDedupeValue(1), encodeDedupeValue("1"), encodeDedupeValue(true), encodeDedupeValue(null), encodeDedupeValue("")]).size).toBe(5);
  });

  it("uses all range columns by default and respects explicit key columns", () => {
    const values = [["a", 1], ["a", 2], ["a", 1]];
    expect([...plan(values).keptSourceOffsets]).toEqual([0, 1]);
    expect([...plan(values, [0]).keptSourceOffsets]).toEqual([0]);
  });

  it("uses computed formula values rather than formula source as keys", () => {
    const source = sheet([["left", 10], ["right", 10]]);
    source.setCell(0, 1, { value: 10, formula: "=A1" });
    source.setCell(1, 1, { value: 10, formula: "=A2*2" });
    const result = buildDedupePlan(source, {
      range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, hasHeader: false, keyColumnOffsets: [1], keep: "first",
    });
    expect([...result.keptSourceOffsets]).toEqual([0]);
    expect([...result.removedSourceOffsets]).toEqual([1]);
  });

  it("keeps a header fixed and treats a header-only range as an empty body", () => {
    const withHeader = plan([["Name"], ["a"], ["a"], ["b"]], [], true);
    expect(withHeader.bodyStartRow).toBe(1);
    expect([...withHeader.keptSourceOffsets]).toEqual([0, 2]);
    expect([...withHeader.removedSourceOffsets]).toEqual([1]);

    const headerOnly = plan([["Name"]], [], true);
    expect(headerOnly.bodyRowCount).toBe(0);
    expect(headerOnly.removedRows).toBe(0);
  });

  it("treats absent sparse cells as null blanks and rejects out-of-bounds ranges", () => {
    const sparse = sheet([[undefined], [null], ["x"]]);
    const result = buildDedupePlan(sparse, {
      range: { startRow: 0, startCol: 0, endRow: 2, endCol: 0 }, hasHeader: false, keyColumnOffsets: [], keep: "first",
    });
    expect([...result.keptSourceOffsets]).toEqual([0, 2]);
    expect([...result.removedSourceOffsets]).toEqual([1]);
    expect(() => buildDedupePlan(sparse, {
      range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 }, hasHeader: false, keyColumnOffsets: [], keep: "first",
    })).toThrow(/exceeds worksheet bounds/);
  });
});
