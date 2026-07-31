import { describe, expect, it } from "vitest";
import { cellsToTSV, parseTSV } from "../tsv.js";
import type { CellPrimitive } from "@opensheet/shared";

describe("cellsToTSV", () => {
  it("encodes a rectangular matrix with tabs and CRLF rows", () => {
    const cells: CellPrimitive[][] = [
      ["a", 1, true],
      [null, "x y", 2.5],
    ];
    expect(cellsToTSV(cells)).toBe("a\t1\tTRUE\r\n\tx y\t2.5");
  });

  it("quotes fields containing tab/newline/quote and doubles inner quotes", () => {
    const cells: CellPrimitive[][] = [
      ['he said "hi"', "a\tb", "line1\nline2"],
    ];
    const tsv = cellsToTSV(cells);
    expect(tsv).toBe('"he said ""hi"""\t"a\tb"\t"line1\nline2"');
  });

  it("round-trips arbitrary strings losslessly (empty field → null, Excel-like)", () => {
    const tricky = ['tab\there', 'new\nline', 'quote"inside', 'both\t"and\nline'];
    const cells: CellPrimitive[][] = [tricky];
    expect(parseTSV(cellsToTSV(cells))).toEqual(cells);
    // An explicitly empty field round-trips to null (same as an empty cell).
    expect(parseTSV(cellsToTSV([["", "x"]]))).toEqual([[null, "x"]]);
  });
});

describe("parseTSV", () => {
  it("parses plain TSV with type inference", () => {
    expect(parseTSV("a\t1\tTRUE\r\n\tx\t2.5")).toEqual([
      ["a", 1, true],
      [null, "x", 2.5],
    ]);
  });

  it("handles \\n-only input and trailing blank lines", () => {
    expect(parseTSV("1\n2\n\n")).toEqual([[1], [2]]);
  });

  it("keeps interior blank rows as [null] (positions do not shift)", () => {
    expect(parseTSV("1\n\n2")).toEqual([[1], [null], [2]]);
    // Round-trip a matrix containing a null row.
    const matrix: CellPrimitive[][] = [[1], [null], [2]];
    expect(parseTSV(cellsToTSV(matrix))).toEqual(matrix);
  });

  it("pads ragged rows with null into a rectangle", () => {
    expect(parseTSV("a\tb\nc")).toEqual([
      ["a", "b"],
      ["c", null],
    ]);
    expect(parseTSV("a\tb\tc\nd")).toEqual([
      ["a", "b", "c"],
      ["d", null, null],
    ]);
  });

  it("unquotes quoted fields including escaped quotes", () => {
    expect(parseTSV('"he said ""hi"""\t"a\tb"')).toEqual([['he said "hi"', "a\tb"]]);
  });

  it("keeps embedded newlines inside quoted fields", () => {
    expect(parseTSV('"line1\nline2"\t3')).toEqual([["line1\nline2", 3]]);
  });

  it("infers numbers and booleans (leading zeros collapse, Excel-like)", () => {
    expect(parseTSV("007\t1.5e3\t-2\tFALSE\t001")).toEqual([[7, 1500, -2, false, 1]]);
  });
});
