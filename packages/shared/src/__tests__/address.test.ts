import { describe, expect, it } from "vitest";
import {
  colFromName,
  colToName,
  formatAddress,
  parseAddress,
  parseCellRef,
} from "../index.js";
import fc from "fast-check";

describe("address parsing", () => {
  it("parses simple addresses (0-based)", () => {
    expect(parseAddress("A1")).toEqual({ row: 0, col: 0 });
    expect(parseAddress("B2")).toEqual({ row: 1, col: 1 });
    expect(parseAddress("AA10")).toEqual({ row: 9, col: 26 });
    expect(parseAddress("XFD1048576")).toEqual({ row: 1048575, col: 16383 });
  });

  it("parses absolute markers", () => {
    expect(parseCellRef("$A$1")).toEqual({ row: 0, col: 0, rowAbs: true, colAbs: true });
    expect(parseCellRef("$A1")).toEqual({ row: 0, col: 0, rowAbs: false, colAbs: true });
    expect(parseCellRef("A$1")).toEqual({ row: 0, col: 0, rowAbs: true, colAbs: false });
  });

  it("is case-insensitive", () => {
    expect(parseAddress("a1")).toEqual({ row: 0, col: 0 });
    expect(parseAddress("aa10")).toEqual({ row: 9, col: 26 });
  });

  it("formats addresses", () => {
    expect(formatAddress({ row: 0, col: 0 })).toBe("A1");
    expect(formatAddress({ row: 9, col: 26 })).toBe("AA10");
    expect(formatAddress({ row: 0, col: 0 }, { rowAbs: true, colAbs: true })).toBe("$A$1");
  });

  it("column name roundtrip", () => {
    expect(colToName(0)).toBe("A");
    expect(colToName(25)).toBe("Z");
    expect(colToName(26)).toBe("AA");
    expect(colToName(16383)).toBe("XFD");
    expect(colFromName("A")).toBe(0);
    expect(colFromName("AA")).toBe(26);
    expect(colFromName("XFD")).toBe(16383);
  });

  it("rejects invalid input", () => {
    expect(() => parseAddress("")).toThrow();
    expect(() => parseAddress("1A")).toThrow();
    expect(() => parseAddress("A0")).toThrow();
    expect(() => parseAddress("A1048577")).toThrow();
    expect(() => parseAddress("XFE1")).toThrow();
  });

  it("property: format(parse(x)) roundtrip", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: 1048575 }),
        fc.integer({ min: 0, max: 16383 }),
        fc.boolean(),
        fc.boolean(),
        (row, col, rowAbs, colAbs) => {
          const text = formatAddress({ row, col }, { rowAbs, colAbs });
          const ref = parseCellRef(text);
          return (
            ref.row === row && ref.col === col && ref.rowAbs === rowAbs && ref.colAbs === colAbs
          );
        },
      ),
    );
  });
});
