import { describe, expect, it } from "vitest";
import {
  clampRange,
  formatRange,
  parseRange,
  rangeCellCount,
  rangeContainsCell,
  rangesIntersect,
} from "../index.js";

describe("range parsing", () => {
  it("parses single cells and rectangles", () => {
    expect(parseRange("A1")).toEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    expect(parseRange("A1:B2")).toEqual({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  });

  it("normalizes reversed corners", () => {
    expect(parseRange("C3:A1")).toEqual({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 });
    expect(parseRange("B5:B2")).toEqual({ startRow: 1, startCol: 1, endRow: 4, endCol: 1 });
  });

  it("accepts absolute markers", () => {
    expect(parseRange("$A$1:$B$2")).toEqual({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  });

  it("formats ranges", () => {
    expect(formatRange({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 })).toBe("A1:B2");
    expect(formatRange({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 })).toBe("A1");
  });

  it("rejects malformed ranges", () => {
    expect(() => parseRange("A1:")).toThrow();
    expect(() => parseRange("A1:B2:C3")).toThrow();
    expect(() => parseRange("")).toThrow();
  });

  it("geometry helpers", () => {
    const r = parseRange("B2:D4");
    expect(rangeCellCount(r)).toBe(9);
    expect(rangeContainsCell(r, 1, 1)).toBe(true);
    expect(rangeContainsCell(r, 0, 0)).toBe(false);
    expect(rangesIntersect(r, parseRange("C1:C10"))).toBe(true);
    expect(rangesIntersect(r, parseRange("E1:E10"))).toBe(false);
  });

  it("clamps to sheet bounds", () => {
    expect(clampRange(parseRange("A1:D10"), 5, 2)).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 4,
      endCol: 1,
    });
    expect(clampRange({ startRow: 10, startCol: 0, endRow: 12, endCol: 1 }, 5, 5)).toBeNull();
  });
});
