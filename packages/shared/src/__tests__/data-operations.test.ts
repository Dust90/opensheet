// M4.0 data operation contract validation.

import { describe, expect, it } from "vitest";
import {
  validateDedupeSpec,
  validateFilterSpec,
  validateFindOptions,
  validateSortSpec,
  FILTER_OPERATORS,
} from "../data-operations.js";
import { parseRange } from "../range.js";
import { SheetError } from "../errors.js";

const range = parseRange("A1:D100");

describe("validateSortSpec", () => {
  it("accepts a valid multi-key spec", () => {
    expect(() =>
      validateSortSpec({
        range,
        hasHeader: true,
        keys: [
          { columnOffset: 2, direction: "desc" },
          { columnOffset: 0, direction: "asc" },
        ],
      }),
    ).not.toThrow();
  });

  it("rejects empty keys", () => {
    expect(() => validateSortSpec({ range, hasHeader: false, keys: [] })).toThrow(SheetError);
  });

  it("rejects columnOffset outside the range width", () => {
    expect(() =>
      validateSortSpec({ range, hasHeader: false, keys: [{ columnOffset: 4, direction: "asc" }] }),
    ).toThrow(/columnOffset/);
    expect(() =>
      validateSortSpec({ range, hasHeader: false, keys: [{ columnOffset: -1, direction: "asc" }] }),
    ).toThrow(/columnOffset/);
  });

  it("rejects a non-normalized range", () => {
    expect(() =>
      validateSortSpec({
        range: { startRow: 5, startCol: 0, endRow: 2, endCol: 3 },
        hasHeader: false,
        keys: [{ columnOffset: 0, direction: "asc" }],
      }),
    ).toThrow(/range/);
  });
});

describe("validateFilterSpec", () => {
  it("accepts valid conditions, including blank operators without value", () => {
    expect(() =>
      validateFilterSpec({
        range,
        hasHeader: true,
        conditions: [
          { columnOffset: 0, operator: "equals", value: "x" },
          { columnOffset: 1, operator: "isBlank" },
          { columnOffset: 2, operator: "greaterThan", value: 3 },
        ],
      }),
    ).not.toThrow();
  });

  it("rejects empty conditions", () => {
    expect(() => validateFilterSpec({ range, hasHeader: false, conditions: [] })).toThrow(SheetError);
  });

  it("requires a value for comparison operators", () => {
    expect(() =>
      validateFilterSpec({
        range,
        hasHeader: false,
        conditions: [{ columnOffset: 0, operator: "equals" }],
      }),
    ).toThrow(/requires a value/);
  });

  it("rejects offsets outside the range width", () => {
    expect(() =>
      validateFilterSpec({
        range,
        hasHeader: false,
        conditions: [{ columnOffset: 99, operator: "notBlank" }],
      }),
    ).toThrow(/columnOffset/);
  });

  it("exports the full operator list", () => {
    expect(FILTER_OPERATORS).toHaveLength(7);
  });
});

describe("validateDedupeSpec", () => {
  it("accepts empty key columns (compare all columns)", () => {
    expect(() =>
      validateDedupeSpec({ range, hasHeader: true, keyColumnOffsets: [], keep: "first" }),
    ).not.toThrow();
  });

  it("rejects key offsets outside the range width", () => {
    expect(() =>
      validateDedupeSpec({ range, hasHeader: false, keyColumnOffsets: [0, 4], keep: "first" }),
    ).toThrow(/columnOffset/);
  });
});

describe("validateFindOptions", () => {
  it("rejects an empty query", () => {
    expect(() =>
      validateFindOptions({
        query: "",
        matchCase: false,
        wholeCell: false,
        searchIn: "values",
        scope: "all",
        direction: "forward",
      }),
    ).toThrow(/query/);
  });
});
