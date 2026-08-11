// M4.0 data operation contract validation, including M4.0.1 runtime hardening:
// validators accept `unknown` (SDK / plugin / Snapshot JSON input) and must
// reject garbage with SheetError — never a raw TypeError.

import { describe, expect, it } from "vitest";
import {
  validateDedupeSpec,
  validateFilterSpec,
  validateFindOptions,
  validateSortSpec,
  FILTER_OPERATORS,
  SORT_DIRECTIONS,
  FIND_SEARCH_IN,
  FIND_SCOPES,
  FIND_DIRECTIONS,
} from "../data-operations.js";
import { parseRange } from "../range.js";
import { SheetError } from "../errors.js";

const range = parseRange("A1:D100");

function expectSheetError(fn: () => void, pattern?: RegExp): void {
  try {
    fn();
  } catch (error) {
    expect(error).toBeInstanceOf(SheetError);
    if (pattern !== undefined) expect((error as Error).message).toMatch(pattern);
    return;
  }
  throw new Error("expected SheetError, but no error was thrown");
}

describe("validateSortSpec", () => {
  it("accepts a valid multi-key spec and narrows the type", () => {
    const input: unknown = {
      range,
      hasHeader: true,
      keys: [
        { columnOffset: 2, direction: "desc" },
        { columnOffset: 0, direction: "asc" },
      ],
      locale: "en",
    };
    validateSortSpec(input);
    // After the assertion the compiler treats input as SortSpec.
    expect(input.keys[0]!.direction).toBe("desc");
  });

  it("rejects non-object input without a raw TypeError", () => {
    expectSheetError(() => validateSortSpec(null));
    expectSheetError(() => validateSortSpec("sort"));
    expectSheetError(() => validateSortSpec([{ range }]));
  });

  it("rejects null keys with SheetError, not TypeError", () => {
    expectSheetError(() => validateSortSpec({ range, hasHeader: true, keys: null }), /keys/);
  });

  it("rejects empty keys", () => {
    expectSheetError(() => validateSortSpec({ range, hasHeader: false, keys: [] }));
  });

  it("rejects a non-boolean hasHeader", () => {
    expectSheetError(
      () => validateSortSpec({ range, hasHeader: "yes", keys: [{ columnOffset: 0, direction: "asc" }] }),
      /hasHeader/,
    );
  });

  it("rejects an invalid direction enum value", () => {
    expectSheetError(
      () => validateSortSpec({ range, hasHeader: true, keys: [{ columnOffset: 0, direction: "sideways" }] }),
      /direction/,
    );
  });

  it("rejects columnOffset outside the range width or unsafe integers", () => {
    expectSheetError(
      () => validateSortSpec({ range, hasHeader: false, keys: [{ columnOffset: 4, direction: "asc" }] }),
      /columnOffset/,
    );
    expectSheetError(
      () => validateSortSpec({ range, hasHeader: false, keys: [{ columnOffset: -1, direction: "asc" }] }),
      /columnOffset/,
    );
    expectSheetError(
      () => validateSortSpec({ range, hasHeader: false, keys: [{ columnOffset: 0.5, direction: "asc" }] }),
      /columnOffset/,
    );
  });

  it("rejects duplicate columnOffset keys", () => {
    expectSheetError(
      () =>
        validateSortSpec({
          range,
          hasHeader: false,
          keys: [
            { columnOffset: 1, direction: "asc" },
            { columnOffset: 1, direction: "desc" },
          ],
        }),
      /duplicate/,
    );
  });

  it("rejects an empty locale string", () => {
    expectSheetError(
      () => validateSortSpec({ range, hasHeader: false, keys: [{ columnOffset: 0, direction: "asc" }], locale: "" }),
      /locale/,
    );
  });

  it("wraps an invalid Intl locale in SheetError", () => {
    expectSheetError(
      () => validateSortSpec({ range, hasHeader: false, keys: [{ columnOffset: 0, direction: "asc" }], locale: "@@@" }),
      /locale/,
    );
  });

  it("rejects a non-normalized or unsafe range", () => {
    expectSheetError(() =>
      validateSortSpec({
        range: { startRow: 5, startCol: 0, endRow: 2, endCol: 3 },
        hasHeader: false,
        keys: [{ columnOffset: 0, direction: "asc" }],
      }),
    );
    expectSheetError(() =>
      validateSortSpec({
        range: { startRow: 0, startCol: 0, endRow: Number.MAX_SAFE_INTEGER + 1, endCol: 3 },
        hasHeader: false,
        keys: [{ columnOffset: 0, direction: "asc" }],
      }),
    );
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
          { columnOffset: 3, operator: "contains", value: "y", caseSensitive: true },
        ],
      }),
    ).not.toThrow();
  });

  it("rejects empty conditions", () => {
    expectSheetError(() => validateFilterSpec({ range, hasHeader: false, conditions: [] }));
  });

  it("requires a value for comparison operators", () => {
    expectSheetError(
      () => validateFilterSpec({ range, hasHeader: false, conditions: [{ columnOffset: 0, operator: "equals" }] }),
      /requires a value/,
    );
  });

  it("rejects an unknown operator", () => {
    expectSheetError(
      () => validateFilterSpec({ range, hasHeader: false, conditions: [{ columnOffset: 0, operator: "regex", value: 1 }] }),
      /operator/,
    );
  });

  it("rejects a non-boolean caseSensitive", () => {
    expectSheetError(
      () =>
        validateFilterSpec({
          range,
          hasHeader: false,
          conditions: [{ columnOffset: 0, operator: "contains", value: "x", caseSensitive: "yes" }],
        }),
      /caseSensitive/,
    );
  });

  it("rejects non-finite numbers and non-primitive values", () => {
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, { nested: true }, [1], undefined].slice(0, 4)) {
      expectSheetError(
        () =>
          validateFilterSpec({
            range,
            hasHeader: false,
            conditions: [{ columnOffset: 0, operator: "equals", value: bad }],
          }),
        /value/,
      );
    }
  });

  it("accepts null as an explicit comparison value", () => {
    expect(() =>
      validateFilterSpec({
        range,
        hasHeader: false,
        conditions: [{ columnOffset: 0, operator: "equals", value: null }],
      }),
    ).not.toThrow();
  });

  it("rejects offsets outside the range width", () => {
    expectSheetError(
      () => validateFilterSpec({ range, hasHeader: false, conditions: [{ columnOffset: 99, operator: "notBlank" }] }),
      /columnOffset/,
    );
  });

  it("exports the full operator and enum lists", () => {
    expect(FILTER_OPERATORS).toHaveLength(7);
    expect(SORT_DIRECTIONS).toEqual(["asc", "desc"]);
    expect(FIND_SEARCH_IN).toEqual(["values", "formulas"]);
    expect(FIND_SCOPES).toEqual(["visible", "all"]);
    expect(FIND_DIRECTIONS).toEqual(["forward", "backward"]);
  });
});

describe("validateDedupeSpec", () => {
  it("accepts empty key columns (compare all columns)", () => {
    expect(() =>
      validateDedupeSpec({ range, hasHeader: true, keyColumnOffsets: [], keep: "first" }),
    ).not.toThrow();
  });

  it("rejects keep values other than \"first\"", () => {
    expectSheetError(
      () => validateDedupeSpec({ range, hasHeader: true, keyColumnOffsets: [], keep: "last" }),
      /keep/,
    );
  });

  it("rejects non-array keyColumnOffsets without a raw TypeError", () => {
    expectSheetError(() => validateDedupeSpec({ range, hasHeader: true, keyColumnOffsets: null, keep: "first" }));
  });

  it("rejects duplicate key offsets", () => {
    expectSheetError(
      () => validateDedupeSpec({ range, hasHeader: false, keyColumnOffsets: [0, 0], keep: "first" }),
      /duplicate/,
    );
  });

  it("rejects key offsets outside the range width", () => {
    expectSheetError(
      () => validateDedupeSpec({ range, hasHeader: false, keyColumnOffsets: [0, 4], keep: "first" }),
      /columnOffset/,
    );
  });
});

describe("validateFindOptions", () => {
  const valid = {
    query: "abc",
    matchCase: false,
    wholeCell: false,
    searchIn: "values",
    scope: "all",
    direction: "forward",
  };

  it("accepts valid options", () => {
    expect(() => validateFindOptions(valid)).not.toThrow();
  });

  it("rejects an empty query", () => {
    expectSheetError(() => validateFindOptions({ ...valid, query: "" }), /query/);
    expectSheetError(() => validateFindOptions({ ...valid, query: 42 }), /query/);
  });

  it("rejects non-boolean flags", () => {
    expectSheetError(() => validateFindOptions({ ...valid, matchCase: "yes" }), /matchCase/);
    expectSheetError(() => validateFindOptions({ ...valid, wholeCell: 1 }), /wholeCell/);
  });

  it("rejects invalid enum values", () => {
    expectSheetError(() => validateFindOptions({ ...valid, searchIn: "anything" }), /searchIn/);
    expectSheetError(() => validateFindOptions({ ...valid, scope: "somewhere" }), /scope/);
    expectSheetError(() => validateFindOptions({ ...valid, direction: "left" }), /direction/);
  });

  it("rejects non-object input without a raw TypeError", () => {
    expectSheetError(() => validateFindOptions(undefined));
    expectSheetError(() => validateFindOptions("find"));
  });
});
