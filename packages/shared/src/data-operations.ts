// @opensheet/shared — data operation contracts (M4: sort / filter / find / dedupe)
//
// Frozen base semantics (see docs/m4-data-operations.md):
// - Blank: `null` is the ONLY true blank. "" is an ordinary string. null !== "".
// - Types: number 1 !== string "1"; booleans never dedupe-equal numbers;
//   CellError compares by (type + message).
// - Conflict: sort/dedupe whose range overlaps an active filter range are
//   REJECTED — hidden rows must never be mutated invisibly (MVP rule).

import type { CellPrimitive } from "./cell.js";
import { SheetError } from "./errors.js";
import type { Range } from "./range.js";
import { rangeWidth } from "./range.js";

// ── Sort ────────────────────────────────────────────────────────────────────

export interface SortKey {
  /** 0-based offset from `range.startCol` (NOT an absolute column). */
  columnOffset: number;
  direction: "asc" | "desc";
}

export interface SortSpec {
  range: Range;
  /** When true, the first row of `range` is a header and never moves. */
  hasHeader: boolean;
  /** Multi-key order: earlier keys dominate; original row index is the final tie-breaker (stable sort). */
  keys: readonly SortKey[];
  /** Fixed collator locale for deterministic string ordering. */
  locale?: string;
}

// ── Filter ──────────────────────────────────────────────────────────────────

export type FilterOperator =
  | "equals"
  | "notEquals"
  | "contains"
  | "greaterThan"
  | "lessThan"
  | "isBlank"
  | "notBlank";

export const FILTER_OPERATORS: readonly FilterOperator[] = [
  "equals",
  "notEquals",
  "contains",
  "greaterThan",
  "lessThan",
  "isBlank",
  "notBlank",
];

export interface FilterCondition {
  /** 0-based offset from `range.startCol` (NOT an absolute column). */
  columnOffset: number;
  operator: FilterOperator;
  /** Ignored for isBlank/notBlank; required for all other operators. */
  value?: CellPrimitive;
  caseSensitive?: boolean;
}

export interface FilterSpec {
  range: Range;
  /** When true, the first row of `range` stays visible as the filter header. */
  hasHeader: boolean;
  /** MVP: conditions combine with AND only. */
  conditions: readonly FilterCondition[];
}

// ── Dedupe ──────────────────────────────────────────────────────────────────

export interface DedupeSpec {
  range: Range;
  /** When true, the first row of `range` is a header and never participates. */
  hasHeader: boolean;
  /** Key columns as offsets from `range.startCol`; empty = compare all columns. */
  keyColumnOffsets: readonly number[];
  /** MVP: stable, always keeps the first occurrence. */
  keep: "first";
}

// ── Find ────────────────────────────────────────────────────────────────────

export interface FindOptions {
  query: string;
  matchCase: boolean;
  /** Whole-cell match; otherwise substring ("contains"). */
  wholeCell: boolean;
  /** Search computed values or formula sources. */
  searchIn: "values" | "formulas";
  /** "visible" skips rows hidden by an active filter; "all" scans every physical row. */
  scope: "visible" | "all";
  direction: "forward" | "backward";
}

// ── Validation ──────────────────────────────────────────────────────────────

function validateRangeShape(range: Range, what: string): void {
  const { startRow, startCol, endRow, endCol } = range;
  if (
    !Number.isInteger(startRow) ||
    !Number.isInteger(startCol) ||
    !Number.isInteger(endRow) ||
    !Number.isInteger(endCol) ||
    startRow < 0 ||
    startCol < 0 ||
    startRow > endRow ||
    startCol > endCol
  ) {
    throw new SheetError(
      "E_INVALID_RANGE",
      `${what}: range must be normalized, non-negative integers (got ${JSON.stringify(range)})`,
    );
  }
}

function validateColumnOffset(offset: number, width: number, what: string): void {
  if (!Number.isInteger(offset) || offset < 0 || offset >= width) {
    throw new SheetError(
      "E_VALIDATION",
      `${what}: columnOffset ${offset} outside range width ${width}`,
    );
  }
}

export function validateSortSpec(spec: SortSpec): void {
  validateRangeShape(spec.range, "SortSpec");
  if (spec.keys.length === 0) {
    throw new SheetError("E_VALIDATION", "SortSpec: at least one sort key is required");
  }
  const width = rangeWidth(spec.range);
  for (const key of spec.keys) {
    validateColumnOffset(key.columnOffset, width, "SortSpec");
  }
}

export function validateFilterSpec(spec: FilterSpec): void {
  validateRangeShape(spec.range, "FilterSpec");
  if (spec.conditions.length === 0) {
    throw new SheetError("E_VALIDATION", "FilterSpec: at least one condition is required");
  }
  const width = rangeWidth(spec.range);
  for (const condition of spec.conditions) {
    validateColumnOffset(condition.columnOffset, width, "FilterSpec");
    if (!FILTER_OPERATORS.includes(condition.operator)) {
      throw new SheetError("E_VALIDATION", `FilterSpec: unknown operator "${condition.operator}"`);
    }
    const needsValue = condition.operator !== "isBlank" && condition.operator !== "notBlank";
    if (needsValue && condition.value === undefined) {
      throw new SheetError(
        "E_VALIDATION",
        `FilterSpec: operator "${condition.operator}" requires a value`,
      );
    }
  }
}

export function validateDedupeSpec(spec: DedupeSpec): void {
  validateRangeShape(spec.range, "DedupeSpec");
  const width = rangeWidth(spec.range);
  for (const offset of spec.keyColumnOffsets) {
    validateColumnOffset(offset, width, "DedupeSpec");
  }
}

export function validateFindOptions(options: FindOptions): void {
  if (options.query.length === 0) {
    throw new SheetError("E_VALIDATION", "FindOptions: query must not be empty");
  }
}
