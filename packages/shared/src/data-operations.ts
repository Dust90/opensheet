// @injoysai/opensheet-shared — data operation contracts (M4: sort / filter / find / dedupe)
//
// Frozen base semantics (see docs/m4-data-operations.md):
// - Blank: `null` is the ONLY true blank. "" is an ordinary string. null !== "".
// - Types: number 1 !== string "1"; booleans never dedupe-equal numbers;
//   CellError compares by (type + message).
// - Conflict: sort/dedupe whose row span overlaps an active filter range are
//   REJECTED — hidden rows must never be mutated invisibly (MVP rule).
//
// M4.0.1: validators accept `unknown` (SDK / plugin / Snapshot JSON input is
// not protected by TypeScript) and narrow via assertion signatures. Enum
// types derive from constant tuples so types and runtime checks never drift.

import type { CellPrimitive } from "./cell.js";
import { SheetError } from "./errors.js";
import type { Range } from "./range.js";
import { rangeWidth } from "./range.js";

// ── Sort ────────────────────────────────────────────────────────────────────

export const SORT_DIRECTIONS = ["asc", "desc"] as const;
export type SortDirection = (typeof SORT_DIRECTIONS)[number];

export interface SortKey {
  /** 0-based offset from `range.startCol` (NOT an absolute column). */
  columnOffset: number;
  direction: SortDirection;
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

export const FILTER_OPERATORS = [
  "equals",
  "notEquals",
  "contains",
  "greaterThan",
  "lessThan",
  "isBlank",
  "notBlank",
] as const;
export type FilterOperator = (typeof FILTER_OPERATORS)[number];

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

export const FIND_SEARCH_IN = ["values", "formulas"] as const;
export type FindSearchIn = (typeof FIND_SEARCH_IN)[number];

export const FIND_SCOPES = ["visible", "all"] as const;
export type FindScope = (typeof FIND_SCOPES)[number];

export const FIND_DIRECTIONS = ["forward", "backward"] as const;
export type FindDirection = (typeof FIND_DIRECTIONS)[number];

export interface FindOptions {
  query: string;
  matchCase: boolean;
  /** Whole-cell match; otherwise substring ("contains"). */
  wholeCell: boolean;
  /** Search computed values or formula sources. */
  searchIn: FindSearchIn;
  /** "visible" skips rows hidden by an active filter; "all" scans every physical row. */
  scope: FindScope;
  direction: FindDirection;
}

// ── Validation ──────────────────────────────────────────────────────────────

function fail(message: string): never {
  throw new SheetError("E_VALIDATION", message);
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function requireEnum(value: unknown, options: readonly string[], what: string): void {
  if (typeof value !== "string" || !options.includes(value)) {
    fail(`${what}: must be one of ${options.map((o) => `"${o}"`).join("/")} (got ${JSON.stringify(value)})`);
  }
}

function requireBoolean(value: unknown, what: string): void {
  if (typeof value !== "boolean") {
    fail(`${what}: must be a boolean (got ${JSON.stringify(value)})`);
  }
}

function validateRangeValue(value: unknown, what: string): Range {
  if (!isPlainObject(value)) {
    fail(`${what}: range must be an object (got ${JSON.stringify(value)})`);
  }
  const { startRow, startCol, endRow, endCol } = value;
  for (const [name, field] of Object.entries({ startRow, startCol, endRow, endCol })) {
    if (typeof field !== "number" || !Number.isSafeInteger(field) || field < 0) {
      fail(`${what}: range.${name} must be a non-negative safe integer (got ${JSON.stringify(field)})`);
    }
  }
  const range = {
    startRow: startRow as number,
    startCol: startCol as number,
    endRow: endRow as number,
    endCol: endCol as number,
  };
  if (range.startRow > range.endRow || range.startCol > range.endCol) {
    fail(`${what}: range must be normalized (start <= end), got ${JSON.stringify(range)}`);
  }
  return range;
}

function validateColumnOffset(value: unknown, width: number, what: string): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0 || value >= width) {
    fail(`${what}: columnOffset ${JSON.stringify(value)} outside range width ${width}`);
  }
  return value;
}

/** CellPrimitive plus the M4 finiteness rule: NaN/Infinity are not valid comparison values. */
function isValidPrimitiveValue(value: unknown): value is CellPrimitive {
  if (value === null) return true;
  const type = typeof value;
  if (type === "string" || type === "boolean") return true;
  if (type === "number") return Number.isFinite(value as number);
  return false;
}

export function validateSortSpec(value: unknown): asserts value is SortSpec {
  const what = "SortSpec";
  if (!isPlainObject(value)) fail(`${what}: must be an object (got ${JSON.stringify(value)})`);
  const range = validateRangeValue(value.range, what);
  requireBoolean(value.hasHeader, `${what}.hasHeader`);
  if (value.locale !== undefined) {
    if (typeof value.locale !== "string" || value.locale.length === 0) {
      fail(`${what}.locale: must be a non-empty string when present`);
    }
    try {
      new Intl.Collator(value.locale);
    } catch {
      fail(`${what}.locale: must be a valid Intl.Collator locale`);
    }
  }
  if (!Array.isArray(value.keys) || value.keys.length === 0) {
    fail(`${what}.keys: must be a non-empty array`);
  }
  const width = rangeWidth(range);
  const seen = new Set<number>();
  for (const key of value.keys as unknown[]) {
    if (!isPlainObject(key)) fail(`${what}.keys: each key must be an object`);
    const offset = validateColumnOffset(key.columnOffset, width, `${what}.keys`);
    if (seen.has(offset)) fail(`${what}.keys: duplicate columnOffset ${offset}`);
    seen.add(offset);
    requireEnum(key.direction, SORT_DIRECTIONS, `${what}.keys.direction`);
  }
}

export function validateFilterSpec(value: unknown): asserts value is FilterSpec {
  const what = "FilterSpec";
  if (!isPlainObject(value)) fail(`${what}: must be an object (got ${JSON.stringify(value)})`);
  const range = validateRangeValue(value.range, what);
  requireBoolean(value.hasHeader, `${what}.hasHeader`);
  if (!Array.isArray(value.conditions) || value.conditions.length === 0) {
    fail(`${what}.conditions: must be a non-empty array`);
  }
  const width = rangeWidth(range);
  for (const condition of value.conditions as unknown[]) {
    if (!isPlainObject(condition)) fail(`${what}.conditions: each condition must be an object`);
    validateColumnOffset(condition.columnOffset, width, `${what}.conditions`);
    requireEnum(condition.operator, FILTER_OPERATORS, `${what}.conditions.operator`);
    if (condition.caseSensitive !== undefined) {
      requireBoolean(condition.caseSensitive, `${what}.conditions.caseSensitive`);
    }
    const needsValue = condition.operator !== "isBlank" && condition.operator !== "notBlank";
    if (needsValue && condition.value === undefined) {
      fail(`${what}.conditions: operator "${condition.operator}" requires a value`);
    }
    if (condition.value !== undefined && !isValidPrimitiveValue(condition.value)) {
      fail(
        `${what}.conditions.value: must be a finite number, string, boolean, or null (got ${JSON.stringify(condition.value)})`,
      );
    }
  }
}

export function validateDedupeSpec(value: unknown): asserts value is DedupeSpec {
  const what = "DedupeSpec";
  if (!isPlainObject(value)) fail(`${what}: must be an object (got ${JSON.stringify(value)})`);
  const range = validateRangeValue(value.range, what);
  requireBoolean(value.hasHeader, `${what}.hasHeader`);
  if (!Array.isArray(value.keyColumnOffsets)) {
    fail(`${what}.keyColumnOffsets: must be an array`);
  }
  const width = rangeWidth(range);
  const seen = new Set<number>();
  for (const raw of value.keyColumnOffsets as unknown[]) {
    const offset = validateColumnOffset(raw, width, `${what}.keyColumnOffsets`);
    if (seen.has(offset)) fail(`${what}.keyColumnOffsets: duplicate columnOffset ${offset}`);
    seen.add(offset);
  }
  if (value.keep !== "first") {
    fail(`${what}.keep: must be "first" (got ${JSON.stringify(value.keep)})`);
  }
}

export function validateFindOptions(value: unknown): asserts value is FindOptions {
  const what = "FindOptions";
  if (!isPlainObject(value)) fail(`${what}: must be an object (got ${JSON.stringify(value)})`);
  if (typeof value.query !== "string" || value.query.length === 0) {
    fail(`${what}.query: must be a non-empty string`);
  }
  requireBoolean(value.matchCase, `${what}.matchCase`);
  requireBoolean(value.wholeCell, `${what}.wholeCell`);
  requireEnum(value.searchIn, FIND_SEARCH_IN, `${what}.searchIn`);
  requireEnum(value.scope, FIND_SCOPES, `${what}.scope`);
  requireEnum(value.direction, FIND_DIRECTIONS, `${what}.direction`);
}
