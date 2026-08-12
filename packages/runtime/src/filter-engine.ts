// M4.2-A: pure FilterSpec evaluation. This module intentionally depends only
// on WorksheetView; it neither mutates worksheet state nor knows about Grid
// projections, commands, or renderer concerns.

import type { WorksheetView } from "@injoysai/opensheet-core";
import {
  isCellError,
  SheetError,
  validateFilterSpec,
  type CellPrimitive,
  type FilterCondition,
  type FilterSpec,
} from "@injoysai/opensheet-shared";

/**
 * Whether one physical row satisfies every condition in `spec`.
 *
 * The header policy belongs to visibility evaluation: callers that evaluate a
 * complete range should use evaluateVisibleRows(), which keeps a header row
 * visible regardless of its cell values.
 */
export function rowMatchesFilter(
  sheet: WorksheetView,
  physicalRow: number,
  spec: FilterSpec,
): boolean {
  validateFilterSpec(spec);
  assertFilterRangeFitsSheet(sheet, spec);
  if (!Number.isSafeInteger(physicalRow) || physicalRow < spec.range.startRow || physicalRow > spec.range.endRow) {
    return false;
  }
  return spec.conditions.every((condition) => {
    const value = sheet.getCell(physicalRow, spec.range.startCol + condition.columnOffset)?.value ?? null;
    return matchesCondition(value, condition);
  });
}

/**
 * Physical rows inside the filter range that remain visible, in ascending
 * order. Rows outside the range deliberately do not appear: FilteredRowProjection
 * retains them through its identity segments. The result is safe to pass
 * directly to that projection without building a sheet-sized visibility map.
 */
export function evaluateVisibleRows(sheet: WorksheetView, spec: FilterSpec): Uint32Array {
  validateFilterSpec(spec);
  assertFilterRangeFitsSheet(sheet, spec);

  const rows: number[] = [];
  for (let row = spec.range.startRow; row <= spec.range.endRow; row++) {
    if ((spec.hasHeader && row === spec.range.startRow) || rowMatchesFilterUnchecked(sheet, row, spec)) {
      rows.push(row);
    }
  }
  return Uint32Array.from(rows);
}

function rowMatchesFilterUnchecked(sheet: WorksheetView, physicalRow: number, spec: FilterSpec): boolean {
  return spec.conditions.every((condition) => {
    const value = sheet.getCell(physicalRow, spec.range.startCol + condition.columnOffset)?.value ?? null;
    return matchesCondition(value, condition);
  });
}

function matchesCondition(value: import("@injoysai/opensheet-shared").CellValue, condition: FilterCondition): boolean {
  // Errors are values for formula display. They are not blank, but do not
  // match ordinary comparison predicates (including notEquals).
  if (isCellError(value)) return condition.operator === "notBlank";

  switch (condition.operator) {
    case "isBlank":
      return value === null;
    case "notBlank":
      return value !== null;
    case "equals":
      return primitivesEqual(value, condition.value!, condition.caseSensitive === true);
    case "notEquals":
      return !primitivesEqual(value, condition.value!, condition.caseSensitive === true);
    case "contains":
      return containsDisplayText(value, condition.value!, condition.caseSensitive === true);
    case "greaterThan": {
      const left = finiteNumericValue(value);
      const right = finiteNumericValue(condition.value!);
      return left !== undefined && right !== undefined && left > right;
    }
    case "lessThan": {
      const left = finiteNumericValue(value);
      const right = finiteNumericValue(condition.value!);
      return left !== undefined && right !== undefined && left < right;
    }
  }
}

function primitivesEqual(left: CellPrimitive, right: CellPrimitive, caseSensitive: boolean): boolean {
  if (typeof left !== typeof right) return false;
  if (typeof left === "string" && typeof right === "string" && !caseSensitive) {
    return left.toLowerCase() === right.toLowerCase();
  }
  return left === right;
}

function containsDisplayText(value: CellPrimitive, needle: CellPrimitive, caseSensitive: boolean): boolean {
  // A null condition value is never a text search. In particular, avoid the
  // accidental `"anything".includes("")` match from displaying null as "".
  if (value === null || needle === null) return false;
  const haystack = displayText(value);
  const query = displayText(needle);
  return caseSensitive
    ? haystack.includes(query)
    : haystack.toLowerCase().includes(query.toLowerCase());
}

function displayText(value: CellPrimitive): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

/** Finite number literals and non-empty finite numeric strings only. */
function finiteNumericValue(value: CellPrimitive): number | undefined {
  if (typeof value === "number") return Number.isFinite(value) ? value : undefined;
  if (typeof value !== "string" || value.trim() === "") return undefined;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : undefined;
}

function assertFilterRangeFitsSheet(sheet: WorksheetView, spec: FilterSpec): void {
  if (spec.range.endRow >= sheet.rowCount || spec.range.endCol >= sheet.columnCount) {
    throw new SheetError(
      "E_INVALID_RANGE",
      `FilterSpec range exceeds worksheet bounds (${sheet.rowCount} rows × ${sheet.columnCount} columns)`,
    );
  }
}
