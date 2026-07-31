// RowProjection (M4.1): visual row ↔ physical row mapping.
//
// COORDINATE DISCIPLINE (docs/m4-data-operations.md):
//   - SelectionModel, Worksheet data, ChangeEvent ranges are ALWAYS physical.
//   - Viewport quadrants, the row AxisMetrics, and canvas Y are ALWAYS visual.
//   - The projection is the ONLY conversion point; never mix the two index
//     spaces in one variable. Name them visualRow / physicalRow explicitly.
//
// Pure module: no DOM, no Worksheet access — unit-testable in node.

import type { Range } from "@opensheet/shared";
import { SheetError } from "@opensheet/shared";

export interface RowProjection {
  readonly physicalRowCount: number;
  readonly visualRowCount: number;

  /** Visual row → physical row. Input is clamped into [0, visualRowCount-1]. */
  visualToPhysical(visualRow: number): number;

  /** Physical row → visual row, or undefined when the row is hidden. */
  physicalToVisual(physicalRow: number): number | undefined;

  isVisible(physicalRow: number): boolean;

  /**
   * Nearest visible physical row after `physicalRow` in `direction`
   * (itself excluded, even when visible). undefined at the sheet edge.
   */
  nextVisible(physicalRow: number, direction: 1 | -1): number | undefined;

  /**
   * Number of VISIBLE physical rows with index < `physicalRowExclusive`.
   * Used to translate a physical frozen-row count into a visual one.
   */
  visibleCountBefore(physicalRowExclusive: number): number;
}

/** No filtering: visual row === physical row. */
export class IdentityRowProjection implements RowProjection {
  readonly physicalRowCount: number;
  readonly visualRowCount: number;

  constructor(physicalRowCount: number) {
    this.physicalRowCount = physicalRowCount;
    this.visualRowCount = physicalRowCount;
  }

  visualToPhysical(visualRow: number): number {
    return Math.min(Math.max(0, visualRow), this.visualRowCount - 1);
  }

  physicalToVisual(physicalRow: number): number | undefined {
    return physicalRow >= 0 && physicalRow < this.physicalRowCount ? physicalRow : undefined;
  }

  isVisible(physicalRow: number): boolean {
    return physicalRow >= 0 && physicalRow < this.physicalRowCount;
  }

  nextVisible(physicalRow: number, direction: 1 | -1): number | undefined {
    const next = physicalRow + direction;
    return next >= 0 && next < this.physicalRowCount ? next : undefined;
  }

  visibleCountBefore(physicalRowExclusive: number): number {
    return Math.min(Math.max(0, physicalRowExclusive), this.physicalRowCount);
  }
}

/**
 * Projection over a sheet with one filtered range. Memory is O(visible rows
 * INSIDE the filter range) — never O(sheet rows) — so a 1M-row sheet with a
 * small filtered block stays cheap.
 *
 * Mapping strategy:
 *   before the filter range : identity
 *   inside the filter range : lookup in `visiblePhysicalRows` (sorted)
 *   after the filter range  : shifted by hiddenCount
 *
 * Costs: visualToPhysical O(1); physicalToVisual / nextVisible O(log v)
 * where v = visible rows inside the filter range.
 */
export class FilteredRowProjection implements RowProjection {
  readonly physicalRowCount: number;
  readonly visualRowCount: number;
  private readonly startRow: number;
  private readonly endRow: number;
  private readonly hiddenCount: number;
  private readonly visible: Uint32Array;

  /**
   * @param filterRange Inclusive physical row bounds the filter applies to.
   * @param visiblePhysicalRows Sorted-ascending physical rows inside
   *   `filterRange` that remain visible (the filter header row is included
   *   here by the caller when `hasHeader`).
   */
  constructor(
    physicalRowCount: number,
    filterRange: { startRow: number; endRow: number },
    visiblePhysicalRows: ArrayLike<number>,
  ) {
    this.physicalRowCount = physicalRowCount;
    const { startRow, endRow } = filterRange;
    if (
      !Number.isSafeInteger(startRow) ||
      !Number.isSafeInteger(endRow) ||
      startRow < 0 ||
      startRow > endRow ||
      endRow >= physicalRowCount
    ) {
      throw new SheetError(
        "E_INVALID_RANGE",
        `FilteredRowProjection: bad filter range ${startRow}..${endRow} for ${physicalRowCount} rows`,
      );
    }
    const visible = new Uint32Array(visiblePhysicalRows.length);
    let previous = -1;
    for (let i = 0; i < visiblePhysicalRows.length; i++) {
      const row = visiblePhysicalRows[i]!;
      if (!Number.isSafeInteger(row) || row < startRow || row > endRow) {
        throw new SheetError(
          "E_INVALID_RANGE",
          `FilteredRowProjection: visible row ${row} outside filter range ${startRow}..${endRow}`,
        );
      }
      if (row <= previous) {
        throw new SheetError(
          "E_VALIDATION",
          `FilteredRowProjection: visible rows must be sorted and unique (row ${row})`,
        );
      }
      visible[i] = row;
      previous = row;
    }
    this.startRow = startRow;
    this.endRow = endRow;
    this.visible = visible;
    this.hiddenCount = endRow - startRow + 1 - visible.length;
    this.visualRowCount = physicalRowCount - this.hiddenCount;
  }

  visualToPhysical(visualRow: number): number {
    const clamped = Math.min(Math.max(0, visualRow), this.visualRowCount - 1);
    if (clamped < this.startRow) return clamped;
    const offset = clamped - this.startRow;
    if (offset < this.visible.length) return this.visible[offset]!;
    return clamped + this.hiddenCount;
  }

  physicalToVisual(physicalRow: number): number | undefined {
    if (physicalRow < 0 || physicalRow >= this.physicalRowCount) return undefined;
    if (physicalRow < this.startRow) return physicalRow;
    if (physicalRow > this.endRow) return physicalRow - this.hiddenCount;
    const index = binarySearch(this.visible, physicalRow);
    return index >= 0 ? this.startRow + index : undefined;
  }

  isVisible(physicalRow: number): boolean {
    return this.physicalToVisual(physicalRow) !== undefined;
  }

  nextVisible(physicalRow: number, direction: 1 | -1): number | undefined {
    const candidate = physicalRow + direction;
    if (candidate < 0 || candidate >= this.physicalRowCount) return undefined;
    if (candidate < this.startRow || candidate > this.endRow) return candidate;
    if (direction === 1) {
      const index = lowerBound(this.visible, candidate);
      if (index < this.visible.length) return this.visible[index]!;
      // Everything below inside the range is hidden → first row after it.
      return this.endRow + 1 < this.physicalRowCount ? this.endRow + 1 : undefined;
    }
    const index = upperBound(this.visible, candidate) - 1;
    if (index >= 0) return this.visible[index]!;
    return this.startRow > 0 ? this.startRow - 1 : undefined;
  }

  visibleCountBefore(physicalRowExclusive: number): number {
    const row = Math.min(Math.max(0, physicalRowExclusive), this.physicalRowCount);
    if (row <= this.startRow) return row;
    if (row > this.endRow) return row - this.hiddenCount;
    return this.startRow + lowerBound(this.visible, row);
  }
}

// ── shared composition helpers (used by SheetGrid AND tests — single source) ──

/**
 * Map a PHYSICAL range onto the visual axis: spans from the first visible row
 * ≥ startRow to the last visible row ≤ endRow. Hidden rows inside the span
 * occupy no pixels, so the visual rect stays continuous. null when no row of
 * the range is visible.
 */
export function physicalRangeToVisualRange(range: Range, projection: RowProjection): Range | null {
  const firstVisible = projection.isVisible(range.startRow)
    ? range.startRow
    : projection.nextVisible(range.startRow, 1);
  const lastVisible = projection.isVisible(range.endRow)
    ? range.endRow
    : projection.nextVisible(range.endRow, -1);
  if (firstVisible === undefined || lastVisible === undefined) return null;
  const startRow = projection.physicalToVisual(firstVisible);
  const endRow = projection.physicalToVisual(lastVisible);
  if (startRow === undefined || endRow === undefined || startRow > endRow) return null;
  return { startRow, startCol: range.startCol, endRow, endCol: range.endCol };
}

/** Physical row of the last visible row, or -1 when nothing is visible. */
export function lastVisiblePhysicalRow(projection: RowProjection): number {
  return projection.visualRowCount === 0
    ? -1
    : projection.visualToPhysical(projection.visualRowCount - 1);
}

/**
 * Hidden-active-cell policy (M4.1 Task 8): keep a visible row if possible —
 * next visible below, else above, else the first visible row on the sheet.
 * undefined only when NOTHING is visible (degenerate, callers keep the row).
 */
export function relocateToVisibleRow(
  projection: RowProjection,
  physicalRow: number,
): number | undefined {
  if (projection.isVisible(physicalRow)) return physicalRow;
  return (
    projection.nextVisible(physicalRow, 1) ??
    projection.nextVisible(physicalRow, -1) ??
    (projection.visualRowCount > 0 ? projection.visualToPhysical(0) : undefined)
  );
}

// ── sorted-array search primitives ──────────────────────────────────────────

/** Index of `value`, or a negative number when absent. */
function binarySearch(sorted: Uint32Array, value: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    const v = sorted[mid]!;
    if (v === value) return mid;
    if (v < value) lo = mid + 1;
    else hi = mid;
  }
  return -1;
}

/** First index whose value is >= `value` (insertion point). */
function lowerBound(sorted: Uint32Array, value: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (sorted[mid]! < value) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** First index whose value is > `value`. */
function upperBound(sorted: Uint32Array, value: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (sorted[mid]! <= value) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}
