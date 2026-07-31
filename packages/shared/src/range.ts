// @opensheet/shared — rectangular cell ranges

import { formatAddress, parseCellRef } from "./address.js";
import { SheetError } from "./errors.js";

/** Inclusive, normalized rectangle of 0-based coordinates (start <= end). */
export interface Range {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/** Normalize any corner ordering into start <= end. */
export function normalizeRange(range: Range): Range {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    startCol: Math.min(range.startCol, range.endCol),
    endRow: Math.max(range.startRow, range.endRow),
    endCol: Math.max(range.startCol, range.endCol),
  };
}

/** Parse "A1" or "A1:B2" (any corner order) into a normalized Range. */
export function parseRange(text: string): Range {
  const parts = text.trim().split(":");
  if (parts.length < 1 || parts.length > 2 || parts.some((p) => p.length === 0)) {
    throw new SheetError("E_INVALID_RANGE", `Invalid range: "${text}"`);
  }
  const start = parseCellRef(parts[0]!);
  const end = parts.length === 2 ? parseCellRef(parts[1]!) : start;
  return normalizeRange({
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  });
}

export function formatRange(range: Range): string {
  const normalized = normalizeRange(range);
  const start = formatAddress({ row: normalized.startRow, col: normalized.startCol });
  const end = formatAddress({ row: normalized.endRow, col: normalized.endCol });
  return start === end ? start : `${start}:${end}`;
}

export function rangeHeight(range: Range): number {
  return range.endRow - range.startRow + 1;
}

export function rangeWidth(range: Range): number {
  return range.endCol - range.startCol + 1;
}

export function rangeCellCount(range: Range): number {
  return rangeHeight(range) * rangeWidth(range);
}

export function rangeContainsCell(range: Range, row: number, col: number): boolean {
  return (
    row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
  );
}

export function rangesIntersect(a: Range, b: Range): boolean {
  return (
    a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol
  );
}

export function rangesEqual(a: Range, b: Range): boolean {
  return (
    a.startRow === b.startRow &&
    a.startCol === b.startCol &&
    a.endRow === b.endRow &&
    a.endCol === b.endCol
  );
}

/** Clamp a range into the given sheet bounds. Returns null if fully outside. */
export function clampRange(range: Range, rowCount: number, columnCount: number): Range | null {
  const clamped: Range = {
    startRow: Math.max(0, range.startRow),
    startCol: Math.max(0, range.startCol),
    endRow: Math.min(rowCount - 1, range.endRow),
    endCol: Math.min(columnCount - 1, range.endCol),
  };
  if (clamped.startRow > clamped.endRow || clamped.startCol > clamped.endCol) return null;
  return clamped;
}
