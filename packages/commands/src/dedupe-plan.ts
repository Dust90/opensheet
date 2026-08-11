import type { WorksheetView } from "@opensheet/core";
import {
  isCellError,
  SheetError,
  validateDedupeSpec,
  type CellValue,
  type DedupeSpec,
} from "@opensheet/shared";

/** A non-mutating description of stable row compaction for `range.dedupe`. */
export interface DedupePlan {
  /** First physical row participating in duplicate detection. */
  readonly bodyStartRow: number;
  /** Number of physical rows participating (excludes a header when present). */
  readonly bodyRowCount: number;
  /** Source body offsets retained in their original, stable order. */
  readonly keptSourceOffsets: Uint32Array;
  /** Source body offsets removed by deduplication, in original order. */
  readonly removedSourceOffsets: Uint32Array;
  readonly keptRowCount: number;
  readonly removedRows: number;
}

/**
 * Encode a cell value with its type so keys never accidentally coerce:
 * `1`, `"1"`, `true`, `null`, and `""` are all distinct.  CellError equality
 * is intentionally based on both its type and optional message.
 */
export function encodeDedupeValue(value: CellValue): string {
  if (value === null) return "null";
  if (isCellError(value)) return `error:${JSON.stringify([value.type, value.message ?? null])}`;
  switch (typeof value) {
    case "string": return `string:${JSON.stringify(value)}`;
    case "number": return `number:${String(value)}`;
    case "boolean": return `boolean:${value ? "true" : "false"}`;
  }
}

/** Build a stable, typed deduplication plan without mutating the worksheet. */
export function buildDedupePlan(sheet: WorksheetView, spec: DedupeSpec): DedupePlan {
  validateDedupeSpec(spec);
  if (spec.range.endRow >= sheet.rowCount || spec.range.endCol >= sheet.columnCount) {
    throw new SheetError("E_INVALID_RANGE", "DedupeSpec range exceeds worksheet bounds");
  }

  const bodyStartRow = spec.range.startRow + (spec.hasHeader ? 1 : 0);
  const bodyRowCount = Math.max(0, spec.range.endRow - bodyStartRow + 1);
  const keyOffsets = spec.keyColumnOffsets.length === 0
    ? Array.from({ length: spec.range.endCol - spec.range.startCol + 1 }, (_, offset) => offset)
    : spec.keyColumnOffsets;
  const seen = new Set<string>();
  const kept: number[] = [];
  const removed: number[] = [];

  for (let sourceOffset = 0; sourceOffset < bodyRowCount; sourceOffset += 1) {
    const row = bodyStartRow + sourceOffset;
    const key = JSON.stringify(keyOffsets.map((offset) =>
      encodeDedupeValue(sheet.getCell(row, spec.range.startCol + offset)?.value ?? null),
    ));
    if (seen.has(key)) removed.push(sourceOffset);
    else {
      seen.add(key);
      kept.push(sourceOffset);
    }
  }

  return {
    bodyStartRow,
    bodyRowCount,
    keptSourceOffsets: Uint32Array.from(kept),
    removedSourceOffsets: Uint32Array.from(removed),
    keptRowCount: kept.length,
    removedRows: removed.length,
  };
}
