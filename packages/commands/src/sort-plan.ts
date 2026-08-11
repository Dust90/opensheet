import type { WorksheetView } from "@opensheet/core";
import {
  isCellError,
  SheetError,
  validateSortSpec,
  type CellValue,
  type SortSpec,
} from "@opensheet/shared";

/** Locale used when a SortSpec does not explicitly specify one. */
export const DEFAULT_SORT_LOCALE = "en-US";

export interface SortPlan {
  readonly bodyStartRow: number;
  readonly bodyRowCount: number;
  /** Destination body offset → original source body offset. */
  readonly destinationToSource: Uint32Array;
  /** Source body offset → destination body offset. */
  readonly sourceToDestination: Uint32Array;
  readonly movedRows: number;
}

interface SortRow {
  sourceOffset: number;
  keys: CellValue[];
}

export { conflictsWithFilter, rowSpansIntersect } from "./data-operation-conflicts.js";

/** Build a deterministic, stable row permutation without mutating the sheet. */
export function buildSortPlan(sheet: WorksheetView, spec: SortSpec): SortPlan {
  validateSortSpec(spec);
  if (spec.range.endRow >= sheet.rowCount || spec.range.endCol >= sheet.columnCount) {
    throw new SheetError("E_INVALID_RANGE", "SortSpec range exceeds worksheet bounds");
  }

  const bodyStartRow = spec.range.startRow + (spec.hasHeader ? 1 : 0);
  const bodyRowCount = Math.max(0, spec.range.endRow - bodyStartRow + 1);
  const rows: SortRow[] = Array.from({ length: bodyRowCount }, (_, sourceOffset) => ({
    sourceOffset,
    keys: spec.keys.map((key) => sheet.getCell(bodyStartRow + sourceOffset, spec.range.startCol + key.columnOffset)?.value ?? null),
  }));
  const collator = new Intl.Collator(spec.locale ?? DEFAULT_SORT_LOCALE, { usage: "sort", sensitivity: "variant", numeric: false });
  rows.sort((left, right) => compareRows(left, right, spec, collator));

  const destinationToSource = new Uint32Array(bodyRowCount);
  const sourceToDestination = new Uint32Array(bodyRowCount);
  let movedRows = 0;
  rows.forEach((row, destinationOffset) => {
    destinationToSource[destinationOffset] = row.sourceOffset;
    sourceToDestination[row.sourceOffset] = destinationOffset;
    if (row.sourceOffset !== destinationOffset) movedRows += 1;
  });
  return { bodyStartRow, bodyRowCount, destinationToSource, sourceToDestination, movedRows };
}

function compareRows(left: SortRow, right: SortRow, spec: SortSpec, collator: Intl.Collator): number {
  for (let index = 0; index < spec.keys.length; index += 1) {
    const key = spec.keys[index]!;
    const compared = compareValue(left.keys[index]!, right.keys[index]!, key.direction, collator);
    if (compared !== 0) return compared;
  }
  return left.sourceOffset - right.sourceOffset;
}

function compareValue(left: CellValue, right: CellValue, direction: "asc" | "desc", collator: Intl.Collator): number {
  const leftClass = valueClass(left);
  const rightClass = valueClass(right);
  // Errors and blanks are terminal classes; direction affects normal values only.
  if (leftClass !== rightClass) return leftClass - rightClass;
  if (leftClass !== 0) return 0;

  let compared: number;
  if (typeof left === "number" && typeof right === "number") compared = left - right;
  else if (typeof left === "string" && typeof right === "string") compared = collator.compare(left, right);
  else if (typeof left === "boolean" && typeof right === "boolean") compared = Number(left) - Number(right);
  else compared = primitiveClass(left as Exclude<CellValue, null>) - primitiveClass(right as Exclude<CellValue, null>);
  return direction === "desc" ? -compared : compared;
}

function valueClass(value: CellValue): number {
  if (value === null) return 2;
  return isCellError(value) ? 1 : 0;
}

function primitiveClass(value: Exclude<CellValue, null>): number {
  if (typeof value === "number") return 0;
  if (typeof value === "string") return 1;
  return 2;
}
