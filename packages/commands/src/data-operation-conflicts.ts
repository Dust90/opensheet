import type { FilterSpec, Range } from "@opensheet/shared";

/** True when an operation would touch physical rows hidden by a filter. */
export function rowSpansIntersect(
  left: Pick<Range, "startRow" | "endRow">,
  right: Pick<Range, "startRow" | "endRow">,
): boolean {
  return left.startRow <= right.endRow && left.endRow >= right.startRow;
}

export function conflictsWithFilter(range: Range, filter: Readonly<FilterSpec> | null): boolean {
  return filter !== null && rowSpansIntersect(range, filter.range);
}
