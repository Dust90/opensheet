// filter.apply / filter.clear: Worksheet filter state with complete inverse
// journals. Filtering affects visibility only, never formula values.

import { SheetError, validateFilterSpec, type FilterSpec, type Range } from "@injoysai/opensheet-shared";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

export interface FilterApplyPayload {
  spec: FilterSpec;
}

export const filterApplyCommand: SheetCommand<FilterApplyPayload> = {
  id: "filter.apply",
  validate(payload, ctx) {
    if (typeof payload !== "object" || payload === null || !("spec" in payload)) {
      throw new SheetError("E_VALIDATION", "filter.apply requires a FilterSpec");
    }
    validateFilterSpec(payload.spec);
    assertFilterFitsSheet(payload.spec, ctx.workbook.getSheet(ctx.sheetId));
  },
  execute(ctx, payload): CommandOutcome {
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    const previous = cloneOrNull(sheet.filter);
    // Detach now: neither the caller nor a later mutation of the payload can
    // corrupt redo's filter state.
    const next = cloneFilterSpec(payload.spec);
    const range = filterAffectedRange(previous, next);
    const apply = (filter: FilterSpec | null, source: typeof ctx.source) => {
      sheet.setFilter(filter);
      emitFilterChange(ctx.workbook, sheet.id, range, source);
    };
    apply(next, ctx.source);
    return {
      result: undefined,
      journal: makeFilterJournal({ label: "filter.apply", sheetId: sheet.id, range, previous, next, apply }),
    };
  },
};

export const filterClearCommand: SheetCommand<unknown> = {
  id: "filter.clear",
  execute(ctx): CommandOutcome {
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    const previous = cloneOrNull(sheet.filter);
    // A clear without an active filter is deliberately invisible to events
    // and HistoryManager; callers can safely issue it idempotently.
    if (previous === null) return { result: undefined, journal: null };
    const range = { ...previous.range };
    const apply = (filter: FilterSpec | null, source: typeof ctx.source) => {
      sheet.setFilter(filter);
      emitFilterChange(ctx.workbook, sheet.id, range, source);
    };
    apply(null, ctx.source);
    return {
      result: undefined,
      journal: makeFilterJournal({ label: "filter.clear", sheetId: sheet.id, range, previous, next: null, apply }),
    };
  },
};

function assertFilterFitsSheet(spec: FilterSpec, sheet: { rowCount: number; columnCount: number }): void {
  if (spec.range.endRow >= sheet.rowCount || spec.range.endCol >= sheet.columnCount) {
    throw new SheetError(
      "E_INVALID_RANGE",
      `FilterSpec range exceeds worksheet bounds (${sheet.rowCount} rows × ${sheet.columnCount} columns)`,
    );
  }
}

function emitFilterChange(
  workbook: import("@injoysai/opensheet-core").Workbook,
  sheetId: string,
  range: Range,
  source: import("@injoysai/opensheet-shared").ChangeSource,
): void {
  workbook.emit({
    workbookId: workbook.id,
    sheetId,
    changes: [{ range, kind: "filter" }],
    source,
    batch: false,
  });
}

function makeFilterJournal(init: {
  label: string;
  sheetId: string;
  range: Range;
  previous: FilterSpec | null;
  next: FilterSpec | null;
  apply: (filter: FilterSpec | null, source: import("@injoysai/opensheet-shared").ChangeSource) => void;
}): JournalEntry {
  const { label, sheetId, range, previous, next, apply } = init;
  return {
    label,
    affected: [{ sheetId, range, kind: "filter" }],
    approxBytes: 256 + JSON.stringify(previous).length + JSON.stringify(next).length,
    undo: (ctx) => apply(previous, ctx.source),
    redo: (ctx) => apply(next, ctx.source),
  };
}

function cloneOrNull(spec: Readonly<FilterSpec> | null): FilterSpec | null {
  return spec === null ? null : cloneFilterSpec(spec);
}

function cloneFilterSpec(spec: Readonly<FilterSpec>): FilterSpec {
  return {
    range: { ...spec.range },
    hasHeader: spec.hasHeader,
    conditions: spec.conditions.map((condition) => ({ ...condition })),
  };
}

/** Every visibility transition must cover both the former and next ranges. */
function filterAffectedRange(previous: FilterSpec | null, next: FilterSpec | null): Range {
  if (previous === null && next === null) {
    throw new SheetError("E_OP_FAILED", "Filter change requires a previous or next FilterSpec");
  }
  if (previous === null) return { ...next!.range };
  if (next === null) return { ...previous.range };
  return {
    startRow: Math.min(previous.range.startRow, next.range.startRow),
    startCol: Math.min(previous.range.startCol, next.range.startCol),
    endRow: Math.max(previous.range.endRow, next.range.endRow),
    endCol: Math.max(previous.range.endCol, next.range.endCol),
  };
}
