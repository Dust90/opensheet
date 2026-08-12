// Command + journal contracts (inverse patch journal, ADR-0003).

import type { Workbook, WorkbookView } from "@injoysai/opensheet-core";
import type { CellValue, ChangeKind, ChangeSource, Range } from "@injoysai/opensheet-shared";

/** Minimal context available while undoing/redoing journal entries. */
export interface JournalReplayContext {
  workbook: Workbook;
  source: ChangeSource;
}

/**
 * One reversible patch produced by executing a command. Closures capture
 * exactly the inverse data (old cell values, old sizes, ...); commands that
 * cannot build an inverse patch must instead deep-copy only the affected
 * ranges into the closure — never the whole sheet.
 */
export interface JournalEntry {
  readonly label: string;
  undo(ctx: JournalReplayContext): void;
  redo(ctx: JournalReplayContext): void;
  /** Ranges touched by this patch (for dirty-region tracking). */
  readonly affected: ReadonlyArray<{ sheetId: string; range: Range; kind: ChangeKind }>;
  /** Rough retained-memory estimate, used by history memory limits. */
  readonly approxBytes: number;
}

export interface CommandContext {
  readonly workbook: Workbook;
  /** Default sheet for operations that do not name one explicitly. */
  readonly sheetId: string;
  readonly source: ChangeSource;
}

export interface CommandOutcome<TResult = void> {
  readonly result: TResult;
  /** Null marks a semantic no-op: no event and no history entry. */
  readonly journal: JournalEntry | null;
}

export interface SheetCommand<TPayload = unknown, TResult = void> {
  readonly id: string;
  /** Throw SheetError on invalid payloads. Runs before execute. */
  validate?(payload: TPayload, ctx: CommandContext): void;
  /** Flush derived hooks from preceding commands before this command reads cached values. */
  readonly requiresFreshDerivedState?: boolean;
  execute(ctx: CommandContext, payload: TPayload): CommandOutcome<TResult>;
}

/** Journal batch recorded as ONE history entry. */
export interface JournalBatch {
  readonly entries: readonly JournalEntry[];
  readonly source: ChangeSource;
  readonly approxBytes: number;
}

/** Implemented by @injoysai/opensheet-history; injected into the bus (no reverse dep). */
export interface HistorySink {
  push(batch: JournalBatch): void;
}

/** A range this transaction is about to touch (drives M3 incremental recalc). */
export interface PendingChange {
  readonly sheetId: string;
  readonly range: Range;
  readonly kind: ChangeKind;
}

/**
 * Transaction-scoped writer for derived (recalculated) values (M3.0).
 * Hooks may ONLY write computed VALUES — metadata (formula, styleId,
 * numberFormat) is preserved automatically, so a sloppy engine can never
 * wipe formatting. Every write: captures the previous value, applies the
 * change, emits a "derived" event into the transaction buffer, AND appends
 * an inverse patch to the current transaction's rollback journal. Derived
 * patches are used for rollback only — they never enter user Undo history.
 */
export interface DerivedWriter {
  setComputedValue(sheetId: string, row: number, col: number, value: CellValue): void;
}

/** Runs inside the open transaction, right before commit (ADR-0003).
 *  The formula engine (M3) hooks here to fold derived recalculation results
 *  into the same merged change event.
 *  - `workbook`: READ-ONLY WorkbookView (M3 guardrail) — hooks cannot mutate.
 *  - `changes`: the ranges/kinds this transaction is about to commit, derived
 *    from the journal's affected ranges. Lets the engine recompute ONLY the
 *    dirty subgraph instead of scanning every formula.
 *  - `derived`: the ONLY write path; journaled for rollback.
 */
export type BeforeCommitHook = (ctx: {
  workbook: WorkbookView;
  source: ChangeSource;
  changes: readonly PendingChange[];
  derived: DerivedWriter;
}) => void;
