// Command + journal contracts (inverse patch journal, ADR-0003).

import type { Workbook } from "@opensheet/core";
import type { ChangeKind, ChangeSource, Range } from "@opensheet/shared";

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
  readonly journal: JournalEntry;
}

export interface SheetCommand<TPayload = unknown, TResult = void> {
  readonly id: string;
  /** Throw SheetError on invalid payloads. Runs before execute. */
  validate?(payload: TPayload, ctx: CommandContext): void;
  execute(ctx: CommandContext, payload: TPayload): CommandOutcome<TResult>;
}

/** Journal batch recorded as ONE history entry. */
export interface JournalBatch {
  readonly entries: readonly JournalEntry[];
  readonly source: ChangeSource;
  readonly approxBytes: number;
}

/** Implemented by @opensheet/history; injected into the bus (no reverse dep). */
export interface HistorySink {
  push(batch: JournalBatch): void;
}

/** Runs inside the open transaction, right before commit (ADR-0003).
 *  The formula engine (M3) hooks here to fold derived recalculation results
 *  into the same merged change event. */
export type BeforeCommitHook = (ctx: { workbook: Workbook; source: ChangeSource }) => void;
