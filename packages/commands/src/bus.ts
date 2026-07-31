// CommandBus: registry + transactions with inverse patch journal and
// buffered events (ADR-0003).

import type { Workbook } from "@opensheet/core";
import { isSheetError, SheetError, type CellData, type ChangeSource } from "@opensheet/shared";
import { ApplyOperationsError, type ApplyOperationsResult } from "./operations.js";
import { CommandRegistry } from "./registry.js";
import type {
  BeforeCommitHook,
  CommandContext,
  DerivedWriter,
  HistorySink,
  JournalBatch,
  JournalEntry,
  PendingChange,
} from "./types.js";

export interface OperationEnvelope {
  type: string;
  payload: unknown;
}

export interface CommandBusOptions {
  history?: HistorySink;
  registry?: CommandRegistry;
}

/**
 * One bus per workbook. All mutations flow through here; the bus owns
 * transaction semantics:
 *
 *   beginBatch → execute commands (events buffered, no recalc, no hooks)
 *   success → beforeCommit hooks (formula recalc lands here, derived events
 *             merge into the same buffer) → endBatch(true): exactly one
 *             merged event per sheet → ONE history entry
 *   failure → reverse-replay journal → endBatch(false): buffer discarded,
 *             no history, observers never saw intermediate state
 */
export class CommandBus {
  readonly registry: CommandRegistry;
  private readonly workbook: Workbook;
  private readonly history: HistorySink | undefined;
  private readonly beforeCommitHooks: BeforeCommitHook[] = [];

  constructor(workbook: Workbook, options?: CommandBusOptions) {
    this.workbook = workbook;
    this.history = options?.history;
    this.registry = options?.registry ?? new CommandRegistry();
  }

  addBeforeCommitHook(hook: BeforeCommitHook): () => void {
    this.beforeCommitHooks.push(hook);
    return () => {
      const i = this.beforeCommitHooks.indexOf(hook);
      if (i >= 0) this.beforeCommitHooks.splice(i, 1);
    };
  }

  /** Execute one command as an implicit single-command transaction. */
  execute<TResult = void>(
    type: string,
    payload: unknown,
    options: { sheetId?: string; source?: ChangeSource } = {},
  ): TResult {
    const { results } = this.runTransaction(
      [{ type, payload }],
      options.sheetId ?? this.workbook.activeSheetId,
      options.source ?? "user",
    );
    return results[0] as TResult;
  }

  /** applyOperations entry point. */
  applyOperations(init: {
    sheetId: string;
    operations: ReadonlyArray<{ type: string; [key: string]: unknown }>;
    atomic?: boolean;
    source?: ChangeSource;
  }): ApplyOperationsResult {
    const operationId = crypto.randomUUID();
    const source = init.source ?? "api";
    const ops = init.operations.map(({ type, ...payload }) => ({ type, payload }));
    const atomic = init.atomic ?? false;

    let affectedCells = 0;
    try {
      if (atomic) {
        const { affected } = this.runTransaction(ops, init.sheetId, source);
        affectedCells = affected;
      } else {
        for (let i = 0; i < ops.length; i++) {
          try {
            const { affected } = this.runTransaction([ops[i]!], init.sheetId, source);
            affectedCells += affected;
          } catch (error) {
            if (error instanceof Error) {
              (error as { __failedIndex?: number }).__failedIndex = i;
            }
            throw error;
          }
        }
      }
    } catch (error) {
      const failedIndex = (error as { __failedIndex?: number }).__failedIndex ?? 0;
      throw new ApplyOperationsError({
        operationId,
        failedOperationIndex: failedIndex,
        errorCode: isSheetError(error) ? error.code : "E_OP_FAILED",
        message: error instanceof Error ? error.message : String(error),
      });
    }
    return { operationId, status: "completed", affectedCells, warnings: [] };
  }

  /**
   * Replay a journal batch (used by history undo/redo). Runs inside the
   * same transaction boundary so observers see a single merged event.
   */
  replayJournal(batch: JournalBatch, direction: "undo" | "redo"): void {
    const source: ChangeSource = direction === "undo" ? "undo" : "redo";
    const derivedJournal: JournalEntry[] = [];
    const completed: JournalEntry[] = [];
    this.workbook.beginBatch();
    const ctx = { workbook: this.workbook, source };
    try {
      const entries = direction === "undo" ? [...batch.entries].reverse() : batch.entries;
      for (const entry of entries) {
        if (direction === "undo") entry.undo(ctx);
        else entry.redo(ctx);
        completed.push(entry);
      }
      this.runBeforeCommitHooks(source, derivedJournal, batch.entries);
      this.workbook.endBatch(true);
    } catch (error) {
      // Best-effort restore of whatever was already replayed, then discard
      // all buffered events: observers must not see a partial replay.
      const restoreCtx = { workbook: this.workbook, source };
      for (let i = derivedJournal.length - 1; i >= 0; i--) derivedJournal[i]!.undo(restoreCtx);
      if (direction === "undo") {
        for (const entry of [...completed].reverse()) entry.redo(restoreCtx);
      } else {
        for (const entry of [...completed].reverse()) entry.undo(restoreCtx);
      }
      this.workbook.endBatch(false);
      throw error;
    }
  }

  private runTransaction(
    ops: readonly OperationEnvelope[],
    sheetId: string,
    source: ChangeSource,
  ): { results: unknown[]; affected: number } {
    const journal: JournalEntry[] = [];
    // Rollback-only journal for derived (hook) writes — never enters history.
    const derivedJournal: JournalEntry[] = [];
    const results: unknown[] = [];
    let affected = 0;
    let index = -1;
    this.workbook.beginBatch();
    try {
      for (index = 0; index < ops.length; index++) {
        const op = ops[index]!;
        const command = this.registry.get(op.type);
        const ctx: CommandContext = { workbook: this.workbook, sheetId, source };
        command.validate?.(op.payload, ctx);
        const outcome = command.execute(ctx, op.payload);
        journal.push(outcome.journal);
        results.push(outcome.result);
        affected += outcome.journal.affected.reduce(
          (sum, a) => sum + (a.range.endRow - a.range.startRow + 1) * (a.range.endCol - a.range.startCol + 1),
          0,
        );
      }
      this.runBeforeCommitHooks(source, derivedJournal, journal);
      this.workbook.endBatch(true);
    } catch (error) {
      // Reverse-replay both journals inside the still-open batch; buffered
      // events from partial execution, hooks and rollback are discarded.
      const replayCtx = { workbook: this.workbook, source };
      for (let i = derivedJournal.length - 1; i >= 0; i--) {
        derivedJournal[i]!.undo(replayCtx);
      }
      for (let i = journal.length - 1; i >= 0; i--) {
        journal[i]!.undo(replayCtx);
      }
      this.workbook.endBatch(false);
      if (error instanceof Error) {
        (error as { __failedIndex?: number }).__failedIndex = Math.max(index, 0);
      }
      throw error;
    }
    if (journal.length > 0 && source !== "derived") {
      this.history?.push({
        entries: journal,
        source,
        approxBytes: journal.reduce((sum, j) => sum + j.approxBytes, 0),
      });
    }
    return { results, affected };
  }

  private runBeforeCommitHooks(
    source: ChangeSource,
    derivedJournal: JournalEntry[],
    journal: readonly JournalEntry[],
  ): void {
    if (this.beforeCommitHooks.length === 0) return;
    const derived = this.makeDerivedWriter(derivedJournal);
    // Guardrail 3 (M2.8): hooks receive a read-only WorkbookView — they can
    // read cells/styles freely but can only WRITE through DerivedWriter,
    // which journals every mutation for rollback.
    const view = this.workbook.asView();
    // M3.0: tell the engine exactly which ranges this transaction touches so
    // it can recompute the dirty subgraph instead of scanning every formula.
    const changes: PendingChange[] = [];
    for (const entry of journal) changes.push(...entry.affected);
    for (const hook of this.beforeCommitHooks) {
      hook({ workbook: view, source, changes, derived });
    }
  }

  /**
   * Derived writes (M3.0): hooks set ONLY computed values; formula, styleId
   * and numberFormat are preserved automatically. Every write captures the
   * previous value, applies the change, emits a "derived" event into the
   * transaction buffer, AND appends an inverse patch to the rollback journal.
   */
  private makeDerivedWriter(derivedJournal: JournalEntry[]): DerivedWriter {
    const workbook = this.workbook;
    return {
      setComputedValue: (sheetId, row, col, value) => {
        const sheet = workbook.getSheet(sheetId);
        if (row < 0 || row >= sheet.rowCount || col < 0 || col >= sheet.columnCount) {
          throw new SheetError("E_INVALID_RANGE", `derived write out of bounds: ${row}:${col}`);
        }
        const previous = sheet.getCell(row, col);
        const previousClone = previous === undefined ? undefined : { ...previous };
        // Preserve formula/styleId/numberFormat; only the value is replaced.
        const next: CellData = { ...(previousClone ?? {}), value };
        sheet.setCell(row, col, next);
        const range = { startRow: row, startCol: col, endRow: row, endCol: col };
        workbook.emit({
          workbookId: workbook.id,
          sheetId,
          changes: [{ range, kind: "cells" }],
          source: "derived",
          batch: false,
        });
        derivedJournal.push({
          label: "derived.write",
          affected: [{ sheetId, range, kind: "cells" }],
          approxBytes: 192,
          undo: () => {
            if (previousClone === undefined) sheet.deleteCell(row, col);
            else sheet.setCell(row, col, { ...previousClone });
          },
          redo: () => {
            sheet.setCell(row, col, { ...next });
          },
        });
      },
    };
  }
}
