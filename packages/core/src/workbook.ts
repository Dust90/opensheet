// Workbook aggregate root + change-event bus with transaction buffering.

import type { ChangeEvent, ChangeListener, Unsubscribe } from "@injoysai/opensheet-shared";
import { SheetError } from "@injoysai/opensheet-shared";
import { StyleTable } from "./styles.js";
import type { WorkbookView, WorksheetView } from "./view.js";
import { Worksheet } from "./worksheet.js";

export interface WorkbookInit {
  id: string;
  name: string;
}

/**
 * Event semantics (transaction boundary, see ADR-0003):
 * - Outside a batch, emit() dispatches immediately.
 * - Inside beginBatch(), emissions accumulate.
 * - endBatch(true)  → at most ONE merged event per (sheet + source) is dispatched.
 * - endBatch(false) → buffer is discarded; observers saw nothing.
 * Listener exceptions are isolated via onListenerError and never affect commits.
 */
export class Workbook {
  readonly id: string;
  name: string;
  version = 1;
  readonly styles = new StyleTable();

  /**
   * Called when a change listener throws. Listener failures NEVER affect
   * transaction outcomes — data stays committed and history is written.
   */
  onListenerError: ((error: unknown, event: ChangeEvent) => void) | undefined;

  private sheets: Worksheet[] = [];
  private activeId = "";
  private readonly listeners = new Set<ChangeListener>();
  private batchDepth = 0;
  private buffer: ChangeEvent[] = [];

  constructor(init: WorkbookInit) {
    this.id = init.id;
    this.name = init.name;
  }

  get activeSheetId(): string {
    return this.activeId;
  }

  setActiveSheet(sheetId: string): void {
    this.getSheet(sheetId); // validates existence
    this.activeId = sheetId;
  }

  listSheets(): readonly Worksheet[] {
    return this.sheets;
  }

  getSheet(sheetId: string): Worksheet {
    const sheet = this.sheets.find((s) => s.id === sheetId);
    if (sheet === undefined) {
      throw new SheetError("E_SHEET_NOT_FOUND", `Sheet not found: "${sheetId}"`);
    }
    return sheet;
  }

  /** Read-only sheet access for consumers outside the command path. */
  getSheetView(sheetId: string): WorksheetView {
    return this.getSheet(sheetId).asView();
  }

  /** Read-only view for beforeCommit hooks (M3 guardrail: hooks get no
   *  writable surface; they write only through DerivedWriter). */
  asView(): WorkbookView {
    return {
      id: this.id,
      name: this.name,
      activeSheetId: this.activeSheetId,
      getSheetView: (sheetId: string) => this.getSheetView(sheetId),
      listSheetViews: () => this.sheets.map((s) => s.asView()),
      resolveStyle: (styleId: string) => this.styles.get(styleId),
    };
  }

  addSheet(sheet: Worksheet): void {
    if (this.sheets.some((s) => s.id === sheet.id)) {
      throw new SheetError("E_VALIDATION", `Duplicate sheet id: "${sheet.id}"`);
    }
    this.sheets.push(sheet);
    if (this.activeId === "") this.activeId = sheet.id;
  }

  removeSheet(sheetId: string): Worksheet {
    const index = this.sheets.findIndex((s) => s.id === sheetId);
    if (index < 0) {
      throw new SheetError("E_SHEET_NOT_FOUND", `Sheet not found: "${sheetId}"`);
    }
    const [removed] = this.sheets.splice(index, 1);
    if (this.activeId === sheetId) {
      this.activeId = this.sheets[0]?.id ?? "";
    }
    return removed!;
  }

  restoreSheet(sheet: Worksheet, index: number): void {
    this.sheets.splice(Math.min(index, this.sheets.length), 0, sheet);
    if (this.activeId === "") this.activeId = sheet.id;
  }

  // --- events -------------------------------------------------------------

  onChange(listener: ChangeListener): Unsubscribe {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  emit(event: ChangeEvent): void {
    if (this.batchDepth > 0) {
      this.buffer.push(event);
      return;
    }
    this.dispatch(event);
  }

  beginBatch(): void {
    this.batchDepth++;
  }

  endBatch(commit: boolean): void {
    if (this.batchDepth === 0) {
      throw new SheetError("E_OP_FAILED", "endBatch() called outside a batch");
    }
    this.batchDepth--;
    if (this.batchDepth > 0) {
      if (!commit) this.buffer = [];
      return; // nested batch: outermost endBatch decides
    }
    const pending = this.buffer;
    this.buffer = [];
    if (!commit) return;
    for (const merged of mergeEvents(pending)) {
      this.dispatch(merged);
    }
  }

  private dispatch(event: ChangeEvent): void {
    for (const listener of this.listeners) {
      try {
        listener(event);
      } catch (error) {
        // Isolated: a broken observer must not corrupt commit/history.
        try {
          this.onListenerError?.(error, event);
        } catch {
          // Error reporting must never affect workbook state.
        }
      }
    }
  }
}

function mergeEvents(events: ChangeEvent[]): ChangeEvent[] {
  const bySheetAndSource = new Map<string, ChangeEvent>();
  for (const event of events) {
    const key = `${event.sheetId}|${event.source}`;
    const existing = bySheetAndSource.get(key);
    if (existing !== undefined) {
      existing.changes.push(...event.changes);
    } else {
      bySheetAndSource.set(key, { ...event, changes: [...event.changes], batch: true });
    }
  }
  return [...bySheetAndSource.values()];
}
