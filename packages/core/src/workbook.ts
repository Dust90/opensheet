// Workbook aggregate root + change-event bus with transaction buffering.

import type { ChangeEvent, ChangeListener, Unsubscribe } from "@opensheet/shared";
import { SheetError } from "@opensheet/shared";
import { StyleTable } from "./styles.js";
import { Worksheet } from "./worksheet.js";

export interface WorkbookInit {
  id: string;
  name: string;
}

/**
 * Event semantics (transaction boundary, see ADR-0003):
 * - Outside a batch, emit() dispatches immediately.
 * - Inside beginBatch(), emissions accumulate.
 * - endBatch(true)  → listeners receive exactly ONE merged event per sheet.
 * - endBatch(false) → buffer is discarded; observers saw nothing.
 */
export class Workbook {
  readonly id: string;
  name: string;
  version = 1;
  readonly styles = new StyleTable();

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
    for (const listener of this.listeners) listener(event);
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
