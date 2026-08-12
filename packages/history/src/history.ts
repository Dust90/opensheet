// Undo/redo history with dual limits: entry count AND retained memory.

import type { CommandBus, HistorySink, JournalBatch } from "@injoysai/opensheet-commands";

export interface HistoryOptions {
  /** Max entries on the undo stack. Default 100. */
  maxEntries?: number;
  /** Max retained journal memory in bytes (sum of approxBytes). Default 16 MiB. */
  maxMemoryBytes?: number;
  /** Called whenever the oldest entries are evicted by either limit. */
  onEvict?: (evictedCount: number) => void;
}

const DEFAULT_MAX_ENTRIES = 100;
const DEFAULT_MAX_MEMORY_BYTES = 16 * 1024 * 1024;

interface StackEntry {
  batch: JournalBatch;
  label: string;
}

/**
 * Records transaction journals pushed by the CommandBus and replays them
 * on undo/redo. "derived" (formula recalc) changes never enter history —
 * the bus simply never pushes them.
 */
export class HistoryManager implements HistorySink {
  private undoStack: StackEntry[] = [];
  private redoStack: StackEntry[] = [];
  private undoBytes = 0;
  private redoBytes = 0;

  private readonly maxEntries: number;
  private readonly maxMemoryBytes: number;
  private readonly onEvict: ((evictedCount: number) => void) | undefined;

  constructor(options?: HistoryOptions) {
    this.maxEntries = options?.maxEntries ?? DEFAULT_MAX_ENTRIES;
    this.maxMemoryBytes = options?.maxMemoryBytes ?? DEFAULT_MAX_MEMORY_BYTES;
    this.onEvict = options?.onEvict;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  get undoDepth(): number {
    return this.undoStack.length;
  }

  get retainedBytes(): number {
    return this.undoBytes + this.redoBytes;
  }

  push(batch: JournalBatch): void {
    // A new user/api action invalidates the redo stack.
    this.redoStack = [];
    this.redoBytes = 0;
    this.undoStack.push({ batch, label: describe(batch) });
    this.undoBytes += batch.approxBytes;
    this.enforceLimits();
  }

  undo(bus: CommandBus): boolean {
    const entry = this.undoStack[this.undoStack.length - 1];
    if (entry === undefined) return false;
    // Replay FIRST: if it throws, both stacks and byte accounting stay intact.
    bus.replayJournal(entry.batch, "undo");
    this.undoStack.pop();
    this.undoBytes -= entry.batch.approxBytes;
    this.redoStack.push(entry);
    this.redoBytes += entry.batch.approxBytes;
    return true;
  }

  redo(bus: CommandBus): boolean {
    const entry = this.redoStack[this.redoStack.length - 1];
    if (entry === undefined) return false;
    bus.replayJournal(entry.batch, "redo");
    this.redoStack.pop();
    this.redoBytes -= entry.batch.approxBytes;
    this.undoStack.push(entry);
    this.undoBytes += entry.batch.approxBytes;
    return true;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.undoBytes = 0;
    this.redoBytes = 0;
  }

  private enforceLimits(): void {
    let evicted = 0;
    while (
      this.undoStack.length > 0 &&
      (this.undoStack.length > this.maxEntries || this.undoBytes > this.maxMemoryBytes)
    ) {
      const removed = this.undoStack.shift()!;
      this.undoBytes -= removed.batch.approxBytes;
      evicted++;
    }
    if (evicted > 0) this.onEvict?.(evicted);
  }
}

function describe(batch: JournalBatch): string {
  const first = batch.entries[0];
  if (first === undefined) return "empty batch";
  if (batch.entries.length === 1) return first.label;
  return `${first.label} +${batch.entries.length - 1}`;
}
