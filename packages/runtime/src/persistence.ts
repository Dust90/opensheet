// Snapshot persistence (M2.6): save committed workbook state to a storage
// backend (localStorage by default) with debounce, and restore it on load.
//
// M2 semantics:
//   - Persisted data comes ONLY from committed transactions: autoSave listens
//     to merged ChangeEvents and snapshots the workbook AFTER each commit.
//     In-editor text never reaches storage.
//   - Corrupt / wrong-version snapshots are REJECTED and do NOT overwrite the
//     current workbook — restore() returns null and leaves storage untouched.
//   - Debounced saves coalesce bursts of edits into one write.

import type { Unsubscribe, WorkbookSnapshot } from "@opensheet/shared";
import { CELL_ERROR_TYPES, MAX_COLS, MAX_ROWS, WORKBOOK_SNAPSHOT_VERSION } from "@opensheet/shared";
import type { OpenSheetAPI, WorkbookInfo } from "./api.js";

export interface StorageLike {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

export interface PersistenceOptions {
  /** Storage backend; defaults to window.localStorage when available. */
  storage?: StorageLike;
  key?: string;
  debounceMs?: number;
}

export interface Persistence {
  /** Try to restore a workbook; null when no snapshot or it is rejected. */
  restore(): WorkbookInfo | null;
  /** Persist immediately (skips debounce). */
  saveNow(): void;
  /** Execute any pending debounced save NOW (pagehide/visibilitychange). */
  flush(): void;
  /**
   * Listen to commits and save (debounced). The returned stop function also
   * FLUSHES any pending save — stopping the listener never drops the last
   * committed edit.
   */
  autoSave(): Unsubscribe;
  /** Remove the stored snapshot (used by tests / "clear data"). */
  clear(): void;
}

const DEFAULT_KEY = "opensheet:workbook";

/**
 * Full structural + range validation. Anything that does not match the
 * current snapshot contract is rejected (returns false) WITHOUT touching
 * storage. Checks beyond shape: positive integer sizes within MAX limits,
 * freeze bounds, activeSheetId existence, cell key legality + bounds, and
 * row-height/column-width entries.
 */
export function validateSnapshot(value: unknown): value is WorkbookSnapshot {
  if (!isPlainRecord(value)) return false;
  const v = value as Record<string, unknown>;
  if (v.version !== WORKBOOK_SNAPSHOT_VERSION) return false;
  if (typeof v.id !== "string" || typeof v.name !== "string") return false;
  if (typeof v.activeSheetId !== "string") return false;
  if (!Array.isArray(v.sheets) || v.sheets.length === 0) return false;
  if (!isPlainRecord(v.styles)) return false;
  if (!v.sheets.every(validateWorksheetSnapshot)) return false;
  // Sheet ids must be unique (loader keys sheets by id).
  const ids = (v.sheets as unknown[]).map((sheet) => (sheet as Record<string, unknown>).id);
  if (new Set(ids).size !== ids.length) return false;
  // activeSheetId must reference an existing sheet.
  return (v.sheets as unknown[]).some(
    (sheet) => (sheet as Record<string, unknown>).id === v.activeSheetId,
  );
}

/**
 * Strict Version-1 validation: every contract field is REQUIRED — no
 * implicit defaults. "Valid" here is exactly "workbookFromSnapshot can load
 * this without throwing".
 */
function validateWorksheetSnapshot(sheet: unknown): boolean {
  if (!isPlainRecord(sheet)) return false;
  const s = sheet as Record<string, unknown>;
  if (typeof s.id !== "string" || typeof s.name !== "string") return false;
  if (!isBoundedSize(s.rowCount, MAX_ROWS) || !isBoundedSize(s.columnCount, MAX_COLS)) return false;
  const rowCount = s.rowCount as number;
  const columnCount = s.columnCount as number;
  // Required fields — absence is a contract violation, not a default.
  if (!isFreeze(s.frozenRows, rowCount) || !isFreeze(s.frozenColumns, columnCount)) return false;
  if (!isPlainRecord(s.cells)) return false;
  for (const [key, data] of Object.entries(s.cells as Record<string, unknown>)) {
    if (!isCellKeyInBounds(key, rowCount, columnCount)) return false;
    if (!isValidCellData(data)) return false;
  }
  if (!isSizeMap(s.rowHeights, rowCount)) return false;
  if (!isSizeMap(s.columnWidths, columnCount)) return false;
  return true;
}

/** Plain object (not null, not an array). */
function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isBoundedSize(value: unknown, max: number): boolean {
  return typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= max;
}

function isFreeze(value: unknown, max: number): boolean {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= max;
}

/** "row:col" with non-negative integers inside the sheet bounds. */
function isCellKeyInBounds(key: string, rowCount: number, columnCount: number): boolean {
  const match = /^(\d+):(\d+)$/.exec(key);
  if (match === null) return false;
  return Number(match[1]) < rowCount && Number(match[2]) < columnCount;
}

const CELL_ERROR_TYPES_SET = new Set<string>(CELL_ERROR_TYPES);

/** CellData shape: legal value + optional string metadata. */
function isValidCellData(value: unknown): boolean {
  if (!isPlainRecord(value)) return false;
  const data = value as Record<string, unknown>;
  if (!isValidCellValue(data.value)) return false;
  for (const field of ["formula", "styleId", "numberFormat"]) {
    if (data[field] !== undefined && typeof data[field] !== "string") return false;
  }
  return true;
}

function isValidCellValue(value: unknown): boolean {
  const primitive = typeof value === "string" || typeof value === "number" || typeof value === "boolean" || value === null;
  if (primitive) return true;
  if (!isPlainRecord(value)) return false;
  // CellError
  if (!CELL_ERROR_TYPES_SET.has(String((value as Record<string, unknown>).type))) return false;
  const message = (value as Record<string, unknown>).message;
  return message === undefined || typeof message === "string";
}

/** Index → positive finite size, with valid integer indices inside bounds. */
function isSizeMap(value: unknown, bound: number): boolean {
  if (!isPlainRecord(value)) return false;
  for (const [index, size] of Object.entries(value)) {
    if (!/^\d+$/.test(index) || Number(index) >= bound) return false;
    if (typeof size !== "number" || !Number.isFinite(size) || size <= 0) return false;
  }
  return true;
}

export function createPersistence(
  api: OpenSheetAPI,
  options?: PersistenceOptions,
): Persistence {
  const storage: StorageLike | undefined =
    options?.storage ??
    (typeof globalThis.localStorage !== "undefined" ? globalThis.localStorage : undefined);
  const key = options?.key ?? DEFAULT_KEY;
  const debounceMs = options?.debounceMs ?? 500;

  let timer: ReturnType<typeof setTimeout> | undefined;
  let disposed = false;

  function readRaw(): string | null {
    if (storage === undefined) return null;
    try {
      return storage.getItem(key);
    } catch {
      return null; // storage unavailable (privacy mode etc.)
    }
  }

  function writeRaw(value: string): void {
    if (storage === undefined) return;
    try {
      storage.setItem(key, value);
    } catch {
      // Quota exceeded / unavailable — persistence is best-effort.
    }
  }

  function restore(): WorkbookInfo | null {
    const raw = readRaw();
    if (raw === null) return null;
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch {
      return null; // corrupt JSON — rejected, storage untouched
    }
    if (!validateSnapshot(parsed)) return null;
    try {
      return api.loadWorkbook(parsed);
    } catch {
      return null; // load failure — rejected, do not clobber storage
    }
  }

  function saveNow(): void {
    if (disposed) return;
    writeRaw(JSON.stringify(api.getWorkbookSnapshot()));
  }

  function flush(): void {
    // Only a PENDING (debounced) save is flushed. With nothing scheduled this
    // is a no-op — callers (pagehide, teardown, React StrictMode remount)
    // must never persist an untouched workbook over an existing snapshot.
    if (timer === undefined) return;
    clearTimeout(timer);
    timer = undefined;
    saveNow();
  }

  function scheduleSave(): void {
    if (disposed) return;
    if (timer !== undefined) clearTimeout(timer);
    timer = setTimeout(() => {
      timer = undefined;
      saveNow();
    }, debounceMs);
  }

  function autoSave(): Unsubscribe {
    // Debounce keyed on commit events; the snapshot is taken AFTER the
    // merged commit, so it always reflects committed data only.
    const unsubscribe = api.onChange(() => scheduleSave());
    return () => {
      unsubscribe();
      // Never drop the last committed edit when the host tears down.
      flush();
    };
  }

  function clear(): void {
    if (storage === undefined) return;
    try {
      storage.removeItem(key);
    } catch {
      // ignore
    }
  }

  return { restore, saveNow, flush, autoSave, clear };
}
