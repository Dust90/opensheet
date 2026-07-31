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
import { WORKBOOK_SNAPSHOT_VERSION } from "@opensheet/shared";
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
  /** Listen to commits and save (debounced). Returns unsubscribe. */
  autoSave(): Unsubscribe;
  /** Remove the stored snapshot (used by tests / "clear data"). */
  clear(): void;
}

const DEFAULT_KEY = "opensheet:workbook";

/**
 * Structural + version validation. Anything that does not match the current
 * snapshot contract is rejected (returns false) WITHOUT touching storage.
 */
export function validateSnapshot(value: unknown): value is WorkbookSnapshot {
  if (typeof value !== "object" || value === null) return false;
  const v = value as Record<string, unknown>;
  if (v.version !== WORKBOOK_SNAPSHOT_VERSION) return false;
  if (typeof v.id !== "string" || typeof v.name !== "string") return false;
  if (typeof v.activeSheetId !== "string") return false;
  if (!Array.isArray(v.sheets) || v.sheets.length === 0) return false;
  if (typeof v.styles !== "object" || v.styles === null) return false;
  return v.sheets.every(
    (sheet) =>
      typeof sheet === "object" &&
      sheet !== null &&
      typeof (sheet as Record<string, unknown>).id === "string" &&
      typeof (sheet as Record<string, unknown>).name === "string" &&
      typeof (sheet as Record<string, unknown>).rowCount === "number" &&
      typeof (sheet as Record<string, unknown>).columnCount === "number" &&
      typeof (sheet as Record<string, unknown>).cells === "object",
  );
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
      if (timer !== undefined) {
        clearTimeout(timer);
        timer = undefined;
      }
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

  return { restore, saveNow, autoSave, clear };
}
