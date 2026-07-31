// M2.6 persistence: restore / autoSave debounce / corrupt-rejection.

import { describe, expect, it, vi } from "vitest";
import { createOpenSheet, createPersistence, validateSnapshot, type StorageLike } from "../index.js";
import { WORKBOOK_SNAPSHOT_VERSION } from "@opensheet/shared";

function memoryStorage(seed: Record<string, string> = {}): StorageLike & { data: Record<string, string> } {
  const data = { ...seed };
  return {
    data,
    getItem: (k) => (k in data ? data[k]! : null),
    setItem: (k, v) => {
      data[k] = v;
    },
    removeItem: (k) => {
      delete data[k];
    },
  };
}

describe("createPersistence", () => {
  it("saveNow → restore round-trips a workbook", () => {
    const api = createOpenSheet();
    const wb = api.createWorkbook({ name: "P" });
    api.applyOperations({
      workbookId: wb.id,
      sheetId: wb.activeSheetId,
      atomic: true,
      operations: [{ type: "range.write", range: "A1:B2", values: [["x", 1], ["y", 2]] }],
    });
    const storage = memoryStorage();
    const persistence = createPersistence(api, { storage, debounceMs: 10 });

    persistence.saveNow();
    expect(storage.data).toHaveProperty("opensheet:workbook");

    // Fresh instance restores the same data.
    const api2 = createOpenSheet();
    const restored = createPersistence(api2, { storage }).restore();
    expect(restored).not.toBeNull();
    expect(restored!.name).toBe("P");
    expect(api2.readRange({ sheetId: restored!.activeSheetId, range: "A1:B2" })).toEqual([
      ["x", 1],
      ["y", 2],
    ]);
  });

  it("autoSave persists committed data after debounce (and not before)", async () => {
    vi.useFakeTimers();
    const api = createOpenSheet();
    const wb = api.createWorkbook({ name: "Auto" });
    const storage = memoryStorage();
    const persistence = createPersistence(api, { storage, debounceMs: 100 });
    persistence.autoSave();

    api.applyOperations({
      workbookId: wb.id,
      sheetId: wb.activeSheetId,
      atomic: true,
      operations: [{ type: "range.write", range: "C3", values: [["committed"]] }],
    });
    // Not yet flushed.
    expect(storage.data["opensheet:workbook"]).toBeUndefined();
    vi.advanceTimersByTime(150);
    expect(storage.data["opensheet:workbook"]).toBeDefined();
    expect(JSON.parse(storage.data["opensheet:workbook"]!).sheets[0].cells["2:2"].value).toBe("committed");
    vi.useRealTimers();
  });

  it("rejects corrupt JSON, wrong version and malformed shapes without touching storage", () => {
    const api = createOpenSheet();
    api.createWorkbook({ name: "X" });
    const storage = memoryStorage();
    const persistence = createPersistence(api, { storage });
    const setRaw = (raw: string) => {
      storage.data["opensheet:workbook"] = raw;
    };

    setRaw("{not json");
    expect(persistence.restore()).toBeNull();
    expect(storage.data["opensheet:workbook"]).toBe("{not json"); // untouched

    setRaw(JSON.stringify({ version: 999, id: "x" }));
    expect(persistence.restore()).toBeNull();
    expect(storage.data["opensheet:workbook"]).toContain('"version":999');

    setRaw(JSON.stringify({ version: WORKBOOK_SNAPSHOT_VERSION, id: "x", name: "y", activeSheetId: "", sheets: [], styles: {} }));
    expect(persistence.restore()).toBeNull(); // empty sheets rejected

    setRaw(JSON.stringify({ version: WORKBOOK_SNAPSHOT_VERSION, id: "x", name: "y", activeSheetId: "", sheets: [{ id: "s", name: "S" }], styles: {} }));
    expect(persistence.restore()).toBeNull(); // missing rowCount etc.
  });

  it("validateSnapshot accepts a well-formed snapshot", () => {
    expect(
      validateSnapshot({
        version: WORKBOOK_SNAPSHOT_VERSION,
        id: "wb",
        name: "n",
        activeSheetId: "s",
        sheets: [{ id: "s", name: "S", rowCount: 1, columnCount: 1, cells: {} }],
        styles: {},
      }),
    ).toBe(true);
    expect(validateSnapshot(null)).toBe(false);
    expect(validateSnapshot("nope")).toBe(false);
    expect(validateSnapshot({ version: 0, id: "x", name: "y", activeSheetId: "s", sheets: [], styles: {} })).toBe(false);
  });
});
