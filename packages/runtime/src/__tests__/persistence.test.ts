// M2.6 persistence: restore / autoSave debounce / corrupt-rejection.

import { describe, expect, it, vi } from "vitest";
import { createOpenSheet, createPersistence, validateSnapshot, type StorageLike } from "../index.js";
import { MAX_COLS, MAX_ROWS, WORKBOOK_SNAPSHOT_VERSION, type WorkbookSnapshot } from "@opensheet/shared";

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

  it("autoSave stop flushes the pending save (no lost edit on teardown)", () => {
    vi.useFakeTimers();
    const api = createOpenSheet();
    const wb = api.createWorkbook({ name: "Flush" });
    const storage = memoryStorage();
    const persistence = createPersistence(api, { storage, debounceMs: 1000 });
    const stop = persistence.autoSave();

    api.applyOperations({
      workbookId: wb.id,
      sheetId: wb.activeSheetId,
      atomic: true,
      operations: [{ type: "range.write", range: "D4", values: [["last-edit"]] }],
    });
    expect(storage.data["opensheet:workbook"]).toBeUndefined(); // still pending
    stop(); // teardown BEFORE the debounce fires
    expect(storage.data["opensheet:workbook"]).toBeDefined();
    expect(JSON.parse(storage.data["opensheet:workbook"]!).sheets[0].cells["3:3"].value).toBe("last-edit");
    vi.useRealTimers();
  });

  it("flush() persists a pending save immediately", () => {
    vi.useFakeTimers();
    const api = createOpenSheet();
    const wb = api.createWorkbook({ name: "F" });
    const storage = memoryStorage();
    const persistence = createPersistence(api, { storage, debounceMs: 1000 });
    persistence.autoSave();
    api.applyOperations({
      workbookId: wb.id,
      sheetId: wb.activeSheetId,
      atomic: true,
      operations: [{ type: "range.write", range: "A1", values: [["x"]] }],
    });
    persistence.flush();
    expect(storage.data["opensheet:workbook"]).toBeDefined();
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
        sheets: [
          {
            id: "s",
            name: "S",
            rowCount: 1,
            columnCount: 1,
            cells: {},
            rowHeights: {},
            columnWidths: {},
            frozenRows: 0,
            frozenColumns: 0,
          },
        ],
        styles: {},
      }),
    ).toBe(true);
    expect(validateSnapshot(null)).toBe(false);
    expect(validateSnapshot("nope")).toBe(false);
    expect(validateSnapshot({ version: 0, id: "x", name: "y", activeSheetId: "s", sheets: [], styles: {} })).toBe(false);
  });

  it("M2.8: rejects out-of-range sizes, freeze, ids, cell keys and size maps", () => {
    const base = {
      version: WORKBOOK_SNAPSHOT_VERSION,
      id: "wb",
      name: "n",
      activeSheetId: "s",
      styles: {},
    };
    const sheet = (over: Record<string, unknown>) => ({
      id: "s",
      name: "S",
      rowCount: 10,
      columnCount: 5,
      cells: {},
      rowHeights: {},
      columnWidths: {},
      frozenRows: 0,
      frozenColumns: 0,
      ...over,
    });
    const make = (s: Record<string, unknown>) => ({ ...base, sheets: [sheet(s)] });

    expect(validateSnapshot(make({ rowCount: 0 }))).toBe(false); // zero rows
    expect(validateSnapshot(make({ rowCount: -3 }))).toBe(false); // negative
    expect(validateSnapshot(make({ rowCount: 2.5 }))).toBe(false); // non-integer
    expect(validateSnapshot(make({ rowCount: MAX_ROWS + 1 }))).toBe(false); // over max
    expect(validateSnapshot(make({ columnCount: MAX_COLS + 1 }))).toBe(false);

    expect(validateSnapshot(make({ frozenRows: 11 }))).toBe(false); // freeze > rows
    expect(validateSnapshot(make({ frozenColumns: -1 }))).toBe(false);

    expect(validateSnapshot({ ...base, activeSheetId: "missing", sheets: [sheet({})] })).toBe(false);

    expect(validateSnapshot(make({ cells: { "9:4": { value: 1 } } }))).toBe(true); // boundary ok
    expect(validateSnapshot(make({ cells: { "10:0": { value: 1 } } }))).toBe(false); // row out of bounds
    expect(validateSnapshot(make({ cells: { "0:5": { value: 1 } } }))).toBe(false); // col out of bounds
    expect(validateSnapshot(make({ cells: { "abc": { value: 1 } } }))).toBe(false); // malformed key
    expect(validateSnapshot(make({ cells: { "-1:0": { value: 1 } } }))).toBe(false);

    expect(validateSnapshot(make({ rowHeights: { 3: 24 } }))).toBe(true);
    expect(validateSnapshot(make({ rowHeights: { 10: 24 } }))).toBe(false); // index out of bounds
    expect(validateSnapshot(make({ rowHeights: { 3: 0 } }))).toBe(false); // non-positive size
    expect(validateSnapshot(make({ rowHeights: { 3: NaN } }))).toBe(false);
    expect(validateSnapshot(make({ columnWidths: { 4: 80 } }))).toBe(true);
    expect(validateSnapshot(make({ columnWidths: { 5: 80 } }))).toBe(false);
  });

  it("M2.9: required fields are enforced — no implicit defaults", () => {
    const base = {
      version: WORKBOOK_SNAPSHOT_VERSION,
      id: "wb",
      name: "n",
      activeSheetId: "s",
      styles: {},
    };
    const full = {
      id: "s",
      name: "S",
      rowCount: 10,
      columnCount: 5,
      cells: {},
      rowHeights: {},
      columnWidths: {},
      frozenRows: 0,
      frozenColumns: 0,
    };
    const make = (s: unknown) => ({ ...base, sheets: [s] });

    expect(validateSnapshot(make(full))).toBe(true);
    // Each required field, when removed, must be rejected (no defaulting).
    for (const field of ["rowHeights", "columnWidths", "frozenRows", "frozenColumns", "cells"] as const) {
      const { [field]: _omitted, ...rest } = full;
      expect(validateSnapshot(make(rest))).toBe(false);
    }
    expect(validateSnapshot(make({ ...full, cells: [] }))).toBe(false); // cells must be an object, not array
    expect(validateSnapshot(make({ ...full, rowHeights: [] }))).toBe(false); // size maps too
    expect(validateSnapshot(make({ ...full, frozenRows: "0" }))).toBe(false); // not a number

    // styles must be a plain object (not array).
    expect(validateSnapshot({ ...base, styles: [] })).toBe(false);

    // Duplicate sheet ids rejected.
    expect(
      validateSnapshot({
        ...base,
        sheets: [
          full,
          { ...full, id: "s", name: "other" },
        ],
      }),
    ).toBe(false);
  });

  it("M2.9: CellData values and metadata are validated", () => {
    const base = {
      version: WORKBOOK_SNAPSHOT_VERSION,
      id: "wb",
      name: "n",
      activeSheetId: "s",
      styles: {},
    };
    const full = {
      id: "s",
      name: "S",
      rowCount: 10,
      columnCount: 5,
      cells: {},
      rowHeights: {},
      columnWidths: {},
      frozenRows: 0,
      frozenColumns: 0,
    };
    const make = (cells: Record<string, unknown>) =>
      validateSnapshot({ ...base, sheets: [{ ...full, cells }] });

    expect(make({ "0:0": { value: "x" } })).toBe(true);
    expect(make({ "0:0": { value: 1.5 } })).toBe(true);
    expect(make({ "0:0": { value: true } })).toBe(true);
    expect(make({ "0:0": { value: null } })).toBe(true);
    expect(make({ "0:0": { value: { type: "#DIV/0!", message: "boom" } } })).toBe(true);
    expect(make({ "0:0": { value: { type: "#NOPE" } } })).toBe(false); // unsupported error type
    expect(make({ "0:0": { value: {} } })).toBe(false); // error without type
    expect(make({ "0:0": { value: undefined } })).toBe(false); // missing value
    expect(make({ "0:0": { value: "x", formula: 5 } })).toBe(false); // metadata must be strings
    expect(make({ "0:0": { value: "x", styleId: true } })).toBe(false);
    expect(make({ "0:0": { value: "x", numberFormat: {} } })).toBe(false);
    expect(make({ "0:0": "not-an-object" })).toBe(false);
    // Numbers must be finite.
    expect(make({ "0:0": { value: NaN } })).toBe(false);
    expect(make({ "0:0": { value: Infinity } })).toBe(false);
  });

  it("M2.9b: styleId must reference an existing style; styles are validated", () => {
    const base = {
      version: WORKBOOK_SNAPSHOT_VERSION,
      id: "wb",
      name: "n",
      activeSheetId: "s",
      styles: {},
    };
    const full = {
      id: "s",
      name: "S",
      rowCount: 10,
      columnCount: 5,
      cells: {},
      rowHeights: {},
      columnWidths: {},
      frozenRows: 0,
      frozenColumns: 0,
    };
    const make = (cells: Record<string, unknown>, styles: Record<string, unknown>) =>
      validateSnapshot({ ...base, styles, sheets: [{ ...full, cells }] });

    // styleId must exist in styles.
    expect(make({ "0:0": { value: "x", styleId: "s1" } }, { s1: { bold: true } })).toBe(true);
    expect(make({ "0:0": { value: "x", styleId: "missing" } }, {})).toBe(false);

    // Style field validation.
    expect(make({}, { s1: { bold: true, italic: false } })).toBe(true);
    expect(make({}, { s1: { bold: "yes" } })).toBe(false); // non-boolean
    expect(make({}, { s1: { fontSize: 14 } })).toBe(true);
    expect(make({}, { s1: { fontSize: -3 } })).toBe(false); // non-positive
    expect(make({}, { s1: { fontSize: NaN } })).toBe(false);
    expect(make({}, { s1: { textColor: 42 } })).toBe(false); // non-string
    expect(make({}, { s1: { horizontalAlign: "center" } })).toBe(true);
    expect(make({}, { s1: { horizontalAlign: "justify" } })).toBe(false); // bad enum
    expect(make({}, { s1: { verticalAlign: "middle" } })).toBe(true);
    expect(make({}, { s1: { verticalAlign: "top-right" } })).toBe(false);
    expect(make({}, { s1: { border: { top: { style: "thin", color: "#000" } } } })).toBe(true);
    expect(make({}, { s1: { border: { top: { style: "fat" } } } })).toBe(false); // bad border style
    expect(make({}, { s1: { border: { top: "thin" } } })).toBe(false); // edge not an object
    expect(make({}, { s1: null })).toBe(false); // style must be an object
    expect(make({}, { s1: [] })).toBe(false); // not an array
    expect(make({}, { s1: { border: { top: { style: "thin", color: 1 } } } })).toBe(false);

    // Canonical keys: leading zeros are coordinate aliases and must be rejected.
    expect(make({ "00:0": { value: 1 } }, {})).toBe(false);
    expect(make({ "0:05": { value: 1 } }, {})).toBe(false);
    expect(make({ "0:0": { value: 1 } }, {})).toBe(true);
    expect(make({}, {})).toBe(true);
  });

  it("M2.9: validate-true implies loadable (property test)", () => {
    // Any snapshot the validator accepts must load without throwing.
    const valid: WorkbookSnapshot = {
      version: WORKBOOK_SNAPSHOT_VERSION,
      id: "wb",
      name: "n",
      activeSheetId: "s",
      sheets: [
        {
          id: "s",
          name: "S",
          rowCount: 3,
          columnCount: 2,
          cells: {
            "0:0": { value: "a1", styleId: "s1" },
            "1:1": { value: { type: "#REF!" }, formula: "=#REF!" },
            "2:0": { value: 42, numberFormat: "0.00" },
          },
          rowHeights: { 1: 30 },
          columnWidths: { 0: 120 },
          frozenRows: 1,
          frozenColumns: 1,
        },
      ],
      styles: { s1: { bold: true } },
    };
    expect(validateSnapshot(valid)).toBe(true);
    const api = createOpenSheet();
    const wb = api.loadWorkbook(valid); // must not throw
    expect(wb.name).toBe("n");
    const value = api.readRange({ sheetId: wb.activeSheetId, range: "A1" })[0]![0];
    expect(value).toBe("a1");
  });
});
