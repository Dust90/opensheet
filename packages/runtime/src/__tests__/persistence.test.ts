// M2.6 persistence: restore / autoSave debounce / corrupt-rejection.

import { describe, expect, it, vi } from "vitest";
import { createOpenSheet, createPersistence, validateSnapshot, type StorageLike } from "../index.js";
import { MAX_COLS, MAX_ROWS, WORKBOOK_SNAPSHOT_VERSION, type CellError, type WorkbookSnapshot } from "@opensheet/shared";

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

// ── M3.6 Fix 8: Snapshot formula syntax validation ────────────────────────

describe("M3.6 Fix 8: Snapshot formula syntax validation", () => {
  const base = {
    version: WORKBOOK_SNAPSHOT_VERSION,
    id: "wb",
    name: "n",
    activeSheetId: "s",
    styles: {},
  };
  const fullSheet = {
    id: "s",
    name: "S",
    rowCount: 10,
    columnCount: 5,
    cells: {} as Record<string, unknown>,
    rowHeights: {},
    columnWidths: {},
    frozenRows: 0,
    frozenColumns: 0,
  };
  const make = (cells: Record<string, unknown>) =>
    validateSnapshot({ ...base, sheets: [{ ...fullSheet, cells }] });

  it("rejects formula that does not start with =", () => {
    expect(make({ "0:0": { value: null, formula: "SUM(A1)" } })).toBe(false);
  });

  it("rejects syntactically invalid formula (unclosed paren)", () => {
    expect(make({ "0:0": { value: null, formula: "=SUM(" } })).toBe(false);
  });

  it("accepts a well-formed formula", () => {
    expect(make({ "0:0": { value: 2, formula: "=1+1" } })).toBe(true);
  });

  it("validateSnapshot(true) implies loadWorkbook does not throw", () => {
    const snapshot = {
      ...base,
      sheets: [{ ...fullSheet, cells: { "0:0": { value: 2, formula: "=1+1" } } }],
    };
    expect(validateSnapshot(snapshot)).toBe(true);
    const api = createOpenSheet();
    expect(() => api.loadWorkbook(snapshot as WorkbookSnapshot)).not.toThrow();
  });
});

// ── M3.6 Fix 1: Multi-sheet formula isolation ─────────────────────────────

describe("M3.6 Fix 1: Multi-sheet formula isolation", () => {
  it("formula on Sheet2 does not affect Sheet1 and vice versa", async () => {
    const api = createOpenSheet();
    const wb = api.createWorkbook({ name: "MS" });
    const [sheet1Id] = [wb.activeSheetId];

    // Create a second sheet
    const sheet2 = api.createSheet({ name: "Sheet2", rows: 100, columns: 10 });

    // Set A1=10 on Sheet1, A1==2+2 on Sheet2
    await api.applyOperations({
      workbookId: wb.id, sheetId: sheet1Id!, atomic: true,
      operations: [{ type: "range.write", range: "A1", values: [[10]] }],
    });
    await api.applyOperations({
      workbookId: wb.id, sheetId: sheet2.id, atomic: true,
      operations: [{ type: "formula.set", range: "A1", formula: "=2+2" }],
    });

    // Sheet1 A1 must still be 10 (literal), Sheet2 A1 must be 4 (formula result).
    const s1 = api.readRange({ sheetId: sheet1Id!, range: "A1" })[0]![0];
    const s2 = api.readRange({ sheetId: sheet2.id, range: "A1" })[0]![0];
    expect(s1).toBe(10);
    expect(s2).toBe(4);
  });

  it("Snapshot load rebuilds formulas on all sheets independently", () => {
    const snapshot: WorkbookSnapshot = {
      version: WORKBOOK_SNAPSHOT_VERSION,
      id: "wb",
      name: "MS",
      activeSheetId: "s1",
      sheets: [
        {
          id: "s1", name: "Sheet1", rowCount: 10, columnCount: 5,
          cells: { "0:0": { value: 99 /* stale */, formula: "=1+1" } },
          rowHeights: {}, columnWidths: {}, frozenRows: 0, frozenColumns: 0,
        },
        {
          id: "s2", name: "Sheet2", rowCount: 10, columnCount: 5,
          cells: { "0:0": { value: 99 /* stale */, formula: "=3+4" } },
          rowHeights: {}, columnWidths: {}, frozenRows: 0, frozenColumns: 0,
        },
      ],
      styles: {},
    };
    expect(validateSnapshot(snapshot)).toBe(true);
    const api = createOpenSheet();
    const wb = api.loadWorkbook(snapshot);
    // Both sheets must be recalculated (stale cache 99 replaced).
    expect(api.readRange({ sheetId: "s1", range: "A1" })[0]![0]).toBe(2);
    expect(api.readRange({ sheetId: "s2", range: "A1" })[0]![0]).toBe(7);
  });
});

// ── M3.6 Fix 3: Snapshot cycle order ─────────────────────────────────────

describe("M3.6 Fix 3: Snapshot rebuild writes #CYCLE! before evaluating downstream", () => {
  it("downstream formula does not see stale cached value from a cycle member", () => {
    const snapshot: WorkbookSnapshot = {
      version: WORKBOOK_SNAPSHOT_VERSION,
      id: "wb",
      name: "CY",
      activeSheetId: "s",
      sheets: [
        {
          id: "s", name: "S", rowCount: 10, columnCount: 5,
          cells: {
            // A1 and B1 form a cycle; C1 depends on A1.
            "0:0": { value: 999 /* stale */, formula: "=B1+1" },  // A1
            "0:1": { value: 999 /* stale */, formula: "=A1+1" },  // B1
            "0:2": { value: 1998 /* stale */, formula: "=A1*2" }, // C1
          },
          rowHeights: {}, columnWidths: {}, frozenRows: 0, frozenColumns: 0,
        },
      ],
      styles: {},
    };
    expect(validateSnapshot(snapshot)).toBe(true);
    const api = createOpenSheet();
    const wb = api.loadWorkbook(snapshot);
    const sheetId = wb.activeSheetId;
    // A1 and B1 must be #CYCLE! (not 999).
    expect(api.readRange({ sheetId, range: "A1" })[0]![0]).toMatchObject({ type: "#CYCLE!" });
    expect(api.readRange({ sheetId, range: "B1" })[0]![0]).toMatchObject({ type: "#CYCLE!" });
    // C1 depends on A1 which is #CYCLE!, so C1 must also be a #CYCLE! error
    // (or any error), NOT the stale value 1998.
    const c1 = api.readRange({ sheetId, range: "C1" })[0]![0];
    expect(c1).not.toBe(1998);
    expect(typeof c1 === "object" && c1 !== null && "type" in (c1 as object)).toBe(true);
  });
});

// ── M3.7 guardrail: transaction total budget ──────────────────────────────

describe("M3.7 guardrail: transaction total budget exhaustion", () => {
  it("formulas in ONE atomic transaction share the budget; excess yields #VALUE! and the transaction still commits", async () => {
    const api = createOpenSheet({
      formula: { maxCellReadsPerFormula: 1_000, maxCellReadsPerTransaction: 4 },
    });
    const wb = api.createWorkbook({ name: "Budget" });
    const sheetId = wb.activeSheetId;

    // ONE atomic transaction carrying both formulas. B1 = SUM(A1:A3) consumes
    // 4 evaluation budget units (3 cell reads + 1 range-node unit) — exactly
    // the shared transaction budget; C1 = B1*2 is evaluated after B1
    // (dependency order) and finds the budget exhausted.
    const result = await api.applyOperations({
      workbookId: wb.id, sheetId, atomic: true,
      operations: [
        { type: "range.write", range: "A1:A3", values: [[1], [2], [3]] },
        { type: "formula.set", range: "B1", formula: "=SUM(A1:A3)" },
        { type: "formula.set", range: "C1", formula: "=B1*2" },
      ],
    });

    // The user transaction committed normally — a budget-exceeded formula is
    // an error VALUE, never a transaction failure.
    expect(result.status).toBe("completed");

    // First formula consumed exactly the shared budget and succeeded.
    // NOTE: `maxCellReadsPerTransaction` counts evaluation budget units
    // (cell reads + range/AST overhead), not purely cell reads.
    expect(api.readRange({ sheetId, range: "B1" })[0]![0]).toBe(6);

    // Second formula exceeds the TRANSACTION budget (its own budget is fine).
    const c1 = api.readRange({ sheetId, range: "C1" })[0]![0];
    expect(c1).toMatchObject({ type: "#VALUE!" });
    expect((c1 as CellError).message).toContain("limit exceeded");

    // Both formula sources persisted — the commit was not rolled back.
    const view = api.getWorksheetView(sheetId);
    expect(view.getCell(0, 1)?.formula).toBe("=SUM(A1:A3)");
    expect(view.getCell(0, 2)?.formula).toBe("=B1*2");

    // The whole transaction is ONE history entry: a single undo removes the
    // seed data and both formulas atomically.
    api.undo();
    const after = api.getWorksheetView(sheetId);
    expect(api.readRange({ sheetId, range: "A1" })[0]![0]).toBeNull();
    expect(after.getCell(0, 1)?.formula).toBeUndefined();
    expect(after.getCell(0, 2)?.formula).toBeUndefined();
  });
});

// ── M3.7 guardrail: structural delete clears graph nodes ─────────────────

describe("M3.7 guardrail: structural delete removes stale graph nodes", () => {
  it("deleting last row removes its formula from the graph and Undo restores it", async () => {
    const api = createOpenSheet();
    const wb = api.createWorkbook({ name: "Del" });
    const sheetId = wb.activeSheetId;

    // Put a formula in row 5 (0-based row 4).
    await api.applyOperations({
      workbookId: wb.id, sheetId, atomic: true,
      operations: [
        { type: "range.write", range: "A1", values: [[10]] },
        { type: "formula.set", range: "A5", formula: "=A1*2" },
      ],
    });
    expect(api.readRange({ sheetId, range: "A5" })[0]![0]).toBe(20);

    // Delete rows 3-5 (0-based: rows 2–4 inclusive = 3 rows starting at index 2).
    await api.applyOperations({
      workbookId: wb.id, sheetId, atomic: true,
      operations: [{ type: "row.delete", at: 2, count: 3 }],
    });

    // The formula should be gone — reading the row that was deleted no longer
    // contains a formula result.  A5 is now what was A8 (empty).
    const afterDelete = api.readRange({ sheetId, range: "A3" })[0]![0];
    // The old formula row is gone; the cell should be null or empty.
    expect(afterDelete).toBeNull();

    // Undo: formula row returns and is recalculated correctly.
    api.undo();
    expect(api.readRange({ sheetId, range: "A5" })[0]![0]).toBe(20);
  });
});

// ── M3.7 guardrail: multi-sheet Undo/Redo isolation ──────────────────────

describe("M3.7 guardrail: multi-sheet Undo/Redo formula isolation", () => {
  it("undo on Sheet1 formula does not touch Sheet2 formula", async () => {
    const api = createOpenSheet();
    const wb = api.createWorkbook({ name: "UndoMS" });
    const sheet1Id = wb.activeSheetId;
    const sheet2 = api.createSheet({ name: "S2", rows: 100, columns: 10 });

    // Set A1=5 on Sheet1, then formula =A1*3 on Sheet1.
    await api.applyOperations({
      workbookId: wb.id, sheetId: sheet1Id, atomic: true,
      operations: [{ type: "range.write", range: "A1", values: [[5]] }],
    });
    await api.applyOperations({
      workbookId: wb.id, sheetId: sheet1Id, atomic: true,
      operations: [{ type: "formula.set", range: "B1", formula: "=A1*3" }],
    });

    // Set formula on Sheet2.
    await api.applyOperations({
      workbookId: wb.id, sheetId: sheet2.id, atomic: true,
      operations: [{ type: "formula.set", range: "A1", formula: "=7+8" }],
    });

    expect(api.readRange({ sheetId: sheet1Id, range: "B1" })[0]![0]).toBe(15);
    expect(api.readRange({ sheetId: sheet2.id, range: "A1" })[0]![0]).toBe(15);

    // Undo the Sheet2 formula.
    api.undo();
    // Sheet2 A1 should revert (formula gone → null or 0).
    const s2After = api.readRange({ sheetId: sheet2.id, range: "A1" })[0]![0];
    expect(s2After === null || s2After === 0).toBe(true);
    // Sheet1 B1 must be completely unaffected.
    expect(api.readRange({ sheetId: sheet1Id, range: "B1" })[0]![0]).toBe(15);
  });
});
