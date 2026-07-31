// M3 integration: formula.set → beforeCommit recalc → derived writes,
// undo/redo, literal overwrite, structure reference rewriting, snapshot
// rebuild — all through the real Command Bus + runtime composition root.

import { describe, expect, it } from "vitest";
import { createOpenSheet } from "@opensheet/runtime";
import type { CellValue } from "@opensheet/shared";

function setup() {
  const api = createOpenSheet();
  const wb = api.createWorkbook({ name: "F" });
  const sheetId = wb.activeSheetId;
  const value = (row: number, col: number): CellValue => api.readRange({ sheetId, range: `${colName(col)}${row + 1}` })[0]![0]!;
  const apply = (operations: Parameters<typeof api.applyOperations>[0]["operations"]) =>
    api.applyOperations({ workbookId: wb.id, sheetId, atomic: true, operations });
  return { api, wb, sheetId, value, apply };
}

function colName(col: number): string {
  let name = "";
  let n = col + 1;
  while (n > 0) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

describe("M3 formula runtime", () => {
  it("formula.set recomputes in the SAME commit; derived adds no history", async () => {
    const { apply, value, api } = setup();
    await apply([{ type: "cell.set", range: "A1", value: 1 }]);
    await apply([{ type: "formula.set", range: "B1", formula: "=A1+1" }]);
    expect(value(0, 1)).toBe(2); // recalculated during the same commit
    api.undo();
    expect(value(0, 1)).toBeNull(); // formula removed
    api.redo();
    expect(value(0, 1)).toBe(2);
  });

  it("changing a referenced cell recalculates dependents; undo/redo keeps values", async () => {
    const { apply, value, api } = setup();
    await apply([
      { type: "cell.set", range: "A1", value: 1 },
      { type: "formula.set", range: "B1", formula: "=A1+1" },
      { type: "formula.set", range: "C1", formula: "=B1*10" },
    ]);
    expect(value(0, 1)).toBe(2);
    expect(value(0, 2)).toBe(20);

    await apply([{ type: "cell.set", range: "A1", value: 2 }]);
    expect(value(0, 1)).toBe(3);
    expect(value(0, 2)).toBe(30);

    api.undo();
    expect(value(0, 0)).toBe(1);
    expect(value(0, 1)).toBe(2);
    expect(value(0, 2)).toBe(20);

    api.redo();
    expect(value(0, 0)).toBe(2);
    expect(value(0, 1)).toBe(3);
    expect(value(0, 2)).toBe(30);
  });

  it("literal overwrite removes the graph node; undo restores the formula", async () => {
    const { apply, value, api } = setup();
    await apply([
      { type: "cell.set", range: "A1", value: 1 },
      { type: "formula.set", range: "B1", formula: "=A1+1" },
    ]);
    expect(value(0, 1)).toBe(2);

    // Overwrite B1 with a literal: formula gone, graph node removed.
    await apply([{ type: "cell.set", range: "B1", value: "text" }]);
    expect(value(0, 1)).toBe("text");
    await apply([{ type: "cell.set", range: "A1", value: 5 }]);
    expect(value(0, 1)).toBe("text"); // B1 no longer recalculates

    // Undo the overwrite → formula restored, recalc resumes.
    api.undo(); // undo cell.set A1=5
    api.undo(); // undo B1 literal overwrite
    expect(value(0, 1)).toBe(2); // A1 back to 1, B1 = 1+1
    await apply([{ type: "cell.set", range: "A1", value: 9 }]);
    expect(value(0, 1)).toBe(10); // graph node restored
  });

  it("syntax errors reject formula.set and leave the cell untouched", async () => {
    const { apply, value } = setup();
    await apply([{ type: "cell.set", range: "A1", value: "keep" }]);
    await expect(apply([{ type: "formula.set", range: "A1", formula: "=SUM(" }])).rejects.toThrow();
    expect(value(0, 0)).toBe("keep");
  });

  it("formula.set preserves cell style metadata", async () => {
    const { apply, value, api } = setup();
    await apply([
      { type: "range.style", range: "B1", style: { bold: true } },
      { type: "cell.set", range: "A1", value: 3 },
      { type: "formula.set", range: "B1", formula: "=A1*2" },
    ]);
    expect(value(0, 1)).toBe(6);
    const sheetView = api.getWorksheetView(sheetIdOf(api));
    expect(sheetView.getCell(0, 1)?.styleId).toBeDefined();
    expect(api.resolveStyle(sheetView.getCell(0, 1)!.styleId!)?.bold).toBe(true);
  });

  it("cycles become #CYCLE! and downstream propagates it", async () => {
    const { apply, value } = setup();
    await apply([
      { type: "formula.set", range: "A1", formula: "=B1+1" },
      { type: "formula.set", range: "B1", formula: "=A1+1" }, // cycle
      { type: "formula.set", range: "C1", formula: "=A1*2" }, // downstream
    ]);
    expect(value(0, 0)).toMatchObject({ type: "#CYCLE!" });
    expect(value(0, 1)).toMatchObject({ type: "#CYCLE!" });
    expect(value(0, 2)).toMatchObject({ type: "#CYCLE!" }); // propagates
  });

  it("structure ops rewrite references ($ semantics) and undo restores them", async () => {
    const { apply, value, api } = setup();
    await apply([
      { type: "cell.set", range: "A1", value: 10 },
      { type: "cell.set", range: "A2", value: 20 },
      { type: "formula.set", range: "B1", formula: "=A1" },
      { type: "formula.set", range: "B2", formula: "=$A$1" }, // absolute
    ]);
    expect(value(0, 1)).toBe(10);
    expect(value(1, 1)).toBe(10);

    // Insert a row above A2 → B2 moves to B3, =A1 stays A1, =$A$1 stays.
    await apply([{ type: "row.insert", at: 1, count: 1 }]);
    expect(value(2, 1)).toBe(10); // B3 (old B2)
    expect(value(1, 1)).toBeNull(); // B2 now blank

    // Delete that row back.
    await apply([{ type: "row.delete", at: 1, count: 1 }]);
    expect(value(1, 1)).toBe(10);
    api.undo(); // undo delete
    expect(value(2, 1)).toBe(10);
    api.undo(); // undo insert
    expect(value(1, 1)).toBe(10);

    // Delete row 1 (A1) → B2's =A1 breaks to #REF!, =$A$1 also breaks.
    await apply([{ type: "row.delete", at: 0, count: 1 }]);
    expect(value(0, 1)).toMatchObject({ type: "#REF!" }); // B2 → B1, ref broken
    expect(value(1, 1)).toBeNull(); // no formula left at (1,1)
  });

  it("snapshot load rebuilds the graph and recalculates from source", async () => {
    const { api, wb, sheetId, value, apply } = setup();
    await apply([
      { type: "cell.set", range: "A1", value: 4 },
      { type: "formula.set", range: "B1", formula: "=A1+1" },
    ]);
    expect(value(0, 1)).toBe(5);

    // Save, corrupt the cached value, reload → graph rebuilt from source.
    const snapshot = api.getWorkbookSnapshot();
    snapshot.sheets[0]!.cells["0:1"] = { value: 999, formula: "=A1+1" };
    const api2 = createOpenSheet();
    const restored = api2.loadWorkbook(snapshot);
    expect(api2.readRange({ sheetId: restored.activeSheetId, range: "B1" })[0]![0]).toBe(5);

    // Recalc continues after load.
    await api2.applyOperations({
      workbookId: restored.id,
      sheetId: restored.activeSheetId,
      atomic: true,
      operations: [{ type: "cell.set", range: "A1", value: 10 }],
    });
    expect(api2.readRange({ sheetId: restored.activeSheetId, range: "B1" })[0]![0]).toBe(11);
    void wb;
  });
});

function sheetIdOf(api: ReturnType<typeof createOpenSheet>): string {
  return api.listSheets()[0]!.id;
}
