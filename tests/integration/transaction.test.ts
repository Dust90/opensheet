// Integration: Command Bus → Core → ChangeEvent, transaction boundaries.

import { describe, expect, it } from "vitest";
import {
  ApplyOperationsError,
  CommandBus,
  createDefaultRegistry,
} from "@injoysai/opensheet-commands";
import { toWorkbookSnapshot, Workbook, Worksheet } from "@injoysai/opensheet-core";
import { HistoryManager } from "@injoysai/opensheet-history";
import type { ChangeEvent } from "@injoysai/opensheet-shared";

function setup() {
  const workbook = new Workbook({ id: "wb", name: "B" });
  workbook.addSheet(new Worksheet({ id: "s1", name: "S1", rowCount: 100, columnCount: 26 }));
  const history = new HistoryManager();
  const bus = new CommandBus(workbook, { history, registry: createDefaultRegistry() });
  const events: ChangeEvent[] = [];
  workbook.onChange((e) => events.push(e));
  return { workbook, history, bus, events };
}

describe("atomic transactions", () => {
  it("success: observers receive exactly ONE merged event and history has ONE entry", () => {
    const { workbook, history, bus, events } = setup();
    const result = bus.applyOperations({
      sheetId: "s1",
      atomic: true,
      operations: [
        { type: "cell.set", range: "A1", value: "a" },
        { type: "cell.set", range: "B2", value: "b" },
        { type: "cell.set", range: "C3", value: "c" },
      ],
    });
    expect(result.status).toBe("completed");
    expect(events).toHaveLength(1);
    expect(events[0]!.batch).toBe(true);
    expect(events[0]!.changes).toHaveLength(3);
    expect(events[0]!.source).toBe("api");
    expect(history.undoDepth).toBe(1);
    expect(workbook.getSheet("s1").cellCount).toBe(3);
  });

  it("failure on 3rd op: full rollback, zero events, zero history, workbook identical to before", () => {
    const { workbook, history, bus, events } = setup();
    bus.execute("cell.set", { range: "A1", value: "keep" }, { sheetId: "s1" });
    events.length = 0;
    const historyDepthBefore = history.undoDepth;
    const before = JSON.stringify(toWorkbookSnapshot(workbook));

    try {
      bus.applyOperations({
        sheetId: "s1",
        atomic: true,
        operations: [
          { type: "cell.set", range: "C1", value: 1 },
          { type: "cell.set", range: "C2", value: 2 },
          { type: "cell.set", range: "ZZ100000", value: 3 }, // out of bounds → throws
        ],
      });
      expect.unreachable("batch should have failed");
    } catch (error) {
      expect(error).toBeInstanceOf(ApplyOperationsError);
      const applyError = error as ApplyOperationsError;
      expect(applyError.failedOperationIndex).toBe(2);
      expect(applyError.errorCode).toBe("E_INVALID_RANGE");
    }

    expect(events).toHaveLength(0); // no intermediate state leaked
    expect(history.undoDepth).toBe(historyDepthBefore); // no history entry
    expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(before); // deep equality
  });

  it("non-atomic: earlier ops survive a later failure", () => {
    const { workbook, bus } = setup();
    try {
      bus.applyOperations({
        sheetId: "s1",
        atomic: false,
        operations: [
          { type: "cell.set", range: "A1", value: 1 },
          { type: "cell.set", range: "ZZ100000", value: 2 },
        ],
      });
      expect.unreachable();
    } catch (error) {
      expect((error as ApplyOperationsError).failedOperationIndex).toBe(1);
    }
    expect(workbook.getSheet("s1").getCell(0, 0)?.value).toBe(1);
  });
});

describe("derived (beforeCommit) channel", () => {
  it("hook output merges into the same commit and never enters history", () => {
    const { workbook, history, bus, events } = setup();
    // Simulates the M3 formula engine: recompute B1 = A1*2 before commit.
    // The hook receives a READ-ONLY WorkbookView (guardrail 3): reads via
    // getSheetView, writes ONLY via derived.setComputedValue (M3.0).
    bus.addBeforeCommitHook(({ workbook: wb, derived }) => {
      const source = wb.getSheetView("s1").getCell(0, 0)?.value;
      if (typeof source === "number") {
        derived.setComputedValue("s1", 0, 1, source * 2);
      }
    });

    bus.execute("cell.set", { range: "A1", value: 21 }, { sheetId: "s1" });

    expect(workbook.getSheet("s1").getCell(0, 1)?.value).toBe(42);
    // One user event + one derived event (merged per sheet+source).
    expect(events).toHaveLength(2);
    expect(events.map((e) => e.source).sort()).toEqual(["derived", "user"]);
    expect(history.undoDepth).toBe(1); // derived never recorded
  });

  it("M3.0: hook receives the transaction's pending changes", () => {
    const { bus } = setup();
    const seen: Array<{ sheetId: string; range: string; kind: string }> = [];
    bus.addBeforeCommitHook(({ changes }) => {
      for (const c of changes) {
        seen.push({
          sheetId: c.sheetId,
          range: `${c.range.startRow}:${c.range.startCol}-${c.range.endRow}:${c.range.endCol}`,
          kind: c.kind,
        });
      }
    });

    bus.applyOperations({
      sheetId: "s1",
      atomic: true,
      operations: [
        { type: "range.write", range: "A1:B2", values: [["a", 1], ["b", 2]] },
        { type: "range.style", range: "A1", style: { bold: true } },
      ],
    });
    expect(seen).toContainEqual({ sheetId: "s1", range: "0:0-1:1", kind: "cells" });
    expect(seen).toContainEqual({ sheetId: "s1", range: "0:0-0:0", kind: "style" });
  });

  it("M3.0: setComputedValue preserves formula/styleId and rejects out-of-bounds", () => {
    const { workbook, bus } = setup();
    // Give B1 a formula + style, then a hook recomputes its value.
    bus.applyOperations({
      sheetId: "s1",
      atomic: true,
      operations: [
        { type: "range.style", range: "B1", style: { bold: true } },
        { type: "cell.set", range: "B1", value: 0 },
      ],
    });
    const styleId = workbook.getSheet("s1").getCell(0, 1)!.styleId!;
    bus.addBeforeCommitHook(({ derived }) => {
      derived.setComputedValue("s1", 0, 1, 7);
    });
    bus.execute("cell.set", { range: "A1", value: 1 }, { sheetId: "s1" });

    const b1 = workbook.getSheet("s1").getCell(0, 1)!;
    expect(b1.value).toBe(7);
    expect(b1.styleId).toBe(styleId); // metadata preserved
    expect(b1.formula).toBeUndefined(); // hook did not set one — stays absent

    // Out-of-bounds derived write must throw (and roll back the transaction).
    bus.addBeforeCommitHook(({ derived }) => {
      derived.setComputedValue("s1", 9999, 0, 1);
    });
    expect(() =>
      bus.execute("cell.set", { range: "A1", value: 2 }, { sheetId: "s1" }),
    ).toThrow(/out of bounds/);
    expect(workbook.getSheet("s1").getCell(0, 1)!.value).toBe(7); // rolled back
  });

  it("hook is NOT called on failed transactions (no partial derived state)", () => {
    const { bus, events } = setup();
    let hookCalls = 0;
    bus.addBeforeCommitHook(() => hookCalls++);
    try {
      bus.applyOperations({
        sheetId: "s1",
        atomic: true,
        operations: [
          { type: "cell.set", range: "A1", value: 1 },
          { type: "cell.set", range: "ZZ100000", value: 2 },
        ],
      });
    } catch {
      // expected
    }
    expect(hookCalls).toBe(0);
    expect(events).toHaveLength(0);
  });

  it("hook throws AFTER writing derived changes: everything rolls back (commands + derived)", () => {
    const { workbook, history, bus, events } = setup();
    bus.execute("cell.set", { range: "A1", value: "keep" }, { sheetId: "s1" });
    events.length = 0;
    const historyDepthBefore = history.undoDepth;
    const before = JSON.stringify(toWorkbookSnapshot(workbook));

    bus.addBeforeCommitHook(({ derived }) => {
      derived.setComputedValue("s1", 0, 1, 999); // derived write happens first...
      throw new Error("recalc exploded"); // ...then the hook fails
    });

    try {
      bus.applyOperations({
        sheetId: "s1",
        atomic: true,
        operations: [
          { type: "cell.set", range: "C1", value: 1 },
          { type: "cell.set", range: "C2", value: 2 },
        ],
      });
      expect.unreachable("transaction should have failed in the hook");
    } catch {
      // expected
    }

    expect(events).toHaveLength(0); // nothing leaked
    expect(history.undoDepth).toBe(historyDepthBefore); // no history
    expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(before); // derived write rolled back too
  });
});

describe("observer failure isolation", () => {
  it("a throwing listener does NOT affect commit, history, or later commands", () => {
    const { workbook, history, bus } = setup();
    const errors: unknown[] = [];
    workbook.onListenerError = (error) => errors.push(error);
    workbook.onChange(() => {
      throw new Error("renderer exploded");
    });

    // Command still succeeds, data is committed, history is written.
    bus.execute("cell.set", { range: "A1", value: "v" }, { sheetId: "s1" });
    expect(workbook.getSheet("s1").getCell(0, 0)?.value).toBe("v");
    expect(history.undoDepth).toBe(1);
    expect(errors).toHaveLength(1);

    // Batch state is clean: later commands and undo still work.
    bus.execute("cell.set", { range: "A2", value: "w" }, { sheetId: "s1" });
    expect(workbook.getSheet("s1").getCell(1, 0)?.value).toBe("w");
    history.undo(bus);
    expect(workbook.getSheet("s1").getCell(1, 0)).toBeUndefined();
    expect(errors.length).toBeGreaterThanOrEqual(3); // each event delivery isolated
  });
});

describe("sheet lifecycle", () => {
  it("sheet.create undo/redo restores the previously active sheet", () => {
    const { workbook, history, bus } = setup();
    bus.execute("sheet.create", { name: "Second", rows: 50, columns: 10 }, { source: "user" });
    const secondId = workbook.activeSheetId;
    expect(workbook.listSheets().map((s) => s.name)).toEqual(["S1", "Second"]);

    history.undo(bus);
    expect(workbook.listSheets().map((s) => s.name)).toEqual(["S1"]);
    expect(workbook.activeSheetId).toBe("s1"); // restored, not just "first sheet"

    history.redo(bus);
    expect(workbook.listSheets().map((s) => s.name)).toEqual(["S1", "Second"]);
    expect(workbook.activeSheetId).toBe(secondId);
  });
});

describe("command → core → event → dirty ranges pipeline", () => {
  it("emits precise dirty ranges for renderer consumption", () => {
    const { bus, events } = setup();
    bus.execute("range.write", { range: "B2:C3", values: [[1, 2], [3, 4]] }, { sheetId: "s1" });
    expect(events).toHaveLength(1);
    expect(events[0]!.changes).toEqual([
      { range: { startRow: 1, startCol: 1, endRow: 2, endCol: 2 }, kind: "cells" },
    ]);
  });

  it("range.write redo uses a cloned payload (caller mutation cannot corrupt redo)", () => {
    const { workbook, history, bus } = setup();
    const values = [["a", "b"]];
    bus.execute("range.write", { range: "A1:B1", values }, { sheetId: "s1" });
    values[0]![0] = "MUTATED"; // caller mutates after execution

    history.undo(bus);
    expect(workbook.getSheet("s1").getCell(0, 0)).toBeUndefined();

    history.redo(bus);
    expect(workbook.getSheet("s1").getCell(0, 0)?.value).toBe("a"); // not "MUTATED"
  });

  it("undo/redo replay through journal restores state and emits source=undo/redo", () => {
    const { workbook, history, bus, events } = setup();
    bus.execute("cell.set", { range: "A1", value: "v1" }, { sheetId: "s1" });
    bus.execute("cell.set", { range: "A1", value: "v2" }, { sheetId: "s1" });
    events.length = 0;

    history.undo(bus);
    expect(workbook.getSheet("s1").getCell(0, 0)?.value).toBe("v1");
    history.undo(bus);
    expect(workbook.getSheet("s1").getCell(0, 0)).toBeUndefined();
    history.redo(bus);
    expect(workbook.getSheet("s1").getCell(0, 0)?.value).toBe("v1");

    expect(events.map((e) => e.source)).toEqual(["undo", "undo", "redo"]);
  });
});
