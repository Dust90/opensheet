// Integration: Command Bus → Core → ChangeEvent, transaction boundaries.

import { describe, expect, it } from "vitest";
import {
  ApplyOperationsError,
  CommandBus,
  createDefaultRegistry,
} from "@opensheet/commands";
import { toWorkbookSnapshot, Workbook, Worksheet } from "@opensheet/core";
import { HistoryManager } from "@opensheet/history";
import type { ChangeEvent } from "@opensheet/shared";

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
  it("hook output merges into the same single event and never enters history", () => {
    const { workbook, history, bus, events } = setup();
    // Simulates the M3 formula engine: recompute B1 = A1 before commit.
    bus.addBeforeCommitHook(({ workbook: wb }) => {
      const sheet = wb.getSheet("s1");
      const source = sheet.getCell(0, 0)?.value;
      if (typeof source === "number") {
        sheet.setCell(0, 1, { value: source * 2, formula: "=A1*2" });
        wb.emit({
          workbookId: wb.id,
          sheetId: "s1",
          changes: [{ range: { startRow: 0, startCol: 1, endRow: 0, endCol: 1 }, kind: "cells" }],
          source: "derived",
          batch: false,
        });
      }
    });

    bus.execute("cell.set", { range: "A1", value: 21 }, { sheetId: "s1" });

    expect(workbook.getSheet("s1").getCell(0, 1)?.value).toBe(42);
    // One user event + one derived event (merged per source) — no intermediate emissions.
    expect(events).toHaveLength(2);
    expect(events.map((e) => e.source).sort()).toEqual(["derived", "user"]);
    expect(history.undoDepth).toBe(1); // derived never recorded
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
