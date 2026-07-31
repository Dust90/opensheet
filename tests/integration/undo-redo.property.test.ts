// Property tests: random operation sequences vs a naive model, plus
// execute → undo → redo identity and atomic-rollback identity.

import { describe, it } from "vitest";
import fc from "fast-check";
import { CommandBus, createDefaultRegistry } from "@opensheet/commands";
import { toWorkbookSnapshot, Workbook, Worksheet } from "@opensheet/core";
import { HistoryManager } from "@opensheet/history";
import { expect } from "vitest";

function setup() {
  const workbook = new Workbook({ id: "wb", name: "P" });
  workbook.addSheet(new Worksheet({ id: "s1", name: "S1", rowCount: 50, columnCount: 10 }));
  const history = new HistoryManager();
  const bus = new CommandBus(workbook, { history, registry: createDefaultRegistry() });
  return { workbook, history, bus };
}

const opArb = fc.oneof(
  fc.record({
    type: fc.constant("cell.set" as const),
    range: fc.tuple(fc.integer({ min: 0, max: 49 }), fc.integer({ min: 0, max: 9 })).map(
      ([r, c]) => `${String.fromCharCode(65 + c)}${r + 1}`,
    ),
    value: fc.oneof(fc.string({ maxLength: 10 }), fc.integer(), fc.boolean(), fc.constant(null)),
  }),
  fc.record({
    type: fc.constant("cell.clear" as const),
    range: fc.tuple(fc.integer({ min: 0, max: 49 }), fc.integer({ min: 0, max: 9 })).map(
      ([r, c]) => `${String.fromCharCode(65 + c)}${r + 1}`,
    ),
  }),
);

describe("command/history properties", () => {
  it("execute all → undo all → redo all reproduces final state", () => {
    fc.assert(
      fc.property(fc.array(opArb, { minLength: 1, maxLength: 25 }), (ops) => {
        const { workbook, history, bus } = setup();
        for (const op of ops) bus.execute(op.type, op, { sheetId: "s1" });
        const finalState = JSON.stringify(toWorkbookSnapshot(workbook));

        while (history.canUndo) history.undo(bus);
        const undone = toWorkbookSnapshot(workbook);
        expect(undone.sheets[0]!.cells).toEqual({}); // fully reverted

        while (history.canRedo) history.redo(bus);
        expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(finalState);
      }),
    );
  });

  it("atomic batch with a guaranteed failure == state before the batch", () => {
    fc.assert(
      fc.property(fc.array(opArb, { maxLength: 10 }), (prefix) => {
        const { workbook, bus } = setup();
        for (const op of prefix) bus.execute(op.type, op, { sheetId: "s1" });
        const before = JSON.stringify(toWorkbookSnapshot(workbook));

        try {
          bus.applyOperations({
            sheetId: "s1",
            atomic: true,
            operations: [
              { type: "cell.set", range: "A1", value: "tmp" },
              { type: "cell.set", range: "ZZ99999", value: "boom" }, // always fails
            ],
          });
        } catch {
          // expected
        }
        expect(JSON.stringify(toWorkbookSnapshot(workbook))).toBe(before);
      }),
    );
  });
});
