import { describe, expect, it, vi } from "vitest";
import { CommandBus, createDefaultRegistry, type JournalBatch } from "@injoysai/opensheet-commands";
import { Workbook, Worksheet } from "@injoysai/opensheet-core";
import { HistoryManager } from "../index.js";

function setup(historyOptions?: ConstructorParameters<typeof HistoryManager>[0]) {
  const workbook = new Workbook({ id: "wb", name: "B" });
  workbook.addSheet(new Worksheet({ id: "s1", name: "S1", rowCount: 100, columnCount: 26 }));
  const history = new HistoryManager(historyOptions);
  const bus = new CommandBus(workbook, { history, registry: createDefaultRegistry() });
  return { workbook, history, bus };
}

function fakeBatch(approxBytes: number): JournalBatch {
  return {
    entries: [
      {
        label: "fake",
        affected: [],
        approxBytes,
        undo: () => {},
        redo: () => {},
      },
    ],
    source: "user",
    approxBytes,
  };
}

describe("HistoryManager", () => {
  it("undo/redo roundtrip through the bus", () => {
    const { workbook, history, bus } = setup();
    bus.execute("cell.set", { range: "A1", value: "x" }, { sheetId: "s1" });
    expect(history.canUndo).toBe(true);

    expect(history.undo(bus)).toBe(true);
    expect(workbook.getSheet("s1").getCell(0, 0)).toBeUndefined();
    expect(history.canRedo).toBe(true);

    expect(history.redo(bus)).toBe(true);
    expect(workbook.getSheet("s1").getCell(0, 0)?.value).toBe("x");
  });

  it("batch is a single history entry", () => {
    const { history, bus } = setup();
    bus.applyOperations({
      sheetId: "s1",
      operations: [
        { type: "cell.set", range: "A1", value: 1 },
        { type: "cell.set", range: "A2", value: 2 },
        { type: "cell.set", range: "A3", value: 3 },
      ],
      atomic: true,
    });
    expect(history.undoDepth).toBe(1);
  });

  it("new action clears the redo stack", () => {
    const { history, bus } = setup();
    bus.execute("cell.set", { range: "A1", value: 1 }, { sheetId: "s1" });
    history.undo(bus);
    expect(history.canRedo).toBe(true);
    bus.execute("cell.set", { range: "A2", value: 2 }, { sheetId: "s1" });
    expect(history.canRedo).toBe(false);
  });

  it("enforces maxEntries", () => {
    const { history } = setup({ maxEntries: 3 });
    for (let i = 0; i < 10; i++) history.push(fakeBatch(100));
    expect(history.undoDepth).toBe(3);
  });

  it("enforces maxMemoryBytes and reports evictions", () => {
    const onEvict = vi.fn();
    const { history } = setup({ maxEntries: 100, maxMemoryBytes: 1000, onEvict });
    for (let i = 0; i < 10; i++) history.push(fakeBatch(300));
    expect(history.undoDepth).toBeLessThanOrEqual(3);
    expect(history.retainedBytes).toBeLessThanOrEqual(1000);
    expect(onEvict).toHaveBeenCalled();
  });

  it("undo returns false on empty stack", () => {
    const { history, bus } = setup();
    expect(history.undo(bus)).toBe(false);
    expect(history.redo(bus)).toBe(false);
  });

  it("failed undo replay keeps both stacks and byte accounting intact", () => {
    const workbook = new Workbook({ id: "wb", name: "B" });
    workbook.addSheet(new Worksheet({ id: "s1", name: "S1", rowCount: 10, columnCount: 5 }));
    const history = new HistoryManager();
    const bus = new CommandBus(workbook, { history, registry: createDefaultRegistry() });

    history.push({
      entries: [
        {
          label: "exploding",
          affected: [],
          approxBytes: 100,
          undo: () => {
            throw new Error("replay exploded");
          },
          redo: () => {},
        },
      ],
      source: "user",
      approxBytes: 100,
    });
    const bytesBefore = history.retainedBytes;

    expect(() => history.undo(bus)).toThrow("replay exploded");
    expect(history.undoDepth).toBe(1); // entry NOT lost
    expect(history.canRedo).toBe(false); // and NOT moved to redo
    expect(history.retainedBytes).toBe(bytesBefore);
  });
});
