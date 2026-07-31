// M2 integration: row/column structure + style commands with full inverse
// journal round-trips through the Command Bus and History.

import { describe, expect, it } from "vitest";
import { Workbook, Worksheet } from "@opensheet/core";
import { CommandBus, createDefaultRegistry } from "@opensheet/commands";
import { HistoryManager } from "@opensheet/history";

function makeWorkbook() {
  const workbook = new Workbook({ id: "wb", name: "wb" });
  const sheet = new Worksheet({ id: "s1", name: "Sheet1", rowCount: 100, columnCount: 10 });
  workbook.addSheet(sheet);
  const history = new HistoryManager();
  const bus = new CommandBus(workbook, { history, registry: createDefaultRegistry() });
  return { workbook, sheet, bus, history };
}

describe("row/column structure commands", () => {
  it("row.delete journal restores cells, row heights and freeze on undo", () => {
    const { workbook, sheet, bus, history } = makeWorkbook();
    // Seed data around the deletion zone.
    bus.applyOperations({
      sheetId: "s1",
      operations: [
        { type: "range.write", range: "A1:C3", values: [["a", 1, "x"], ["b", 2, "y"], ["c", 3, "z"]] },
        { type: "row.insert", at: 2, count: 1 }, // blank row 3 (index 2)
      ],
      atomic: true,
      source: "user",
    });
    // Custom row height inside the to-be-deleted zone + freeze row 1.
    sheet.rowHeights.set(1, 42);
    sheet.frozenRows = 1;
    history.clear();

    bus.applyOperations({
      sheetId: "s1",
      operations: [{ type: "row.delete", at: 1, count: 2 }],
      atomic: true,
      source: "user",
    });
    // rows 1-2 (index) removed; A1 "a" remains, old row 3 (index 2+1) moved up.
    // (100 base rows + 1 inserted = 101 before the delete.)
    expect(sheet.rowCount).toBe(99);
    expect(sheet.getCell(0, 0)?.value).toBe("a");
    expect(sheet.getCell(0, 1)?.value).toBe(1);
    // Frozen row is index 0 (frozenRows=1); deleting index 1-2 leaves it intact.
    expect(sheet.frozenRows).toBe(1);
    expect(sheet.rowHeights.has(1)).toBe(false);

    history.undo(bus);
    expect(sheet.rowCount).toBe(101);
    expect(sheet.getCell(0, 0)?.value).toBe("a");
    // Old rows restored at their original coordinates.
    expect(sheet.getCell(1, 0)?.value).toBe("b");
    expect(sheet.getCell(1, 1)?.value).toBe(2);
    expect(sheet.getCell(3, 0)?.value).toBe("c");
    expect(sheet.rowHeights.get(1)).toBe(42);
    expect(sheet.frozenRows).toBe(1);

    history.redo(bus);
    expect(sheet.rowCount).toBe(99);
    expect(sheet.getCell(0, 0)?.value).toBe("a");
    expect(sheet.frozenRows).toBe(1);
  });

  it("column.insert/delete round-trip preserves data and column widths", () => {
    const { workbook, sheet, bus, history } = makeWorkbook();
    bus.applyOperations({
      sheetId: "s1",
      operations: [
        { type: "range.write", range: "A1:C2", values: [["a", "b", "c"], ["d", "e", "f"]] },
      ],
      atomic: true,
      source: "user",
    });
    sheet.columnWidths.set(1, 300);
    history.clear();

    bus.applyOperations({
      sheetId: "s1",
      operations: [{ type: "column.insert", at: 1, count: 1 }],
      atomic: true,
      source: "user",
    });
    expect(sheet.columnCount).toBe(11);
    expect(sheet.getCell(0, 1)?.value).toBeUndefined(); // blank inserted col
    expect(sheet.getCell(0, 2)?.value).toBe("b"); // shifted
    expect(sheet.columnWidths.get(2)).toBe(300); // width shifted with the col

    history.undo(bus);
    expect(sheet.columnCount).toBe(10);
    expect(sheet.getCell(0, 0)?.value).toBe("a");
    expect(sheet.getCell(0, 1)?.value).toBe("b");
    expect(sheet.columnWidths.get(1)).toBe(300);

    bus.applyOperations({
      sheetId: "s1",
      operations: [{ type: "column.delete", at: 0, count: 2 }],
      atomic: true,
      source: "user",
    });
    expect(sheet.columnCount).toBe(8);
    expect(sheet.getCell(0, 0)?.value).toBe("c");
    history.undo(bus);
    expect(sheet.columnCount).toBe(10);
    expect(sheet.getCell(0, 0)?.value).toBe("a");
    expect(sheet.getCell(0, 1)?.value).toBe("b");
  });

  it("rejects out-of-bounds structure operations atomically", () => {
    const { sheet, bus } = makeWorkbook();
    expect(() =>
      bus.applyOperations({
        sheetId: "s1",
        operations: [{ type: "row.delete", at: 95, count: 10 }],
        atomic: true,
        source: "user",
      }),
    ).toThrow(/exceeds sheet rows/);
    expect(sheet.rowCount).toBe(100); // nothing changed
    expect(() =>
      bus.applyOperations({
        sheetId: "s1",
        operations: [{ type: "row.insert", at: 101 }],
        atomic: true,
        source: "user",
      }),
    ).toThrow(/exceeds sheet rows/);
  });
});

describe("range.style command", () => {
  it("merges style attributes, dedupes via StyleTable, undo restores styleId", () => {
    const { workbook, sheet, bus, history } = makeWorkbook();
    bus.applyOperations({
      sheetId: "s1",
      operations: [
        { type: "range.write", range: "A1:B1", values: [["x", "y"]] },
        { type: "range.style", range: "A1:B1", style: { bold: true } },
      ],
      atomic: true,
      source: "user",
    });
    const id1 = sheet.getCell(0, 0)?.styleId;
    const id2 = sheet.getCell(0, 1)?.styleId;
    expect(id1).toBeDefined();
    expect(id2).toBe(id1); // identical styles deduplicated
    expect(workbook.styles.get(id1!)?.bold).toBe(true);
    history.clear();

    // Add italic to A1 only: merged style, new id; B1 keeps bold-only id.
    bus.applyOperations({
      sheetId: "s1",
      operations: [{ type: "range.style", range: "A1", style: { italic: true } }],
      atomic: true,
      source: "user",
    });
    const boldItalic = sheet.getCell(0, 0)?.styleId;
    expect(workbook.styles.get(boldItalic!)?.bold).toBe(true);
    expect(workbook.styles.get(boldItalic!)?.italic).toBe(true);
    expect(sheet.getCell(0, 1)?.styleId).toBe(id1);

    history.undo(bus);
    expect(sheet.getCell(0, 0)?.styleId).toBe(id1);
    expect(workbook.styles.get(sheet.getCell(0, 0)!.styleId!)?.italic).toBeUndefined();

    history.redo(bus);
    expect(sheet.getCell(0, 0)?.styleId).toBe(boldItalic);
  });

  it("styles a previously empty cell and undo removes it again", () => {
    const { sheet, bus, history } = makeWorkbook();
    bus.applyOperations({
      sheetId: "s1",
      operations: [{ type: "range.style", range: "C5", style: { backgroundColor: "#ff0000" } }],
      atomic: true,
      source: "user",
    });
    const cell = sheet.getCell(4, 2);
    expect(cell?.styleId).toBeDefined();
    expect(cell?.value).toBeNull();

    history.undo(bus);
    expect(sheet.getCell(4, 2)).toBeUndefined();
  });

  it("rejects empty style payloads", () => {
    const { bus } = makeWorkbook();
    expect(() =>
      bus.applyOperations({
        sheetId: "s1",
        operations: [{ type: "range.style", range: "A1", style: {} }],
        atomic: true,
        source: "user",
      }),
    ).toThrow(/at least one style attribute/);
  });
});
