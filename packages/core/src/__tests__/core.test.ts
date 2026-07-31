import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { parseRange } from "@opensheet/shared";
import {
  ChunkedCellStore,
  NumberKeyCellStore,
  StringKeyCellStore,
  StyleTable,
  toWorkbookSnapshot,
  Workbook,
  workbookFromSnapshot,
  Worksheet,
} from "../index.js";

function makeSheet() {
  return new Worksheet({ id: "s1", name: "Sheet1", rowCount: 100, columnCount: 26 });
}

describe("sparse cell model", () => {
  it("does not allocate entries for empty cells", () => {
    const sheet = makeSheet();
    expect(sheet.cellCount).toBe(0);
    expect(sheet.getCell(50, 10)).toBeUndefined();
    sheet.setCell(50, 10, { value: 42 });
    expect(sheet.cellCount).toBe(1);
    sheet.deleteCell(50, 10);
    expect(sheet.cellCount).toBe(0);
  });

  it("all three stores behave identically", () => {
    for (const Store of [StringKeyCellStore, NumberKeyCellStore, ChunkedCellStore]) {
      const store = new Store();
      store.set(3, 4, { value: "x" });
      store.set(0, 0, { value: 1 });
      expect(store.size).toBe(2);
      expect(store.get(3, 4)).toEqual({ value: "x" });
      const seen: string[] = [];
      store.forEachInRange(parseRange("A1:Z100"), (r, c) => seen.push(`${r},${c}`));
      expect(seen.sort()).toEqual(["0,0", "3,4"]);
      const outOfRange: string[] = [];
      store.forEachInRange(parseRange("B10:B20"), (r, c) => outOfRange.push(`${r},${c}`));
      expect(outOfRange).toEqual([]);
    }
  });
});

describe("row/column structure", () => {
  it("insertRows shifts cells and row heights", () => {
    const sheet = makeSheet();
    sheet.setCell(0, 0, { value: "top" });
    sheet.setCell(5, 0, { value: "bottom" });
    sheet.rowHeights.set(5, 40);
    sheet.insertRows(1, 2);
    expect(sheet.getCell(0, 0)?.value).toBe("top");
    expect(sheet.getCell(5, 0)).toBeUndefined();
    expect(sheet.getCell(7, 0)?.value).toBe("bottom");
    expect(sheet.rowHeights.get(7)).toBe(40);
    expect(sheet.rowCount).toBe(102);
  });

  it("deleteRows removes cells in range and shifts up", () => {
    const sheet = makeSheet();
    sheet.setCell(1, 0, { value: "doomed" });
    sheet.setCell(3, 0, { value: "survivor" });
    sheet.deleteRows(0, 2);
    expect(sheet.getCell(1, 0)?.value).toBe("survivor");
    expect(sheet.cellCount).toBe(1);
  });

  it("insertColumns/deleteColumns mirror row behavior", () => {
    const sheet = makeSheet();
    sheet.setCell(0, 2, { value: "c" });
    sheet.insertColumns(1, 1);
    expect(sheet.getCell(0, 3)?.value).toBe("c");
    sheet.deleteColumns(1, 1);
    expect(sheet.getCell(0, 2)?.value).toBe("c");
  });
});

describe("style table", () => {
  it("deduplicates identical styles", () => {
    const table = new StyleTable();
    const a = table.register({ bold: true, fontSize: 12 });
    const b = table.register({ fontSize: 12, bold: true }); // key order differs
    const c = table.register({ bold: true, fontSize: 14 });
    expect(a).toBe(b);
    expect(a).not.toBe(c);
    expect(table.size).toBe(2);
  });

  it("toJSON returns a deep copy (callers cannot mutate internal styles)", () => {
    const table = new StyleTable();
    const id = table.register({ bold: true, border: { top: { style: "thin" } } });
    const json = table.toJSON();
    json[id]!.bold = false;
    json[id]!.border!.top!.style = "thick";
    expect(table.get(id)).toEqual({ bold: true, border: { top: { style: "thin" } } });
  });
});

describe("snapshot", () => {
  it("roundtrips cells, sizes, freeze, active sheet", () => {
    const workbook = new Workbook({ id: "wb1", name: "Book" });
    const sheet = makeSheet();
    sheet.setCell(0, 0, { value: "hello", styleId: "st1" });
    sheet.setCell(9, 3, { value: 42, formula: "=A1" });
    sheet.rowHeights.set(0, 30);
    sheet.columnWidths.set(3, 120);
    sheet.frozenRows = 1;
    sheet.frozenColumns = 2;
    workbook.addSheet(sheet);
    workbook.styles.register({ bold: true });

    const snapshot = toWorkbookSnapshot(workbook);
    const restored = workbookFromSnapshot(JSON.parse(JSON.stringify(snapshot)));

    expect(restored.id).toBe("wb1");
    expect(restored.activeSheetId).toBe("s1");
    const restoredSheet = restored.getSheet("s1");
    expect(restoredSheet.getCell(0, 0)).toEqual({ value: "hello", styleId: "st1" });
    expect(restoredSheet.getCell(9, 3)).toEqual({ value: 42, formula: "=A1" });
    expect(restoredSheet.rowHeights.get(0)).toBe(30);
    expect(restoredSheet.columnWidths.get(3)).toBe(120);
    expect(restoredSheet.frozenRows).toBe(1);
    expect(restoredSheet.frozenColumns).toBe(2);
    expect(restoredSheet.cellCount).toBe(2);
  });

  it("rejects unknown snapshot versions", () => {
    const workbook = new Workbook({ id: "wb", name: "B" });
    workbook.addSheet(makeSheet());
    const snapshot = toWorkbookSnapshot(workbook);
    expect(() => workbookFromSnapshot({ ...snapshot, version: 999 })).toThrow(/version/);
  });

  it("property: save/load is identity for random sparse sheets", () => {
    fc.assert(
      fc.property(
        fc.array(
          fc.record({
            row: fc.integer({ min: 0, max: 99 }),
            col: fc.integer({ min: 0, max: 25 }),
            value: fc.oneof(fc.string({ maxLength: 20 }), fc.integer(), fc.boolean(), fc.constant(null)),
          }),
          { maxLength: 200 },
        ),
        (cells) => {
          const workbook = new Workbook({ id: "wb", name: "P" });
          const sheet = makeSheet();
          for (const { row, col, value } of cells) sheet.setCell(row, col, { value });
          workbook.addSheet(sheet);
          const restored = workbookFromSnapshot(toWorkbookSnapshot(workbook));
          const a = toWorkbookSnapshot(workbook);
          const b = toWorkbookSnapshot(restored);
          return JSON.stringify(a) === JSON.stringify(b);
        },
      ),
    );
  });
});

describe("workbook event buffering", () => {
  it("merges events inside a batch, discards on rollback", () => {
    const workbook = new Workbook({ id: "wb", name: "E" });
    workbook.addSheet(makeSheet());
    const events: string[] = [];
    workbook.onChange((e) => events.push(`${e.sheetId}:${e.changes.length}:${e.batch}`));

    workbook.emit({
      workbookId: "wb",
      sheetId: "s1",
      changes: [{ range: parseRange("A1"), kind: "cells" }],
      source: "user",
      batch: false,
    });
    expect(events).toEqual(["s1:1:false"]);

    workbook.beginBatch();
    workbook.emit({
      workbookId: "wb",
      sheetId: "s1",
      changes: [{ range: parseRange("A1"), kind: "cells" }],
      source: "user",
      batch: false,
    });
    workbook.emit({
      workbookId: "wb",
      sheetId: "s1",
      changes: [{ range: parseRange("B2"), kind: "cells" }],
      source: "user",
      batch: false,
    });
    workbook.endBatch(true);
    expect(events).toEqual(["s1:1:false", "s1:2:true"]);

    workbook.beginBatch();
    workbook.emit({
      workbookId: "wb",
      sheetId: "s1",
      changes: [{ range: parseRange("C3"), kind: "cells" }],
      source: "user",
      batch: false,
    });
    workbook.endBatch(false);
    expect(events).toHaveLength(2); // nothing more delivered
  });
});
