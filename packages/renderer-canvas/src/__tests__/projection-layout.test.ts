// M4.1 renderer regression: the exact compositions SheetGrid wires together
// (projection + visual row axis + viewport + hit test + dirty rects), tested
// in pure node. Grid DOM paths stay covered by the E2E suite in identity mode.
//
// Scenario: 20 rows × 3 cols, filter range 5..14, visible [5,7,9,12].

import { describe, expect, it } from "vitest";
import { SheetError, type CellData, type Range } from "@opensheet/shared";
import type { WorksheetView } from "@opensheet/core";
import { AxisMetrics } from "../axis-metrics.js";
import { cellRectInCanvas, hitTestCell } from "../coordinate-mapper.js";
import { DirtyRegionTracker } from "../dirty-region.js";
import {
  FilteredRowProjection,
  IdentityRowProjection,
  lastVisiblePhysicalRow,
  physicalRangeToVisualRange,
  relocateToVisibleRow,
  type RowProjection,
} from "../row-projection.js";
import { clampScroll, computeScrollToCell, computeViewport } from "../viewport.js";

const PHYSICAL = 20;
const COLS = 3;
const HEADER_W = 48;
const HEADER_H = 26;

function fakeSheet(frozenRows = 0): WorksheetView {
  const cells = new Map<string, CellData>();
  for (let r = 0; r < PHYSICAL; r++) {
    for (let c = 0; c < COLS; c++) cells.set(`${r}:${c}`, { value: `R${r + 1}C${c + 1}` });
  }
  return {
    id: "s1",
    name: "S",
    rowCount: PHYSICAL,
    columnCount: COLS,
    frozenRows,
    frozenColumns: 0,
    cellCount: cells.size,
    getCell: (r, c) => cells.get(`${r}:${c}`),
    *cellEntries() {
      for (const [key, data] of cells) {
        const [r, c] = key.split(":").map(Number);
        yield [r!, c!, data];
      }
    },
    forEachCellInRange(range: Range, cb: (r: number, c: number, d: Readonly<CellData>) => void) {
      for (let r = range.startRow; r <= range.endRow; r++)
        for (let c = range.startCol; c <= range.endCol; c++) {
          const d = cells.get(`${r}:${c}`);
          if (d !== undefined) cb(r, c, d);
        }
    },
    getRowHeight: () => undefined,
    getColumnWidth: () => undefined,
  };
}

function makeProjection(): FilteredRowProjection {
  return new FilteredRowProjection(PHYSICAL, { startRow: 5, endRow: 14 }, [5, 7, 9, 12]);
}

/** Same construction as SheetGrid.rebuildMetrics. */
function visualAxis(sheet: WorksheetView, projection: RowProjection): AxisMetrics {
  return new AxisMetrics(projection.visualRowCount, 24, (visualRow) =>
    sheet.getRowHeight(projection.visualToPhysical(visualRow)),
  );
}

function colAxis(sheet: WorksheetView): AxisMetrics {
  return new AxisMetrics(sheet.columnCount, 100, (c) => sheet.getColumnWidth(c));
}

function layoutFor(rows: AxisMetrics, cols: AxisMetrics, projection: RowProjection, sheet: WorksheetView, scrollY = 0) {
  return computeViewport({
    scrollX: 0,
    scrollY,
    width: 800,
    height: 600,
    rows,
    cols,
    frozenRowCount: projection.visibleCountBefore(sheet.frozenRows),
    frozenColCount: sheet.frozenColumns,
    bufferPx: 0,
    headerWidth: HEADER_W,
    headerHeight: HEADER_H,
  });
}

describe("hit test → physical row", () => {
  it("clicking the 3rd VISIBLE row of the filtered area hits the right physical row", () => {
    const sheet = fakeSheet();
    const p = makeProjection();
    const rows = visualAxis(sheet, p);
    const cols = colAxis(sheet);
    const layout = layoutFor(rows, cols, p, sheet);
    // Visual row 6 = 3rd visible row inside the filter range (5,7,9).
    const y = HEADER_H + rows.positionOf(6) + 2;
    const hit = hitTestCell({
      x: HEADER_W + 5,
      y,
      layout,
      rows,
      cols,
      scrollX: 0,
      scrollY: 0,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
      scrollbarSize: 10,
      rowCount: rows.length,
      colCount: sheet.columnCount,
    });
    expect(hit.zone).toBe("cell");
    // The returned row is VISUAL; the grid converts once → physical 7.
    expect(hit.row).toBe(6);
    expect(p.visualToPhysical(hit.row)).toBe(7);
  });
});

describe("paint + header mapping", () => {
  it("cells painted for the viewport come from the correct physical rows", () => {
    const sheet = fakeSheet();
    const p = makeProjection();
    const rows = visualAxis(sheet, p);
    const cols = colAxis(sheet);
    const layout = layoutFor(rows, cols, p, sheet);
    // SheetGrid.paintQuadrantCells composition: visual loop → physical data.
    const painted: string[] = [];
    for (let visualRow = layout.main.rowStart; visualRow <= layout.main.rowEnd; visualRow++) {
      const physicalRow = p.visualToPhysical(visualRow);
      painted.push(String(sheet.getCell(physicalRow, 0)?.value));
    }
    expect(painted).toEqual(["R1C1", "R2C1", "R3C1", "R4C1", "R5C1", "R6C1", "R8C1", "R10C1", "R13C1", "R16C1", "R17C1", "R18C1", "R19C1", "R20C1"]);
  });

  it("row headers display PHYSICAL numbers, not compacted visual indices", () => {
    const p = makeProjection();
    const labels: string[] = [];
    for (let visualRow = 0; visualRow < p.visualRowCount; visualRow++) {
      labels.push(String(p.visualToPhysical(visualRow) + 1));
    }
    expect(labels).toEqual(["1", "2", "3", "4", "5", "6", "8", "10", "13", "16", "17", "18", "19", "20"]);
  });
});

describe("keyboard navigation over visible rows", () => {
  const p = makeProjection();

  it("ArrowDown skips hidden rows (moveActiveRow composition)", () => {
    expect(p.nextVisible(5, 1)).toBe(7);
    expect(p.nextVisible(7, 1)).toBe(9);
    expect(p.nextVisible(12, 1)).toBe(15);
    // At the bottom edge the selection stays (grid: ?? baseRow).
    expect(p.nextVisible(19, 1)).toBeUndefined();
  });

  it("ArrowUp skips hidden rows", () => {
    expect(p.nextVisible(15, -1)).toBe(12);
    expect(p.nextVisible(7, -1)).toBe(5);
  });

  it("PageDown moves ±pageRows on the VISUAL axis", () => {
    const pageRows = 3;
    const fromPhysical = 5;
    const baseVisual = p.physicalToVisual(fromPhysical)!;
    const target = p.visualToPhysical(Math.min(baseVisual + pageRows, p.visualRowCount - 1));
    expect(target).toBe(12); // visual 5 + 3 = visual 8 → physical 12
  });

  it("Ctrl+End targets the last VISIBLE physical row", () => {
    expect(lastVisiblePhysicalRow(p)).toBe(19);
    const hiddenTail = new FilteredRowProjection(PHYSICAL, { startRow: 15, endRow: 19 }, []);
    expect(lastVisiblePhysicalRow(hiddenTail)).toBe(14);
  });

  it("hidden active cell relocates to the next visible physical row", () => {
    expect(relocateToVisibleRow(p, 6)).toBe(7);
    expect(relocateToVisibleRow(p, 14)).toBe(15);
    expect(relocateToVisibleRow(p, 4)).toBe(4); // already visible: unchanged
  });
});

describe("frozen zone with hidden rows", () => {
  it("frozenHeight covers only the VISIBLE frozen rows", () => {
    const sheet = fakeSheet(8); // physical rows 0..7 frozen; row 6 hidden
    const p = makeProjection();
    expect(p.visibleCountBefore(8)).toBe(7);
    const rows = visualAxis(sheet, p);
    const cols = colAxis(sheet);
    const layout = layoutFor(rows, cols, p, sheet);
    expect(layout.frozenHeight).toBe(7 * 24);
    expect(layout.top!.rowStart).toBe(0);
    expect(layout.top!.rowEnd).toBe(6);
    // The frozen strip shows physical 0,1,2,3,4,5,7 — row 6 is not pinned.
    const frozenPhysical: number[] = [];
    for (let v = layout.top!.rowStart; v <= layout.top!.rowEnd; v++) {
      frozenPhysical.push(p.visualToPhysical(v));
    }
    expect(frozenPhysical).toEqual([0, 1, 2, 3, 4, 5, 7]);
  });
});

describe("editor overlay position", () => {
  it("cellRectInCanvas uses the VISUAL row of the edited physical cell", () => {
    const sheet = fakeSheet();
    const p = makeProjection();
    const rows = visualAxis(sheet, p);
    const cols = colAxis(sheet);
    const layout = layoutFor(rows, cols, p, sheet);
    const visualRow = p.physicalToVisual(7)!;
    expect(visualRow).toBe(6);
    const rect = cellRectInCanvas({ row: visualRow, col: 0 }, layout, rows, cols);
    // Visual row 6 starts at header + 6*24 (identity would give 7*24 for
    // physical row 7 — off by one hidden row's height).
    expect(rect.y).toBe(HEADER_H + 6 * 24);
    expect(rect.height).toBe(24);
  });
});

describe("dirty region conversion", () => {
  it("physical change ranges map to visual rects; fully-hidden ranges produce none", () => {
    const sheet = fakeSheet();
    const p = makeProjection();
    const rows = visualAxis(sheet, p);
    const cols = colAxis(sheet);
    const layout = layoutFor(rows, cols, p, sheet);
    const tracker = new DirtyRegionTracker();
    tracker.consume(layout, rows, cols, p); // clear initial full flag

    tracker.pushEvent({
      workbookId: "wb",
      sheetId: "s1",
      changes: [{ range: { startRow: 10, startCol: 0, endRow: 11, endCol: 2 }, kind: "cells" }],
      source: "user",
      batch: true,
    });
    expect(tracker.consume(layout, rows, cols, p).rects).toHaveLength(0);

    tracker.pushEvent({
      workbookId: "wb",
      sheetId: "s1",
      changes: [{ range: { startRow: 5, startCol: 0, endRow: 14, endCol: 2 }, kind: "cells" }],
      source: "user",
      batch: true,
    });
    const { rects } = tracker.consume(layout, rows, cols, p);
    expect(rects).toHaveLength(1);
    // Visual span 5..8 (physical 5,7,9,12): y = header + 5*24, height 4*24.
    expect(rects[0]).toEqual({ x: HEADER_W, y: HEADER_H + 5 * 24, width: 3 * 100, height: 4 * 24 });
  });

  it("filter/reorder events force a full redraw (structure path)", () => {
    const tracker = new DirtyRegionTracker();
    for (const kind of ["filter", "reorder"] as const) {
      tracker.pushEvent({
        workbookId: "wb",
        sheetId: "s1",
        changes: [{ range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, kind }],
        source: "user",
        batch: true,
      });
      expect(tracker.needsStructureRebuild).toBe(true);
      const sheet = fakeSheet();
      const p = makeProjection();
      const rows = visualAxis(sheet, p);
      const cols = colAxis(sheet);
      expect(tracker.consume(layoutFor(rows, cols, p, sheet), rows, cols, p).full).toBe(true);
    }
  });
});

describe("scroll-to-active with projection", () => {
  it("scrolls the visual axis so the active physical cell becomes visible", () => {
    const sheet = fakeSheet();
    const p = makeProjection();
    const rows = visualAxis(sheet, p);
    const cols = colAxis(sheet);
    // Active physical 15 → visual 9; bottom edge = header + 10*24 = 266 ≤ 574
    // → no scroll needed yet. Use physical 19 (visual 13) → needs scrolling.
    const activeVisual = p.physicalToVisual(19)!;
    expect(activeVisual).toBe(13);
    const next = computeScrollToCell({ row: activeVisual, col: 0 }, { scrollX: 0, scrollY: 0 }, {
      viewportWidth: 800 - HEADER_W,
      viewportHeight: 600 - HEADER_H,
      rows,
      cols,
      frozenRowCount: p.visibleCountBefore(sheet.frozenRows),
      frozenColCount: sheet.frozenColumns,
    });
    // Visual cell 13 spans 312..336; viewport 574 → still no scroll.
    expect(next.scrollY).toBe(0);
    // Force a scroll: far-down sheet.
    const bigSheet = (() => {
      const s = fakeSheet();
      return { ...s, rowCount: 1000 };
    })();
    const bigProjection = new FilteredRowProjection(1000, { startRow: 5, endRow: 14 }, [5, 7, 9, 12]);
    const bigRows = new AxisMetrics(bigProjection.visualRowCount, 24, (v) =>
      bigSheet.getRowHeight(bigProjection.visualToPhysical(v)),
    );
    const lastVisual = bigProjection.physicalToVisual(999)!;
    const scrolled = computeScrollToCell({ row: lastVisual, col: 0 }, { scrollX: 0, scrollY: 0 }, {
      viewportWidth: 800 - HEADER_W,
      viewportHeight: 600 - HEADER_H,
      rows: bigRows,
      cols,
      frozenRowCount: 0,
      frozenColCount: 0,
    });
    expect(scrolled.scrollY).toBe(bigRows.positionOf(lastVisual + 1) - (600 - HEADER_H));
  });
});

describe("M4.1.1 hardening", () => {
  it("visualToPhysical throws SheetError when no rows are visible", () => {
    const empty = new FilteredRowProjection(10, { startRow: 0, endRow: 9 }, []);
    expect(empty.visualRowCount).toBe(0);
    expect(() => empty.visualToPhysical(0)).toThrow(SheetError);
    expect(() => new IdentityRowProjection(0).visualToPhysical(0)).toThrow(SheetError);
  });

  it("zero-visible-row compositions degrade safely (no -1 coordinates)", () => {
    const empty = new FilteredRowProjection(10, { startRow: 0, endRow: 9 }, []);
    expect(lastVisiblePhysicalRow(empty)).toBe(-1);
    expect(relocateToVisibleRow(empty, 3)).toBeUndefined();
    expect(physicalRangeToVisualRange({ startRow: 0, startCol: 0, endRow: 9, endCol: 2 }, empty)).toBeNull();
    // Hit test with an empty axis reports "outside", never a cell.
    const sheet = fakeSheet();
    const rows = new AxisMetrics(empty.visualRowCount, 24, () => undefined);
    const cols = colAxis(sheet);
    const layout = computeViewport({
      scrollX: 0, scrollY: 0, width: 800, height: 600, rows, cols,
      frozenRowCount: 0, frozenColCount: 0, bufferPx: 0,
      headerWidth: HEADER_W, headerHeight: HEADER_H,
    });
    expect(layout.main.rowEnd).toBe(-1); // paint loops no-op
    const hit = hitTestCell({
      x: HEADER_W + 5, y: HEADER_H + 5, layout, rows, cols,
      scrollX: 0, scrollY: 0, headerWidth: HEADER_W, headerHeight: HEADER_H,
      scrollbarSize: 10, rowCount: rows.length, colCount: COLS,
    });
    expect(hit.zone).toBe("outside");
  });

  it("clampScroll pulls a deep scroll back after the axis shrinks", () => {
    const sheet = fakeSheet();
    const cols = colAxis(sheet);
    // Deep scroll on the full 1000-row axis.
    const bigRows = new AxisMetrics(1000, 24, () => undefined);
    const deep = { scrollX: 0, scrollY: bigRows.totalSize - 500 };
    // Filter hides 90% → axis shrinks to 100 rows.
    const shrunk = new AxisMetrics(100, 24, () => undefined);
    const clamped = clampScroll(deep, shrunk, cols, 800 - HEADER_W, 600 - HEADER_H);
    expect(clamped.scrollY).toBeLessThanOrEqual(Math.max(0, shrunk.totalSize - (600 - HEADER_H)));
    expect(clamped.scrollY).toBeGreaterThan(0); // still shows the last screen
  });

  it("cellRectInCanvas: non-frozen cell above the viewport does not crash (no freeze)", () => {
    const sheet = fakeSheet();
    const p = new IdentityRowProjection(PHYSICAL);
    const rows = visualAxis(sheet, p);
    const cols = colAxis(sheet);
    // Scrolled deep: main.rowStart = 10; cell row 2 is above the viewport and
    // NOT frozen — previously misclassified via `cell.row < main.rowStart`.
    const layout = computeViewport({
      scrollX: 0, scrollY: 10 * 24, width: 800, height: 600, rows, cols,
      frozenRowCount: 0, frozenColCount: 0, bufferPx: 0,
      headerWidth: HEADER_W, headerHeight: HEADER_H,
    });
    expect(layout.main.rowStart).toBe(10);
    expect(layout.top).toBeNull();
    const rect = cellRectInCanvas({ row: 2, col: 0 }, layout, rows, cols);
    // Main-quadrant math: 2 is 8 rows above the window → negative offset.
    expect(rect.y).toBe(layout.main.originY - 8 * 24);
    expect(rect.height).toBe(24);
  });

  it("cellRectInCanvas: frozen membership comes from quadrant bounds", () => {
    const sheet = fakeSheet(1); // physical row 0 frozen
    const p = new IdentityRowProjection(PHYSICAL);
    const rows = visualAxis(sheet, p);
    const cols = colAxis(sheet);
    const layout = computeViewport({
      scrollX: 0, scrollY: 10 * 24, width: 800, height: 600, rows, cols,
      frozenRowCount: p.visibleCountBefore(sheet.frozenRows), frozenColCount: 0, bufferPx: 0,
      headerWidth: HEADER_W, headerHeight: HEADER_H,
    });
    // Frozen row 0 → top strip at the header.
    expect(cellRectInCanvas({ row: 0, col: 0 }, layout, rows, cols).y).toBe(HEADER_H);
    // Row 2 is not inside top.rowStart..rowEnd (only row 0 is frozen) — it
    // resolves against the MAIN quadrant even though it sits above the
    // scrolled window (offscreen, but quadrant-correct and crash-free).
    const rect = cellRectInCanvas({ row: 2, col: 0 }, layout, rows, cols);
    expect(rect.y).toBe(layout.main.originY + rows.positionOf(2) - rows.positionOf(layout.main.rowStart));
    expect(rect.y).not.toBe(HEADER_H);
  });
});

describe("projection cleared → identity", () => {
  it("restores 1:1 mapping, physical selection stays aligned", () => {
    const sheet = fakeSheet();
    let projection: RowProjection = makeProjection();
    // With the filter: physical 7 sits at visual 6.
    expect(projection.physicalToVisual(7)).toBe(6);
    // Clear (grid.setRowProjection(null)).
    projection = new IdentityRowProjection(sheet.rowCount);
    const rows = visualAxis(sheet, projection);
    const cols = colAxis(sheet);
    expect(rows.length).toBe(PHYSICAL);
    expect(projection.physicalToVisual(7)).toBe(7);
    const layout = layoutFor(rows, cols, projection, sheet);
    expect(layout.frozenHeight).toBe(0);
    // Hit test maps 1:1 again.
    const hit = hitTestCell({
      x: HEADER_W + 5,
      y: HEADER_H + rows.positionOf(7) + 2,
      layout,
      rows,
      cols,
      scrollX: 0,
      scrollY: 0,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
      scrollbarSize: 10,
      rowCount: rows.length,
      colCount: COLS,
    });
    expect(projection.visualToPhysical(hit.row)).toBe(7);
  });
});
