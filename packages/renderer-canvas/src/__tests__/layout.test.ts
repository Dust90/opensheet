import { describe, expect, it } from "vitest";
import { AxisMetrics } from "../axis-metrics.js";
import { computeScrollbarGeometry, hitTestCell } from "../coordinate-mapper.js";
import { DirtyRegionTracker, rangeToCanvasRects } from "../dirty-region.js";
import { IdentityRowProjection } from "../row-projection.js";
import { SelectionModel } from "../selection.js";
import { clampScroll, computeScrollToCell, computeViewport, type ViewportInput } from "../viewport.js";

const ROWS = 1000;
const COLS = 26;
const IDENTITY = new IdentityRowProjection(ROWS);

function makeAxes(rowOverrides = new Map<number, number>(), colOverrides = new Map<number, number>()) {
  const rows = new AxisMetrics(ROWS, 24, (i) => rowOverrides.get(i));
  const cols = new AxisMetrics(COLS, 100, (i) => colOverrides.get(i));
  return { rows, cols };
}

function viewportInput(rows: AxisMetrics, cols: AxisMetrics, over: Partial<ViewportInput> = {}): ViewportInput {
  return {
    scrollX: 0,
    scrollY: 0,
    width: 800,
    height: 600,
    rows,
    cols,
    frozenRowCount: 0,
    frozenColCount: 0,
    bufferPx: 0,
    headerWidth: 48,
    headerHeight: 26,
    ...over,
  };
}

describe("AxisMetrics", () => {
  it("uniform sizes: positionOf/indexAt are exact", () => {
    const { rows } = makeAxes();
    expect(rows.positionOf(0)).toBe(0);
    expect(rows.positionOf(10)).toBe(240);
    expect(rows.totalSize).toBe(24000);
    expect(rows.indexAt(0)).toBe(0);
    expect(rows.indexAt(239)).toBe(9);
    expect(rows.indexAt(240)).toBe(10);
    expect(rows.indexAt(241)).toBe(10);
  });

  it("variable sizes: binary search respects overrides", () => {
    const { rows } = makeAxes(new Map([[0, 100], [1, 10]]));
    // prefix: 0, 100, 110, 134, ...
    expect(rows.positionOf(1)).toBe(100);
    expect(rows.positionOf(2)).toBe(110);
    expect(rows.positionOf(3)).toBe(134);
    expect(rows.indexAt(99)).toBe(0);
    expect(rows.indexAt(100)).toBe(1);
    expect(rows.indexAt(109)).toBe(1);
    expect(rows.indexAt(110)).toBe(2);
    expect(rows.sizeOf(0)).toBe(100);
    expect(rows.sizeOf(1)).toBe(10);
    expect(rows.sizeOf(2)).toBe(24);
  });

  it("clamps out-of-range positions", () => {
    const { rows } = makeAxes();
    expect(rows.indexAt(-50)).toBe(0);
    expect(rows.indexAt(999999)).toBe(ROWS - 1);
  });
});

describe("computeViewport", () => {
  it("no freeze: main covers the visible window with correct origins", () => {
    const { rows, cols } = makeAxes();
    const layout = computeViewport(viewportInput(rows, cols, { scrollX: 250, scrollY: 120 }));
    // scrollY=120 / 24px rows → row 5; scrollX=250 / 100px → col 2
    expect(layout.main.rowStart).toBe(5);
    expect(layout.main.colStart).toBe(2);
    expect(layout.main.originY).toBe(26 + (120 - 120)); // header + (pos(5)-scroll) = 26 + 0
    expect(layout.main.originX).toBe(48 + (200 - 250)); // 48 + (-50) = -2
    expect(layout.corner).toBeNull();
    expect(layout.top).toBeNull();
    expect(layout.left).toBeNull();
  });

  it("frozen panes: four quadrants with independent scroll handling", () => {
    const { rows, cols } = makeAxes();
    const layout = computeViewport(
      viewportInput(rows, cols, {
        scrollX: 300,
        scrollY: 240,
        frozenRowCount: 2,
        frozenColCount: 1,
      }),
    );
    // frozen: 2 rows (48px), 1 col (100px)
    expect(layout.frozenHeight).toBe(48);
    expect(layout.frozenWidth).toBe(100);
    expect(layout.corner).not.toBeNull();
    expect(layout.corner!.rowEnd).toBe(1);
    expect(layout.corner!.colEnd).toBe(0);
    // top strip: frozen rows, scrolled cols
    expect(layout.top!.rowStart).toBe(0);
    // scroll is relative to the non-frozen area: content pos = frozen(100) + 300
    expect(layout.top!.colStart).toBe(4); // indexAt(400) = col 4
    expect(layout.top!.originY).toBe(26); // directly under header
    // left strip: scrolled rows, frozen cols
    expect(layout.left!.colStart).toBe(0);
    expect(layout.left!.rowStart).toBe(12); // indexAt(48 + 240) = row 12
    expect(layout.left!.originX).toBe(48);
    // main starts after frozen zones
    expect(layout.main.rowStart).toBe(12);
    expect(layout.main.colStart).toBe(4);
    expect(layout.main.originY).toBe(26 + 48 + (rows.positionOf(12) - 48 - 240)); // 0
    expect(layout.main.originX).toBe(48 + 100 + (cols.positionOf(4) - 100 - 300)); // 0
  });

  it("scroll=0 + freeze 1x1: main starts at B2, frozen content not duplicated", () => {
    const { rows, cols } = makeAxes();
    const layout = computeViewport(
      viewportInput(rows, cols, { frozenRowCount: 1, frozenColCount: 1 }),
    );
    expect(layout.main.rowStart).toBe(1);
    expect(layout.main.colStart).toBe(1);
    expect(layout.main.originX).toBe(48 + 100);
    expect(layout.main.originY).toBe(26 + 24);
    expect(layout.corner).not.toBeNull();
    expect(layout.top!.colStart).toBe(1); // top strip scrolls cols from B
    expect(layout.left!.rowStart).toBe(1); // left strip scrolls rows from row 2
  });

  it("scroll=0 + freeze: hit test at first main cell returns B2, not A1", () => {
    const { rows, cols } = makeAxes();
    const layout = computeViewport(viewportInput(rows, cols, { frozenRowCount: 1, frozenColCount: 1 }));
    const hit = hitTestCell({
      x: 48 + 100 + 10,
      y: 26 + 24 + 10,
      layout,
      rows,
      cols,
      scrollX: 0,
      scrollY: 0,
      headerWidth: 48,
      headerHeight: 26,
      scrollbarSize: 10,
      rowCount: ROWS,
      colCount: COLS,
    });
    expect(hit.zone).toBe("cell");
    expect(hit.row).toBe(1);
    expect(hit.col).toBe(1);
  });
});

describe("computeScrollToCell / clampScroll", () => {
  it("scrolls minimally to reveal a cell below the viewport", () => {
    const { rows, cols } = makeAxes();
    const next = computeScrollToCell({ row: 40, col: 0 }, { scrollX: 0, scrollY: 0 }, {
      viewportWidth: 752,
      viewportHeight: 574,
      rows,
      cols,
      frozenRowCount: 0,
      frozenColCount: 0,
    });
    // cell 40 spans 960..984; viewport shows 0..574 → scrollY = 984-574 = 410
    expect(next.scrollY).toBe(984 - 574);
  });

  it("cells inside frozen rows never trigger scroll", () => {
    const { rows, cols } = makeAxes();
    const next = computeScrollToCell({ row: 0, col: 0 }, { scrollX: 100, scrollY: 200 }, {
      viewportWidth: 752,
      viewportHeight: 574,
      rows,
      cols,
      frozenRowCount: 1,
      frozenColCount: 1,
    });
    expect(next).toEqual({ scrollX: 100, scrollY: 200 });
  });

  it("clampScroll limits to content extent", () => {
    const { rows, cols } = makeAxes();
    const clamped = clampScroll({ scrollX: 999999, scrollY: -5 }, rows, cols, 752, 574);
    expect(clamped.scrollY).toBe(0);
    expect(clamped.scrollX).toBe(cols.totalSize - 752);
  });

  it("clampScroll max matches scrollbar geometry max under freeze", () => {
    const { rows, cols } = makeAxes();
    const layout = computeViewport(
      viewportInput(rows, cols, { frozenRowCount: 1, frozenColCount: 1 }),
    );
    const geo = computeScrollbarGeometry({
      layout,
      rows,
      cols,
      scrollX: 0,
      scrollY: 0,
      width: 800,
      height: 600,
      headerWidth: 48,
      headerHeight: 26,
      scrollbarSize: 10,
    });
    // Same viewport semantics as SheetGrid.setScroll: header-exclusive size.
    const clamped = clampScroll(
      { scrollX: Number.MAX_SAFE_INTEGER, scrollY: Number.MAX_SAFE_INTEGER },
      rows,
      cols,
      800 - 48,
      600 - 26,
    );
    expect(clamped.scrollX).toBe(geo.horizontal!.maxScroll);
    expect(clamped.scrollY).toBe(geo.vertical!.maxScroll);
  });

  it("computeScrollToCell + clampScroll makes the last cell fully visible", () => {
    const { rows, cols } = makeAxes();
    const frozen = { frozenRowCount: 1, frozenColCount: 1 };
    const viewportWidth = 800 - 48;
    const viewportHeight = 600 - 26;
    const last = { row: ROWS - 1, col: COLS - 1 };

    const target = computeScrollToCell(last, { scrollX: 0, scrollY: 0 }, {
      viewportWidth,
      viewportHeight,
      rows,
      cols,
      ...frozen,
    });
    const clamped = clampScroll(target, rows, cols, viewportWidth, viewportHeight);
    const layout = computeViewport(
      viewportInput(rows, cols, { scrollX: clamped.scrollX, scrollY: clamped.scrollY, ...frozen }),
    );

    // Last cell's right/bottom edges must land inside the main window.
    const cellRight = cols.positionOf(last.col + 1) - layout.frozenWidth - clamped.scrollX;
    const cellBottom = rows.positionOf(last.row + 1) - layout.frozenHeight - clamped.scrollY;
    expect(cellRight).toBeLessThanOrEqual(layout.mainWidth);
    expect(cellBottom).toBeLessThanOrEqual(layout.mainHeight);
    // And the main quadrant must actually include the last cell.
    expect(layout.main.rowEnd).toBe(ROWS - 1);
    expect(layout.main.colEnd).toBe(COLS - 1);
  });

  it("scrollbar geometry excludes the frozen zone from extents", () => {
    const { rows, cols } = makeAxes();
    const layout = computeViewport(
      viewportInput(rows, cols, { frozenRowCount: 1, frozenColCount: 1 }),
    );
    const geo = computeScrollbarGeometry({
      layout,
      rows,
      cols,
      scrollX: 0,
      scrollY: 0,
      width: 800,
      height: 600,
      headerWidth: 48,
      headerHeight: 26,
      scrollbarSize: 10,
    });
    expect(geo.vertical).not.toBeNull();
    expect(geo.horizontal).not.toBeNull();
    // maxScroll = scrollable - main window (frozen zone not double-counted)
    expect(geo.vertical!.maxScroll).toBe(rows.totalSize - 24 - (600 - 26 - 24));
    expect(geo.horizontal!.maxScroll).toBe(cols.totalSize - 100 - (800 - 48 - 100));
  });
});

describe("SelectionModel", () => {
  it("setActive collapses; extendTo keeps anchor", () => {
    const sel = new SelectionModel(() => ROWS, () => COLS);
    sel.setActive({ row: 2, col: 3 });
    expect(sel.state.range).toEqual({ startRow: 2, startCol: 3, endRow: 2, endCol: 3 });
    sel.extendTo({ row: 5, col: 1 });
    expect(sel.state.range).toEqual({ startRow: 2, startCol: 1, endRow: 5, endCol: 3 });
    expect(sel.state.active).toEqual({ row: 5, col: 1 });
  });

  it("moveBy without shift moves from range edge and collapses", () => {
    const sel = new SelectionModel(() => ROWS, () => COLS);
    sel.setActive({ row: 2, col: 2 });
    sel.extendTo({ row: 5, col: 4 });
    sel.moveBy(1, 0, false);
    expect(sel.state.active).toEqual({ row: 6, col: 4 });
    sel.moveBy(-1, 0, false);
    expect(sel.state.active).toEqual({ row: 5, col: 4 });
  });

  it("moveBy with shift extends the range", () => {
    const sel = new SelectionModel(() => ROWS, () => COLS);
    sel.setActive({ row: 2, col: 2 });
    sel.moveBy(2, 1, true);
    expect(sel.state.range).toEqual({ startRow: 2, startCol: 2, endRow: 4, endCol: 3 });
  });

  it("clamps to sheet bounds", () => {
    const sel = new SelectionModel(() => 10, () => 5);
    sel.moveBy(-100, -100, false);
    expect(sel.state.active).toEqual({ row: 0, col: 0 });
    sel.moveBy(100, 100, false);
    expect(sel.state.active).toEqual({ row: 9, col: 4 });
  });
});

describe("DirtyRegionTracker", () => {
  it("cell events produce pixel rects; structure events force full redraw", () => {
    const { rows, cols } = makeAxes();
    const layout = computeViewport(viewportInput(rows, cols));
    const tracker = new DirtyRegionTracker();
    tracker.consume(layout, rows, cols, IDENTITY); // clear initial full flag

    tracker.pushEvent({
      workbookId: "wb",
      sheetId: "s1",
      changes: [{ range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, kind: "cells" }],
      source: "user",
      batch: true,
    });
    const result = tracker.consume(layout, rows, cols, IDENTITY);
    expect(result.full).toBe(false);
    expect(result.rects).toHaveLength(1);
    expect(result.rects[0]).toEqual({ x: 48, y: 26, width: 100, height: 24 });

    tracker.pushEvent({
      workbookId: "wb",
      sheetId: "s1",
      changes: [{ range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, kind: "structure" }],
      source: "api",
      batch: true,
    });
    expect(tracker.needsStructureRebuild).toBe(true);
    expect(tracker.consume(layout, rows, cols, IDENTITY).full).toBe(true);
  });

  it("rangeToCanvasRects splits ranges across frozen quadrants", () => {
    const { rows, cols } = makeAxes();
    const layout = computeViewport(
      viewportInput(rows, cols, { frozenRowCount: 1, frozenColCount: 1 }),
    );
    // Range spanning frozen + main in both axes: A1:B2 (row 0 frozen, col 0 frozen)
    const rects = rangeToCanvasRects(
      { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      layout,
      rows,
      cols,
    );
    // Expect corner (A1), top (B1), left (A2), main (B2) = 4 rects
    expect(rects).toHaveLength(4);
  });
});
