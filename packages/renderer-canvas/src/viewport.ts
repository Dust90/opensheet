// Viewport: visible range computation + frozen-pane quadrant geometry.
// Pure functions — no DOM, unit-testable in node.
//
// SCROLL SEMANTICS (ADR: M1.9): scrollX/scrollY are RELATIVE TO THE
// NON-FROZEN AREA. At scroll=0 the main quadrant starts at the first
// non-frozen row/col (frozenRowCount/frozenColCount), so frozen content is
// never duplicated.

import type { CellAddress } from "@opensheet/shared";
import type { AxisMetrics } from "./axis-metrics.js";

export interface ViewportInput {
  scrollX: number;
  scrollY: number;
  /** Canvas CSS pixel size. */
  width: number;
  height: number;
  rows: AxisMetrics;
  cols: AxisMetrics;
  frozenRowCount: number;
  frozenColCount: number;
  /** Extra pixels of overscan beyond each edge (buffer zone). */
  bufferPx: number;
  headerWidth: number;
  headerHeight: number;
}

/** Inclusive index range plus its canvas pixel origin. */
export interface Quadrant {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
  /** Canvas x of colStart's left edge. */
  originX: number;
  /** Canvas y of rowStart's top edge. */
  originY: number;
  /** Canvas clip rect for this quadrant. */
  clipX: number;
  clipY: number;
  clipWidth: number;
  clipHeight: number;
}

export interface ViewportLayout {
  /** Frozen corner (frozen rows × frozen cols), null when no freeze. */
  corner: Quadrant | null;
  /** Frozen rows strip (scrolls horizontally), null when no frozen rows. */
  top: Quadrant | null;
  /** Frozen cols strip (scrolls vertically), null when no frozen cols. */
  left: Quadrant | null;
  /** Main scrollable quadrant. */
  main: Quadrant;
  frozenWidth: number;
  frozenHeight: number;
  /** Canvas rect of the scrollable (non-frozen) area. */
  mainX: number;
  mainY: number;
  mainWidth: number;
  mainHeight: number;
}

export function computeViewport(input: ViewportInput): ViewportLayout {
  const { rows, cols, bufferPx } = input;
  const frozenRowCount = Math.min(input.frozenRowCount, rows.length);
  const frozenColCount = Math.min(input.frozenColCount, cols.length);
  const frozenWidth = cols.positionOf(frozenColCount);
  const frozenHeight = rows.positionOf(frozenRowCount);

  const mainX = input.headerWidth + frozenWidth;
  const mainY = input.headerHeight + frozenHeight;
  const mainWidth = Math.max(0, input.width - mainX);
  const mainHeight = Math.max(0, input.height - mainY);

  // scroll is relative to the non-frozen area: absolute content position of
  // the main window start = frozenHeight + scrollY (rows) / frozenWidth + scrollX.
  const mainRowStart = Math.max(
    frozenRowCount,
    rows.indexAt(frozenHeight + Math.max(0, input.scrollY) - bufferPx),
  );
  const mainRowEnd = Math.min(
    rows.length - 1,
    rows.indexAt(frozenHeight + input.scrollY + mainHeight + bufferPx),
  );
  const mainColStart = Math.max(
    frozenColCount,
    cols.indexAt(frozenWidth + Math.max(0, input.scrollX) - bufferPx),
  );
  const mainColEnd = Math.min(
    cols.length - 1,
    cols.indexAt(frozenWidth + input.scrollX + mainWidth + bufferPx),
  );

  const main: Quadrant = {
    rowStart: mainRowStart,
    rowEnd: mainRowEnd,
    colStart: mainColStart,
    colEnd: mainColEnd,
    originX: mainX + cols.positionOf(mainColStart) - frozenWidth - input.scrollX,
    originY: mainY + rows.positionOf(mainRowStart) - frozenHeight - input.scrollY,
    clipX: mainX,
    clipY: mainY,
    clipWidth: mainWidth,
    clipHeight: mainHeight,
  };

  let corner: Quadrant | null = null;
  let top: Quadrant | null = null;
  let left: Quadrant | null = null;

  if (frozenRowCount > 0) {
    top = {
      rowStart: 0,
      rowEnd: frozenRowCount - 1,
      colStart: mainColStart,
      colEnd: mainColEnd,
      originX: main.originX,
      originY: input.headerHeight,
      clipX: mainX,
      clipY: input.headerHeight,
      clipWidth: mainWidth,
      clipHeight: frozenHeight,
    };
  }
  if (frozenColCount > 0) {
    left = {
      rowStart: mainRowStart,
      rowEnd: mainRowEnd,
      colStart: 0,
      colEnd: frozenColCount - 1,
      originX: input.headerWidth,
      originY: main.originY,
      clipX: input.headerWidth,
      clipY: mainY,
      clipWidth: frozenWidth,
      clipHeight: mainHeight,
    };
  }
  if (frozenRowCount > 0 && frozenColCount > 0) {
    corner = {
      rowStart: 0,
      rowEnd: frozenRowCount - 1,
      colStart: 0,
      colEnd: frozenColCount - 1,
      originX: input.headerWidth,
      originY: input.headerHeight,
      clipX: input.headerWidth,
      clipY: input.headerHeight,
      clipWidth: frozenWidth,
      clipHeight: frozenHeight,
    };
  }

  return { corner, top, left, main, frozenWidth, frozenHeight, mainX, mainY, mainWidth, mainHeight };
}

export interface ScrollPosition {
  scrollX: number;
  scrollY: number;
}

/**
 * Minimal scroll adjustment so `cell` becomes fully visible. Cells inside
 * frozen zones are always visible. scroll is relative to the non-frozen
 * area, so cell positions are shifted by the frozen sizes.
 */
export function computeScrollToCell(
  cell: CellAddress,
  scroll: ScrollPosition,
  input: {
    viewportWidth: number;
    viewportHeight: number;
    rows: AxisMetrics;
    cols: AxisMetrics;
    frozenRowCount: number;
    frozenColCount: number;
  },
): ScrollPosition {
  const { rows, cols } = input;
  const frozenWidth = cols.positionOf(Math.min(input.frozenColCount, cols.length));
  const frozenHeight = rows.positionOf(Math.min(input.frozenRowCount, rows.length));
  const viewWidth = input.viewportWidth - frozenWidth;
  const viewHeight = input.viewportHeight - frozenHeight;

  let { scrollX, scrollY } = scroll;

  if (cell.col >= input.frozenColCount) {
    const cellStart = cols.positionOf(cell.col) - frozenWidth;
    const cellEnd = cellStart + cols.sizeOf(cell.col);
    if (cellStart < scrollX) scrollX = cellStart;
    else if (cellEnd > scrollX + viewWidth) scrollX = cellEnd - viewWidth;
  }
  if (cell.row >= input.frozenRowCount) {
    const cellStart = rows.positionOf(cell.row) - frozenHeight;
    const cellEnd = cellStart + rows.sizeOf(cell.row);
    if (cellStart < scrollY) scrollY = cellStart;
    else if (cellEnd > scrollY + viewHeight) scrollY = cellEnd - viewHeight;
  }

  // viewWidth/viewHeight already exclude the frozen zone, so:
  //   maxScroll = (totalSize - frozen) - view = totalSize - viewportSize.
  const maxScrollX = Math.max(0, cols.totalSize - frozenWidth - viewWidth);
  const maxScrollY = Math.max(0, rows.totalSize - frozenHeight - viewHeight);
  return {
    scrollX: Math.min(Math.max(0, scrollX), maxScrollX),
    scrollY: Math.min(Math.max(0, scrollY), maxScrollY),
  };
}

/**
 * Clamp scroll into valid range.
 *
 * viewportWidth/Height are the header-exclusive canvas size (i.e. they still
 * INCLUDE the frozen zone). Scroll is relative to the non-frozen area, so:
 *
 *   scrollable content = totalSize - frozenSize
 *   main window        = viewportSize - frozenSize
 *   maxScroll          = scrollable content - main window
 *                      = totalSize - viewportSize        (frozen cancels out)
 *
 * This must stay consistent with computeScrollbarGeometry().maxScroll, which
 * computes (totalSize - frozen) - (viewport - frozen) = totalSize - viewport.
 */
export function clampScroll(
  scroll: ScrollPosition,
  rows: AxisMetrics,
  cols: AxisMetrics,
  viewportWidth: number,
  viewportHeight: number,
): ScrollPosition {
  const maxScrollX = Math.max(0, cols.totalSize - viewportWidth);
  const maxScrollY = Math.max(0, rows.totalSize - viewportHeight);
  return {
    scrollX: Math.min(Math.max(0, scroll.scrollX), maxScrollX),
    scrollY: Math.min(Math.max(0, scroll.scrollY), maxScrollY),
  };
}
