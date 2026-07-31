// coordinate-mapper: pure hit-testing and scrollbar geometry. Shared by the
// grid input layer and (M2) the DOM editor positioning.

import type { AxisMetrics } from "./axis-metrics.js";
import type { ViewportLayout } from "./viewport.js";

export type HitZone =
  | "cell"
  | "colHeader"
  | "rowHeader"
  | "cornerHeader"
  | "vScrollbar"
  | "hScrollbar"
  | "outside";

export interface CellHit {
  zone: HitZone;
  row: number;
  col: number;
}

export interface HitTestInput {
  x: number;
  y: number;
  layout: ViewportLayout;
  rows: AxisMetrics;
  cols: AxisMetrics;
  scrollX: number;
  scrollY: number;
  headerWidth: number;
  headerHeight: number;
  scrollbarSize: number;
  rowCount: number;
  colCount: number;
}

/**
 * Map canvas point → cell/zone. Scroll semantics match computeViewport:
 * scroll is relative to the non-frozen area.
 */
export function hitTestCell(input: HitTestInput): CellHit {
  const { x, y, layout } = input;
  const width = layout.mainX + layout.mainWidth;
  const height = layout.mainY + layout.mainHeight;

  if (x >= width - input.scrollbarSize && y >= input.headerHeight && y <= height - input.scrollbarSize) {
    return { zone: "vScrollbar", row: -1, col: -1 };
  }
  if (y >= height - input.scrollbarSize && x >= input.headerWidth && x <= width - input.scrollbarSize) {
    return { zone: "hScrollbar", row: -1, col: -1 };
  }
  if (x < input.headerWidth && y < input.headerHeight) {
    return { zone: "cornerHeader", row: -1, col: -1 };
  }
  if (y < input.headerHeight) {
    return { zone: "colHeader", row: -1, col: colAt(x, input) };
  }
  if (x < input.headerWidth) {
    return { zone: "rowHeader", row: rowAt(y, input), col: -1 };
  }
  const row = rowAt(y, input);
  const col = colAt(x, input);
  if (row >= input.rowCount || col >= input.colCount) {
    return { zone: "outside", row, col };
  }
  return { zone: "cell", row, col };
}

function rowAt(y: number, input: HitTestInput): number {
  const { layout, rows } = input;
  const inFrozenRow = y < layout.mainY;
  const row = inFrozenRow
    ? rows.indexAt(y - input.headerHeight)
    : rows.indexAt(input.layout.frozenHeight + input.scrollY + (y - layout.mainY));
  return row;
}

function colAt(x: number, input: HitTestInput): number {
  const { layout, cols } = input;
  const inFrozenCol = x < layout.mainX;
  const col = inFrozenCol
    ? cols.indexAt(x - input.headerWidth)
    : cols.indexAt(input.layout.frozenWidth + input.scrollX + (x - layout.mainX));
  return col;
}

export interface ScrollbarThumb {
  x: number;
  y: number;
  width: number;
  height: number;
  trackStart: number;
  trackSize: number;
  maxScroll: number;
}

export interface ScrollbarGeometry {
  vertical: ScrollbarThumb | null;
  horizontal: ScrollbarThumb | null;
}

export interface ScrollbarInput {
  layout: ViewportLayout;
  rows: AxisMetrics;
  cols: AxisMetrics;
  scrollX: number;
  scrollY: number;
  width: number;
  height: number;
  headerWidth: number;
  headerHeight: number;
  scrollbarSize: number;
}

/** Scrollbar track/thumb geometry. Scroll extents exclude the frozen zone. */
export function computeScrollbarGeometry(input: ScrollbarInput): ScrollbarGeometry {
  const { layout } = input;
  const minThumb = 24;
  const scrollableH = Math.max(0, input.rows.totalSize - layout.frozenHeight);
  const scrollableW = Math.max(0, input.cols.totalSize - layout.frozenWidth);
  const mainH = layout.mainHeight;
  const mainW = layout.mainWidth;

  let vertical: ScrollbarThumb | null = null;
  if (mainH > 0 && scrollableH > mainH) {
    const trackStart = input.headerHeight;
    const trackSize = input.height - input.headerHeight - input.scrollbarSize;
    const maxScroll = scrollableH - mainH;
    const thumbHeight = Math.max(minThumb, (mainH / scrollableH) * trackSize);
    const thumbY = trackStart + (input.scrollY / maxScroll) * (trackSize - thumbHeight);
    vertical = { x: input.width - input.scrollbarSize, y: thumbY, width: input.scrollbarSize, height: thumbHeight, trackStart, trackSize, maxScroll };
  }

  let horizontal: ScrollbarThumb | null = null;
  if (mainW > 0 && scrollableW > mainW) {
    const trackStart = input.headerWidth;
    const trackSize = input.width - input.headerWidth - input.scrollbarSize;
    const maxScroll = scrollableW - mainW;
    const thumbWidth = Math.max(minThumb, (mainW / scrollableW) * trackSize);
    const thumbX = trackStart + (input.scrollX / maxScroll) * (trackSize - thumbWidth);
    horizontal = { x: thumbX, y: input.height - input.scrollbarSize, width: thumbWidth, height: input.scrollbarSize, trackStart, trackSize, maxScroll };
  }

  return { vertical, horizontal };
}
