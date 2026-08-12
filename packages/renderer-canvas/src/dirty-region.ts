// DirtyRegionTracker: ChangeEvent ranges → canvas pixel rects. Pure logic.

import type { ChangeEvent, Range } from "@injoysai/opensheet-shared";
import type { AxisMetrics } from "./axis-metrics.js";
import { physicalRangeToVisualRange, type RowProjection } from "./row-projection.js";
import type { ViewportLayout } from "./viewport.js";

export interface PixelRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Accumulates change events between frames and answers: "full content
 * redraw" or "only these pixel rects". Scroll changes force a full redraw
 * (viewport content is entirely new); cell/style changes produce rects for
 * the quadrants they intersect.
 */
export class DirtyRegionTracker {
  private ranges: Range[] = [];
  private full = true; // first frame paints everything
  private structureChanged = false;

  markFullRedraw(): void {
    this.full = true;
  }

  pushEvent(event: ChangeEvent): void {
    for (const change of event.changes) {
      if (
        change.kind === "rows" ||
        change.kind === "columns" ||
        change.kind === "structure" ||
        change.kind === "metadata" ||
        // M4: a filter re-projects the whole visual axis; a reorder moved
        // every value in the range — both need metrics + full repaint.
        change.kind === "filter" ||
        change.kind === "reorder"
      ) {
        this.structureChanged = true;
      } else {
        this.ranges.push(change.range);
      }
    }
  }

  /** True when metrics (row/col sizes, freeze) must be recomputed. */
  get needsStructureRebuild(): boolean {
    return this.structureChanged;
  }

  /**
   * Consume the accumulated state. Caller re-runs viewport computation first
   * when `needsStructureRebuild` was true. ChangeEvent ranges are PHYSICAL;
   * they are mapped onto the visual axis through `projection` before being
   * converted to pixel rects (fully-hidden ranges produce no rect).
   */
  consume(layout: ViewportLayout, rows: AxisMetrics, cols: AxisMetrics, projection: RowProjection): {
    full: boolean;
    rects: PixelRect[];
  } {
    const full = this.full || this.structureChanged;
    this.full = false;
    this.structureChanged = false;
    if (full) {
      this.ranges = [];
      return { full: true, rects: [] };
    }
    const rects: PixelRect[] = [];
    for (const range of this.ranges) {
      const visualRange = physicalRangeToVisualRange(range, projection);
      if (visualRange === null) continue;
      rects.push(...rangeToCanvasRects(visualRange, layout, rows, cols));
    }
    this.ranges = [];
    return { full: false, rects: mergeRects(rects) };
  }
}

/** Map a VISUAL-axis range to canvas rects (one per intersected quadrant). */
export function rangeToCanvasRects(
  range: Range,
  layout: ViewportLayout,
  rows: AxisMetrics,
  cols: AxisMetrics,
): PixelRect[] {
  const rects: PixelRect[] = [];
  const quadrants = [layout.main, layout.top, layout.left, layout.corner];
  for (const q of quadrants) {
    if (q === null) continue;
    const rowStart = Math.max(range.startRow, q.rowStart);
    const rowEnd = Math.min(range.endRow, q.rowEnd);
    const colStart = Math.max(range.startCol, q.colStart);
    const colEnd = Math.min(range.endCol, q.colEnd);
    if (rowStart > rowEnd || colStart > colEnd) continue;
    const x = q.originX + (cols.positionOf(colStart) - cols.positionOf(q.colStart));
    const y = q.originY + (rows.positionOf(rowStart) - rows.positionOf(q.rowStart));
    const width = cols.positionOf(colEnd + 1) - cols.positionOf(colStart);
    const height = rows.positionOf(rowEnd + 1) - rows.positionOf(rowStart);
    rects.push(intersectRect({ x, y, width, height }, {
      x: q.clipX,
      y: q.clipY,
      width: q.clipWidth,
      height: q.clipHeight,
    }));
  }
  return rects.filter((r) => r.width > 0 && r.height > 0);
}

function intersectRect(a: PixelRect, b: PixelRect): PixelRect {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  return {
    x,
    y,
    width: Math.min(a.x + a.width, b.x + b.width) - x,
    height: Math.min(a.y + a.height, b.y + b.height) - y,
  };
}

/** Merge overlapping/adjacent rects to reduce repaint passes. */
export function mergeRects(rects: PixelRect[]): PixelRect[] {
  const merged: PixelRect[] = [];
  for (const rect of rects) {
    let target = rect;
    let i = 0;
    while (i < merged.length) {
      const other = merged[i]!;
      if (rectsOverlapOrTouch(target, other)) {
        target = unionRect(target, other);
        merged.splice(i, 1);
      } else {
        i++;
      }
    }
    merged.push(target);
  }
  return merged;
}

function rectsOverlapOrTouch(a: PixelRect, b: PixelRect): boolean {
  return (
    a.x <= b.x + b.width &&
    b.x <= a.x + a.width &&
    a.y <= b.y + b.height &&
    b.y <= a.y + a.height
  );
}

function unionRect(a: PixelRect, b: PixelRect): PixelRect {
  const x = Math.min(a.x, b.x);
  const y = Math.min(a.y, b.y);
  return {
    x,
    y,
    width: Math.max(a.x + a.width, b.x + b.width) - x,
    height: Math.max(a.y + a.height, b.y + b.height) - y,
  };
}
