// M4.1 RowProjection unit tests (pure, no DOM).
//
// Scenario used throughout: 20 physical rows, filter range rows 5..14
// (10 rows), visible inside = [5(header), 7, 9, 12] → 6 hidden,
// visualRowCount = 14.

import { describe, expect, it } from "vitest";
import { SheetError } from "@injoysai/opensheet-shared";
import { AxisMetrics } from "../axis-metrics.js";
import {
  FilteredRowProjection,
  IdentityRowProjection,
  lastVisiblePhysicalRow,
  physicalRangeToVisualRange,
  relocateToVisibleRow,
  type RowProjection,
} from "../row-projection.js";

const PHYSICAL = 20;
const RANGE = { startRow: 5, endRow: 14 };
const VISIBLE = [5, 7, 9, 12];

function makeProjection(): FilteredRowProjection {
  return new FilteredRowProjection(PHYSICAL, RANGE, VISIBLE);
}

describe("IdentityRowProjection", () => {
  it("maps both directions 1:1", () => {
    const p = new IdentityRowProjection(PHYSICAL);
    expect(p.visualRowCount).toBe(PHYSICAL);
    for (let row = 0; row < PHYSICAL; row++) {
      expect(p.visualToPhysical(row)).toBe(row);
      expect(p.physicalToVisual(row)).toBe(row);
      expect(p.isVisible(row)).toBe(true);
    }
    expect(p.physicalToVisual(PHYSICAL)).toBeUndefined();
    expect(p.nextVisible(0, -1)).toBeUndefined();
    expect(p.nextVisible(PHYSICAL - 1, 1)).toBeUndefined();
    expect(p.visibleCountBefore(7)).toBe(7);
  });
});

describe("FilteredRowProjection mapping", () => {
  const p = makeProjection();

  it("counts: hidden rows shrink the visual axis", () => {
    expect(p.physicalRowCount).toBe(20);
    expect(p.visualRowCount).toBe(14);
  });

  it("rows before the filter range map by identity", () => {
    for (let row = 0; row < 5; row++) {
      expect(p.visualToPhysical(row)).toBe(row);
      expect(p.physicalToVisual(row)).toBe(row);
    }
  });

  it("visible rows inside the filter range map compactly", () => {
    expect(p.physicalToVisual(5)).toBe(5);
    expect(p.physicalToVisual(7)).toBe(6);
    expect(p.physicalToVisual(9)).toBe(7);
    expect(p.physicalToVisual(12)).toBe(8);
    expect(p.visualToPhysical(5)).toBe(5);
    expect(p.visualToPhysical(6)).toBe(7);
    expect(p.visualToPhysical(7)).toBe(9);
    expect(p.visualToPhysical(8)).toBe(12);
  });

  it("hidden rows inside the filter range return undefined", () => {
    for (const hidden of [6, 8, 10, 11, 13, 14]) {
      expect(p.physicalToVisual(hidden)).toBeUndefined();
      expect(p.isVisible(hidden)).toBe(false);
    }
  });

  it("rows after the filter range shift by hiddenCount", () => {
    expect(p.physicalToVisual(15)).toBe(9);
    expect(p.physicalToVisual(19)).toBe(13);
    expect(p.visualToPhysical(9)).toBe(15);
    expect(p.visualToPhysical(13)).toBe(19);
  });

  it("rejects invalid construction input with SheetError", () => {
    expect(() => new FilteredRowProjection(20, { startRow: 5, endRow: 4 }, [])).toThrow(SheetError);
    expect(() => new FilteredRowProjection(20, { startRow: 5, endRow: 20 }, [])).toThrow(SheetError);
    expect(() => new FilteredRowProjection(20, RANGE, [7, 5])).toThrow(SheetError); // unsorted
    expect(() => new FilteredRowProjection(20, RANGE, [5, 5])).toThrow(SheetError); // duplicate
    expect(() => new FilteredRowProjection(20, RANGE, [3])).toThrow(SheetError); // outside range
  });
});

describe("FilteredRowProjection.nextVisible", () => {
  const p = makeProjection();

  it("forward skips hidden rows", () => {
    expect(p.nextVisible(4, 1)).toBe(5);
    expect(p.nextVisible(5, 1)).toBe(7);
    expect(p.nextVisible(7, 1)).toBe(9);
    expect(p.nextVisible(9, 1)).toBe(12);
    expect(p.nextVisible(12, 1)).toBe(15); // 13,14 hidden → lands after range
    expect(p.nextVisible(19, 1)).toBeUndefined();
  });

  it("backward skips hidden rows", () => {
    expect(p.nextVisible(15, -1)).toBe(12);
    expect(p.nextVisible(12, -1)).toBe(9);
    expect(p.nextVisible(5, -1)).toBe(4);
    expect(p.nextVisible(0, -1)).toBeUndefined();
  });

  it("works from a hidden row (both directions)", () => {
    expect(p.nextVisible(6, 1)).toBe(7);
    expect(p.nextVisible(6, -1)).toBe(5);
    expect(p.nextVisible(14, 1)).toBe(15);
    expect(p.nextVisible(10, -1)).toBe(9);
  });

  it("all data rows hidden: falls through to rows outside the range", () => {
    const all = new FilteredRowProjection(20, { startRow: 2, endRow: 4 }, []);
    expect(all.visualRowCount).toBe(17);
    expect(all.nextVisible(1, 1)).toBe(5);
    expect(all.nextVisible(5, -1)).toBe(1);
    expect(all.physicalToVisual(3)).toBeUndefined();
    expect(all.visualToPhysical(2)).toBe(5);
  });
});

describe("visibleCountBefore (frozen boundary)", () => {
  const p = makeProjection();

  it("counts only visible rows below the bound", () => {
    expect(p.visibleCountBefore(0)).toBe(0);
    expect(p.visibleCountBefore(3)).toBe(3);
    expect(p.visibleCountBefore(5)).toBe(5);
    expect(p.visibleCountBefore(6)).toBe(6); // 0..5 all visible
    expect(p.visibleCountBefore(7)).toBe(6); // row 6 hidden
    expect(p.visibleCountBefore(8)).toBe(7); // rows 0..7: 8 minus hidden {6}
    expect(p.visibleCountBefore(13)).toBe(9); // rows 0..12: 13 minus hidden {6,8,10,11}
    expect(p.visibleCountBefore(15)).toBe(9); // rows 0..14: 15 minus 6 hidden
    expect(p.visibleCountBefore(20)).toBe(14);
  });

  it("frozen zone entirely before the filter range is unaffected", () => {
    expect(p.visibleCountBefore(4)).toBe(4);
  });

  it("header stays visible and counts toward the frozen zone", () => {
    // hasHeader filters include the header row in the visible list (done at
    // construction time by the caller): freezing the header keeps it pinned.
    expect(p.isVisible(5)).toBe(true);
    expect(p.visibleCountBefore(6)).toBe(6); // physical rows 0..5 → 6 visual rows
  });
});

describe("AxisMetrics over the visual axis", () => {
  it("custom PHYSICAL row heights land on the correct visual rows", () => {
    const p = makeProjection();
    const heights = new Map<number, number>([[7, 100], [15, 50]]); // physical rows
    const rows = new AxisMetrics(p.visualRowCount, 24, (visualRow) =>
      heights.get(p.visualToPhysical(visualRow)),
    );
    expect(rows.length).toBe(14);
    // Visual row 6 = physical 7 → custom height 100.
    expect(rows.sizeOf(6)).toBe(100);
    // Visual row 9 = physical 15 → custom height 50.
    expect(rows.sizeOf(9)).toBe(50);
    // Hidden physical rows occupy no pixels anywhere.
    expect(rows.totalSize).toBe(14 * 24 - 24 - 24 + 100 + 50);
    // A hidden physical row's (absent) height never leaks into the axis.
    expect(rows.sizeOf(5)).toBe(24); // physical 5, default
  });
});

describe("composition helpers", () => {
  const p = makeProjection();

  it("physicalRangeToVisualRange spans visible rows, skipping hidden edges", () => {
    // Fully visible span outside the range.
    expect(physicalRangeToVisualRange({ startRow: 0, startCol: 0, endRow: 4, endCol: 2 }, p))
      .toEqual({ startRow: 0, startCol: 0, endRow: 4, endCol: 2 });
    // Edges hidden: physical 5..14 → visual 5..8 (physical 5 and 12 visible).
    expect(physicalRangeToVisualRange({ startRow: 5, startCol: 0, endRow: 14, endCol: 2 }, p))
      .toEqual({ startRow: 5, startCol: 0, endRow: 8, endCol: 2 });
    // Hidden start expands down to the first visible row.
    expect(physicalRangeToVisualRange({ startRow: 6, startCol: 0, endRow: 8, endCol: 2 }, p))
      .toEqual({ startRow: 6, startCol: 0, endRow: 6, endCol: 2 }); // only physical 7
    // No visible rows at all → null.
    expect(physicalRangeToVisualRange({ startRow: 10, startCol: 0, endRow: 11, endCol: 2 }, p)).toBeNull();
    // Cross-range span: physical 3..16 → visual 3..10.
    expect(physicalRangeToVisualRange({ startRow: 3, startCol: 1, endRow: 16, endCol: 1 }, p))
      .toEqual({ startRow: 3, startCol: 1, endRow: 10, endCol: 1 });
  });

  it("lastVisiblePhysicalRow is the Ctrl+End target", () => {
    expect(lastVisiblePhysicalRow(p)).toBe(19);
    const allHiddenTail = new FilteredRowProjection(20, { startRow: 15, endRow: 19 }, []);
    expect(lastVisiblePhysicalRow(allHiddenTail)).toBe(14);
  });

  it("relocateToVisibleRow implements the hidden-active-cell policy", () => {
    expect(relocateToVisibleRow(p, 3)).toBe(3); // already visible
    expect(relocateToVisibleRow(p, 6)).toBe(7); // next below
    expect(relocateToVisibleRow(p, 14)).toBe(15); // next below (13,14 hidden)
    const tail = new FilteredRowProjection(20, { startRow: 15, endRow: 19 }, []);
    expect(relocateToVisibleRow(tail, 17)).toBe(14); // none below → above
  });

  it("clearing the projection restores the identity mapping", () => {
    // Grid teardown path: setRowProjection(null) → IdentityRowProjection.
    let projection: RowProjection = makeProjection();
    expect(projection.physicalToVisual(7)).toBe(6);
    projection = new IdentityRowProjection(PHYSICAL);
    expect(projection.visualRowCount).toBe(PHYSICAL);
    expect(projection.physicalToVisual(7)).toBe(7);
    expect(projection.visualToPhysical(13)).toBe(13);
  });
});
