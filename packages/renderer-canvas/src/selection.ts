// SelectionModel: single-range selection with active cell. Pure, no DOM.

import type { CellAddress, Range } from "@opensheet/shared";
import { normalizeRange } from "@opensheet/shared";

export interface SelectionState {
  /** The cell that receives keyboard input / editing. */
  active: CellAddress;
  /** Current selection (normalized, may be a single cell). */
  range: Range;
}

/**
 * anchor + focus model: a drag or shift-move keeps `anchor` fixed and moves
 * `focus`; range = normalize(anchor..focus). Plain moves reset both.
 */
export class SelectionModel {
  private anchor: CellAddress = { row: 0, col: 0 };
  private focus: CellAddress = { row: 0, col: 0 };
  private readonly rowCount: () => number;
  private readonly colCount: () => number;

  constructor(rowCount: () => number, colCount: () => number) {
    this.rowCount = rowCount;
    this.colCount = colCount;
  }

  get state(): SelectionState {
    return {
      active: { ...this.focus },
      range: normalizeRange({
        startRow: this.anchor.row,
        startCol: this.anchor.col,
        endRow: this.focus.row,
        endCol: this.focus.col,
      }),
    };
  }

  private clampRow(row: number): number {
    return Math.min(Math.max(0, row), this.rowCount() - 1);
  }

  private clampCol(col: number): number {
    return Math.min(Math.max(0, col), this.colCount() - 1);
  }

  private clamp(addr: CellAddress): CellAddress {
    return { row: this.clampRow(addr.row), col: this.clampCol(addr.col) };
  }

  /** Plain selection: collapse to a single cell. */
  setActive(addr: CellAddress): void {
    const clamped = this.clamp(addr);
    this.anchor = clamped;
    this.focus = clamped;
  }

  /** Extend the range: anchor stays, focus moves. */
  extendTo(addr: CellAddress): void {
    this.focus = this.clamp(addr);
  }

  /** Keyboard move. extend=true keeps the anchor (shift+arrows). */
  moveBy(deltaRow: number, deltaCol: number, extend: boolean): void {
    if (!extend) {
      // Move from the range edge in the direction of travel, like desktop sheets.
      const range = this.state.range;
      const base: CellAddress = {
        row: deltaRow > 0 ? range.endRow : deltaRow < 0 ? range.startRow : this.focus.row,
        col: deltaCol > 0 ? range.endCol : deltaCol < 0 ? range.startCol : this.focus.col,
      };
      this.setActive({ row: base.row + deltaRow, col: base.col + deltaCol });
      return;
    }
    this.extendTo({ row: this.focus.row + deltaRow, col: this.focus.col + deltaCol });
  }

  /** Jump to first/last cell on an axis (Home/End, Ctrl+Home/End). */
  jumpTo(addr: CellAddress, extend: boolean): void {
    if (extend) this.extendTo(addr);
    else this.setActive(addr);
  }

  lastCell(): CellAddress {
    return { row: this.rowCount() - 1, col: this.colCount() - 1 };
  }

  /**
   * Force both anchor and focus back into bounds. Called by the renderer
   * after row/col structure changes so a stale out-of-range selection is
   * never painted.
   */
  clampSelection(): void {
    const anchor = this.clamp(this.anchor);
    const focus = this.clamp(this.focus);
    this.anchor = anchor;
    this.focus = focus;
  }
}
