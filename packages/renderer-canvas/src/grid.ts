// SheetGrid: dual-canvas grid renderer.
//
// Boundary (ADR-0001, M1 guardrail): the renderer receives ONLY a
// WorksheetView — it cannot mutate workbook data. All it emits back is
// selection/scroll state via callbacks.

import type { WorksheetView } from "@injoysai/opensheet-core";
import type {
  CellPrimitive,
  CellStyle,
  ChangeEvent,
  ChangeListener,
  Unsubscribe,
} from "@injoysai/opensheet-shared";
import { isCellError, SheetError } from "@injoysai/opensheet-shared";
import { AxisMetrics } from "./axis-metrics.js";
import {
  cellRectInCanvas,
  computeScrollbarGeometry,
  hitTestCell,
  type ScrollbarGeometry,
} from "./coordinate-mapper.js";
import { DirtyRegionTracker, rangeToCanvasRects, type PixelRect } from "./dirty-region.js";
import { CellEditor } from "./editor/cell-editor.js";
import {
  cellDisplayText,
  decideKeyInPhase,
  isPrintableKey,
} from "./editor/editor-state.js";
import {
  IdentityRowProjection,
  lastVisiblePhysicalRow,
  physicalRangeToVisualRange,
  relocateToVisibleRow,
  type RowProjection,
} from "./row-projection.js";
import { SelectionModel } from "./selection.js";
import { lightTheme, type GridTheme } from "./theme.js";
import {
  clampScroll,
  computeScrollToCell,
  computeViewport,
  type Quadrant,
  type ViewportLayout,
} from "./viewport.js";

export interface SheetGridOptions {
  container: HTMLElement;
  /** READONLY data source. The renderer never sees Worksheet/Workbook/CommandBus. */
  worksheet: WorksheetView;
  /** Merged change events from the host (runtime.onChange). */
  onChange: (listener: ChangeListener) => Unsubscribe;
  /** Optional style lookup for cell presentation (host-owned style table). */
  resolveStyle?: (styleId: string) => Readonly<CellStyle> | undefined;
  theme?: GridTheme;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
  headerWidth?: number;
  headerHeight?: number;
  onSelectionChange?: (state: { activeRow: number; activeCol: number }) => void;
  /** Perf probe: called after each rendered frame with paint timings. */
  onFrame?: (stats: { paintMs: number; full: boolean; paintedCells: number }) => void;
  /** Inline editor committed text for `cell`. Host routes it to the Command Bus. */
  onCommitCell?: (init: { row: number; col: number; text: string }) => void;
  /** Selection copied (Cmd/Ctrl+C): host serializes to TSV and writes the clipboard. */
  onCopyCells?: (cells: CellPrimitive[][]) => void;
  /** Paste requested (Cmd/Ctrl+V): host reads the clipboard, parses TSV, writes ONE transaction anchored at `active`. */
  onPasteRequest?: (active: { row: number; col: number }) => void;
}

const DEFAULT_ROW_HEIGHT = 26;
const DEFAULT_COL_WIDTH = 100;
const DEFAULT_HEADER_W = 48;
const DEFAULT_HEADER_H = 26;
const BUFFER_PX = 240;
const SCROLLBAR = 10;
const CELL_PAD_X = 5;

type DragMode =
  | { kind: "select" }
  | { kind: "v-scroll"; grabOffset: number }
  | { kind: "h-scroll"; grabOffset: number }
  | null;

export class SheetGrid {
  private readonly container: HTMLElement;
  private readonly worksheet: WorksheetView;
  private readonly resolveStyleFn: ((styleId: string) => Readonly<CellStyle> | undefined) | undefined;
  private readonly theme: GridTheme;
  private readonly headerWidth: number;
  private readonly headerHeight: number;
  private readonly onSelectionChange: ((state: { activeRow: number; activeCol: number }) => void) | undefined;
  private readonly onFrame: ((stats: { paintMs: number; full: boolean; paintedCells: number }) => void) | undefined;
  private readonly onCommitCell: ((init: { row: number; col: number; text: string }) => void) | undefined;
  private readonly onCopyCells: ((cells: CellPrimitive[][]) => void) | undefined;
  private readonly onPasteRequest: ((active: { row: number; col: number }) => void) | undefined;
  private paintedCells = 0;
  private readonly defaultRowHeight: number;
  private readonly defaultColWidth: number;
  private headerDirty = true;
  private readonly containerPosition: string;
  private readonly containerOverflow: string;
  private readonly containerTabIndex: number;
  private readonly containerOutline: string;
  private readonly handleWheelBound: (e: WheelEvent) => void;
  private readonly handleMouseDownBound: (e: MouseEvent) => void;
  private readonly handleDblClickBound: (e: MouseEvent) => void;
  private readonly handleMouseMoveBound: (e: MouseEvent) => void;
  private readonly handleMouseUpBound: () => void;
  private readonly handleKeyDownBound: (e: KeyboardEvent) => void;

  private readonly contentCanvas: HTMLCanvasElement;
  private readonly overlayCanvas: HTMLCanvasElement;
  private readonly contentCtx: CanvasRenderingContext2D;
  private readonly overlayCtx: CanvasRenderingContext2D;

  private rows: AxisMetrics;
  private cols: AxisMetrics;
  /**
   * Visual ↔ physical row mapping (M4.1). The row AxisMetrics, viewport
   * quadrants and canvas Y are ALWAYS visual; SelectionModel, Worksheet data
   * and ChangeEvent ranges are ALWAYS physical. Identity when no filter is
   * active.
   */
  private projection: RowProjection;
  private scrollX = 0;
  private scrollY = 0;
  private layout: ViewportLayout | null = null;

  private readonly selection: SelectionModel;
  private readonly dirty = new DirtyRegionTracker();
  private readonly editor: CellEditor;
  private editingCell: { row: number; col: number } | null = null;
  private rafId = 0;
  private destroyed = false;
  private drag: DragMode = null;
  private readonly unsubscribe: Unsubscribe;
  private readonly resizeObserver: ResizeObserver;

  constructor(options: SheetGridOptions) {
    this.container = options.container;
    this.worksheet = options.worksheet;
    this.resolveStyleFn = options.resolveStyle;
    this.theme = options.theme ?? lightTheme;
    this.headerWidth = options.headerWidth ?? DEFAULT_HEADER_W;
    this.headerHeight = options.headerHeight ?? DEFAULT_HEADER_H;
    this.onSelectionChange = options.onSelectionChange;
    this.onFrame = options.onFrame;
    this.onCommitCell = options.onCommitCell;
    this.onCopyCells = options.onCopyCells;
    this.onPasteRequest = options.onPasteRequest;

    const initialPosition = getComputedStyle(this.container).position;
    this.containerPosition = initialPosition;
    this.containerOverflow = this.container.style.overflow;
    this.containerTabIndex = this.container.tabIndex;
    this.containerOutline = this.container.style.outline;
    if (initialPosition === "static") this.container.style.position = "relative";
    this.container.style.overflow = "hidden";
    this.container.tabIndex = 0;
    this.container.style.outline = "none";

    this.contentCanvas = document.createElement("canvas");
    this.overlayCanvas = document.createElement("canvas");
    for (const canvas of [this.contentCanvas, this.overlayCanvas]) {
      canvas.style.position = "absolute";
      canvas.style.inset = "0";
    }
    this.container.append(this.contentCanvas, this.overlayCanvas);
    this.contentCtx = this.contentCanvas.getContext("2d")!;
    this.overlayCtx = this.overlayCanvas.getContext("2d")!;

    this.defaultRowHeight = options.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;
    this.defaultColWidth = options.defaultColumnWidth ?? DEFAULT_COL_WIDTH;
    this.projection = new IdentityRowProjection(this.worksheet.rowCount);
    // Row axis spans VISUAL rows; heights resolve through the projection so
    // hidden rows occupy no pixels and custom physical heights still apply.
    this.rows = new AxisMetrics(this.projection.visualRowCount, this.defaultRowHeight, (visualRow) =>
      this.worksheet.getRowHeight(this.projection.visualToPhysical(visualRow)),
    );
    this.cols = new AxisMetrics(this.worksheet.columnCount, this.defaultColWidth, (i) =>
      this.worksheet.getColumnWidth(i),
    );
    this.selection = new SelectionModel(
      () => this.worksheet.rowCount,
      () => this.worksheet.columnCount,
    );

    this.unsubscribe = options.onChange((event) => this.handleChangeEvent(event));
    this.resizeObserver = new ResizeObserver(() => this.resize());
    this.resizeObserver.observe(this.container);
    this.editor = new CellEditor(this.container, {
      onCommit: (text, move) => this.commitEditor(text, move),
      onCancel: () => {
        this.editingCell = null;
        this.container.focus();
      },
    });
    this.handleWheelBound = (e) => this.onWheel(e);
    this.handleMouseDownBound = (e) => this.onMouseDown(e);
    this.handleDblClickBound = (e) => this.onDblClick(e);
    this.handleMouseMoveBound = (e) => this.onMouseMove(e);
    this.handleMouseUpBound = () => {
      this.drag = null;
    };
    this.handleKeyDownBound = (e) => this.onKeyDown(e);
    this.attachInputHandlers();
    this.resize();
  }

  /** Current selection state (host toolbar / copy reads this). */
  getSelectionState(): import("./selection.js").SelectionState {
    return this.selection.state;
  }

  /** Return keyboard focus to the grid after a host-owned overlay closes. */
  focus(): void {
    this.container.focus();
  }

  /**
   * Select a physical cell supplied by a host feature such as Find, then
   * reveal it through the current visual row projection.  The renderer keeps
   * the physical/visual conversion private: callers never need canvas rows.
   */
  setActiveCell(cell: { row: number; col: number }): void {
    if (
      !Number.isSafeInteger(cell.row) ||
      !Number.isSafeInteger(cell.col) ||
      cell.row < 0 ||
      cell.col < 0 ||
      cell.row >= this.worksheet.rowCount ||
      cell.col >= this.worksheet.columnCount
    ) {
      throw new SheetError("E_INVALID_RANGE", "Active cell is outside worksheet bounds");
    }
    if (!this.projection.isVisible(cell.row)) {
      throw new SheetError("E_VALIDATION", "Cannot select a row hidden by the active projection");
    }
    this.selection.setActive(cell);
    this.scrollActiveIntoView();
    this.notifySelection();
    this.scheduleOverlay();
  }

  /** Current projection (identity unless the host installed a filtered one). */
  getRowProjection(): RowProjection {
    return this.projection;
  }

  /**
   * Install a new row projection (M4.1; `null` restores identity, e.g. when a
   * filter is cleared). Rebuilds the visual row axis and applies the
   * hidden-active-cell policy: an active cell that became hidden moves to the
   * next visible physical row so the overlay, F2 and keyboard navigation
   * always have a visible anchor — then scrolls it into the viewport.
   *
   * Order (M4.1.1): validate → install → rebuild metrics → clamp scroll →
   * relocate active cell → scroll active into view → sync layout → redraw.
   */
  setRowProjection(projection: RowProjection | null): void {
    if (projection !== null && projection.physicalRowCount !== this.worksheet.rowCount) {
      // A projection built for another sheet (or before a structural change)
      // would map onto out-of-bounds physical rows — reject, never install.
      throw new SheetError(
        "E_VALIDATION",
        `RowProjection physicalRowCount ${projection.physicalRowCount} does not match worksheet rowCount ${this.worksheet.rowCount}`,
      );
    }
    this.projection = projection ?? new IdentityRowProjection(this.worksheet.rowCount);
    this.rebuildMetrics();
    // The visual axis may have shrunk drastically (e.g. filter hid 90% of
    // rows): a stale deep scrollY would paint an empty canvas.
    this.reconcileViewportAfterMetrics();
    const active = this.selection.state.active;
    if (!this.projection.isVisible(active.row)) {
      const relocated = relocateToVisibleRow(this.projection, active.row);
      if (relocated !== undefined) {
        this.selection.setActive({ row: relocated, col: active.col });
      }
    }
    this.selection.clampSelection();
    // "Visible in the projection" is not "visible in the viewport" — reveal
    // the (possibly relocated) active cell before the next frame.
    this.scrollActiveIntoView();
    // Refresh layout synchronously: a click before the next rAF must never
    // hit-test against stale frozen geometry (same rule as structure events).
    this.layout = this.computeLayout();
    this.dirty.markFullRedraw();
    this.notifySelection();
    this.scheduleFrame();
  }

  destroy(): void {
    this.destroyed = true;
    if (this.rafId !== 0) cancelAnimationFrame(this.rafId);
    this.unsubscribe();
    this.resizeObserver.disconnect();
    this.container.removeEventListener("wheel", this.handleWheelBound);
    this.container.removeEventListener("keydown", this.handleKeyDownBound);
    this.overlayCanvas.removeEventListener("mousedown", this.handleMouseDownBound);
    this.overlayCanvas.removeEventListener("dblclick", this.handleDblClickBound);
    globalThis.removeEventListener("mousemove", this.handleMouseMoveBound);
    globalThis.removeEventListener("mouseup", this.handleMouseUpBound);
    this.editor.destroy();
    this.contentCanvas.remove();
    this.overlayCanvas.remove();
    // Restore container attributes the renderer modified.
    this.container.style.position = this.containerPosition;
    this.container.style.overflow = this.containerOverflow;
    this.container.style.outline = this.containerOutline;
    this.container.tabIndex = this.containerTabIndex;
  }

  // --- change events --------------------------------------------------------

  private handleChangeEvent(event: ChangeEvent): void {
    if (event.sheetId !== this.worksheet.id) return;
    this.dirty.pushEvent(event);
    // MVP rule (docs/m4-data-operations.md): row/column insert/delete clears
    // an active filter — the projection resets to identity rather than
    // rewriting filter-range coordinates. The host re-applies if needed.
    for (const change of event.changes) {
      if (change.kind === "rows" || change.kind === "columns") {
        this.projection = new IdentityRowProjection(this.worksheet.rowCount);
        break;
      }
    }
    if (this.dirty.needsStructureRebuild) {
      this.rebuildMetrics();
      // Deleted rows/cols shrink the axes: clamp scroll before any hit test
      // or paint reads the stale offset (M4.1.1).
      this.reconcileViewportAfterMetrics();
      // Row/col/sheet structure changed: the active cell may now be out of
      // bounds. SelectionModel re-clamps on next set; force it now so the
      // overlay does not paint a stale out-of-range selection.
      this.selection.clampSelection();
      this.notifySelection();
      // Refresh the cached layout IMMEDIATELY — hit-testing (e.g. a click
      // arriving before the next rAF) must never use stale frozen geometry.
      this.layout = this.computeLayout();
    }
    this.scheduleFrame();
  }

  private rebuildMetrics(): void {
    const projection = this.projection;
    this.rows.rebuild(projection.visualRowCount, this.defaultRowHeight, (visualRow) =>
      this.worksheet.getRowHeight(projection.visualToPhysical(visualRow)),
    );
    this.cols.rebuild(this.worksheet.columnCount, this.defaultColWidth, (i) =>
      this.worksheet.getColumnWidth(i),
    );
  }

  // --- sizing ---------------------------------------------------------------

  private resize(): void {
    const rect = this.container.getBoundingClientRect();
    const ratio = globalThis.devicePixelRatio || 1;
    for (const [canvas, ctx] of [
      [this.contentCanvas, this.contentCtx],
      [this.overlayCanvas, this.overlayCtx],
    ] as const) {
      canvas.style.width = `${rect.width}px`;
      canvas.style.height = `${rect.height}px`;
      canvas.width = Math.max(1, Math.round(rect.width * ratio));
      canvas.height = Math.max(1, Math.round(rect.height * ratio));
      ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
    }
    // A larger viewport shrinks maxScroll (totalSize - viewport): re-clamp so
    // the scroll offset never points past the content (M4.1.1).
    this.reconcileViewportAfterMetrics();
    this.dirty.markFullRedraw();
    this.scheduleFrame();
  }

  /**
   * Re-clamp the current scroll offset against the CURRENT axis sizes and
   * viewport. Must run after every metrics rebuild or resize — unlike
   * setScroll it does not schedule a frame (callers handle invalidation).
   */
  private reconcileViewportAfterMetrics(): void {
    const clamped = clampScroll(
      { scrollX: this.scrollX, scrollY: this.scrollY },
      this.rows,
      this.cols,
      this.contentWidth() - this.headerWidth,
      this.contentHeight() - this.headerHeight,
    );
    this.scrollX = clamped.scrollX;
    this.scrollY = clamped.scrollY;
  }

  // --- scrolling ------------------------------------------------------------

  private setScroll(scrollX: number, scrollY: number): void {
    // viewport size excludes headers but includes the frozen zone; clampScroll
    // derives maxScroll = totalSize - viewport (frozen cancels out), matching
    // computeScrollbarGeometry().maxScroll exactly.
    const clamped = clampScroll(
      { scrollX, scrollY },
      this.rows,
      this.cols,
      this.contentWidth() - this.headerWidth,
      this.contentHeight() - this.headerHeight,
    );
    if (clamped.scrollX === this.scrollX && clamped.scrollY === this.scrollY) return;
    this.scrollX = clamped.scrollX;
    this.scrollY = clamped.scrollY;
    this.dirty.markFullRedraw(); // viewport content is entirely new
    this.scheduleFrame();
  }

  private contentWidth(): number {
    return this.container.clientWidth;
  }

  private contentHeight(): number {
    return this.container.clientHeight;
  }

  // --- frame scheduling -----------------------------------------------------

  private scheduleFrame(): void {
    if (this.rafId !== 0 || this.destroyed) return;
    this.rafId = requestAnimationFrame(() => {
      this.rafId = 0;
      this.renderFrame();
    });
  }

  private computeLayout(): ViewportLayout {
    return computeViewport({
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      width: this.contentWidth(),
      height: this.contentHeight(),
      rows: this.rows,
      cols: this.cols,
      // Viewport frozen counts are VISUAL: hidden rows inside the physical
      // frozen zone occupy no height, visible frozen rows stay pinned.
      frozenRowCount: this.projection.visibleCountBefore(this.worksheet.frozenRows),
      frozenColCount: this.worksheet.frozenColumns,
      bufferPx: BUFFER_PX,
      headerWidth: this.headerWidth,
      headerHeight: this.headerHeight,
    });
  }

  private renderFrame(): void {
    if (this.destroyed) return;
    const t0 = performance.now();
    this.paintedCells = 0;
    this.layout = this.computeLayout();
    const { full, rects } = this.dirty.consume(this.layout, this.rows, this.cols, this.projection);
    if (full) {
      this.paintAll(this.layout);
    } else {
      // Dirty cell rects and the header highlight are independent: a single
      // frame can carry both (e.g. editor commit + Enter moves the selection),
      // so each gets its own branch instead of an else-if chain.
      if (rects.length > 0) {
        this.paintDirtyRects(this.layout, rects);
      }
      if (this.headerDirty) {
        this.paintHeaders(this.layout);
      }
    }
    this.headerDirty = false;
    // Overlay is cheap: repaint on any content change too (selection geometry
    // may shift with scroll), or when selection itself changed.
    this.paintOverlay(this.layout);
    this.onFrame?.({ paintMs: performance.now() - t0, full, paintedCells: this.paintedCells });
  }

  // --- content painting -----------------------------------------------------

  private paintAll(layout: ViewportLayout): void {
    const ctx = this.contentCtx;
    ctx.save();
    ctx.clearRect(0, 0, this.contentWidth(), this.contentHeight());
    ctx.fillStyle = this.theme.background;
    ctx.fillRect(0, 0, this.contentWidth(), this.contentHeight());
    this.paintQuadrants(layout, null);
    this.paintHeaders(layout);
    ctx.restore();
  }

  private paintDirtyRects(layout: ViewportLayout, rects: PixelRect[]): void {
    const ctx = this.contentCtx;
    ctx.save();
    for (const rect of rects) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(rect.x, rect.y, rect.width, rect.height);
      ctx.clip();
      ctx.fillStyle = this.theme.background;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      this.paintQuadrants(layout, rect);
      ctx.restore();
    }
    ctx.restore();
  }

  private paintQuadrants(layout: ViewportLayout, dirtyRect: PixelRect | null): void {
    for (const quadrant of [layout.corner, layout.top, layout.left, layout.main]) {
      if (quadrant === null) continue;
      this.paintQuadrantCells(quadrant, dirtyRect);
    }
  }

  private paintQuadrantCells(
    q: Quadrant,
    dirtyRect: PixelRect | null,
  ): void {
    const ctx = this.contentCtx;
    ctx.save();
    ctx.beginPath();
    if (dirtyRect !== null) {
      ctx.rect(dirtyRect.x, dirtyRect.y, dirtyRect.width, dirtyRect.height);
    }
    ctx.rect(q.clipX, q.clipY, q.clipWidth, q.clipHeight);
    ctx.clip();

    // Limit iteration to rows/cols overlapping the dirty rect.
    let rowStart = q.rowStart;
    let rowEnd = q.rowEnd;
    let colStart = q.colStart;
    let colEnd = q.colEnd;
    if (dirtyRect !== null) {
      const baseRowPos = this.rows.positionOf(q.rowStart);
      const baseColPos = this.cols.positionOf(q.colStart);
      rowStart = Math.max(
        q.rowStart,
        this.rows.indexAt(baseRowPos + (dirtyRect.y - q.originY)) ,
      );
      rowEnd = Math.min(
        q.rowEnd,
        this.rows.indexAt(baseRowPos + (dirtyRect.y + dirtyRect.height - q.originY)),
      );
      colStart = Math.max(
        q.colStart,
        this.cols.indexAt(baseColPos + (dirtyRect.x - q.originX)),
      );
      colEnd = Math.min(
        q.colEnd,
        this.cols.indexAt(baseColPos + (dirtyRect.x + dirtyRect.width - q.originX)),
      );
      if (rowStart > rowEnd || colStart > colEnd) {
        ctx.restore();
        return;
      }
    }

    const theme = this.theme;
    const sheet = this.worksheet;
    const projection = this.projection;

    // 1. backgrounds — quadrant indices are VISUAL rows; data access maps
    // each visual row back to its physical row exactly once.
    for (let visualRow = rowStart; visualRow <= rowEnd; visualRow++) {
      const physicalRow = projection.visualToPhysical(visualRow);
      const y = q.originY + (this.rows.positionOf(visualRow) - this.rows.positionOf(q.rowStart));
      const h = this.rows.sizeOf(visualRow);
      for (let col = colStart; col <= colEnd; col++) {
        const cell = sheet.getCell(physicalRow, col);
        const style = cell?.styleId !== undefined ? this.resolveStyleFn?.(cell.styleId) : undefined;
        if (style?.backgroundColor === undefined) continue;
        const x = q.originX + (this.cols.positionOf(col) - this.cols.positionOf(q.colStart));
        ctx.fillStyle = style.backgroundColor;
        ctx.fillRect(x, y, this.cols.sizeOf(col), h);
      }
    }

    // 2. grid lines (drawn for the visible window of this quadrant)
    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let visualRow = rowStart; visualRow <= rowEnd + 1; visualRow++) {
      const y = Math.round(q.originY + (this.rows.positionOf(visualRow) - this.rows.positionOf(q.rowStart))) + 0.5;
      ctx.moveTo(q.clipX, y);
      ctx.lineTo(q.clipX + q.clipWidth, y);
    }
    for (let col = colStart; col <= colEnd + 1; col++) {
      const x = Math.round(q.originX + (this.cols.positionOf(col) - this.cols.positionOf(q.colStart))) + 0.5;
      ctx.moveTo(x, q.clipY);
      ctx.lineTo(x, q.clipY + q.clipHeight);
    }
    ctx.stroke();

    // 3. cell text
    ctx.textBaseline = "middle";
    for (let visualRow = rowStart; visualRow <= rowEnd; visualRow++) {
      const physicalRow = projection.visualToPhysical(visualRow);
      const y = q.originY + (this.rows.positionOf(visualRow) - this.rows.positionOf(q.rowStart));
      const h = this.rows.sizeOf(visualRow);
      for (let col = colStart; col <= colEnd; col++) {
        const cell = sheet.getCell(physicalRow, col);
        this.paintedCells++;
        if (cell === undefined || cell.value === null) continue;
        const style = cell.styleId !== undefined ? this.resolveStyleFn?.(cell.styleId) : undefined;
        const x = q.originX + (this.cols.positionOf(col) - this.cols.positionOf(q.colStart));
        const w = this.cols.sizeOf(col);
        this.paintCellText(ctx, cell.value, style, x, y, w, h);
      }
    }
    ctx.restore();
  }

  private paintCellText(
    ctx: CanvasRenderingContext2D,
    value: import("@injoysai/opensheet-shared").CellValue,
    style: Readonly<CellStyle> | undefined,
    x: number,
    y: number,
    w: number,
    h: number,
  ): void {
    let text: string;
    let color = style?.textColor ?? this.theme.cellText;
    if (isCellError(value)) {
      text = value.type;
      color = this.theme.cellErrorText;
    } else if (typeof value === "boolean") {
      text = value ? "TRUE" : "FALSE";
    } else {
      text = String(value);
    }

    const fontSize = style?.fontSize ?? 13;
    const weight = style?.bold === true ? "600" : "400";
    const italic = style?.italic === true ? "italic " : "";
    ctx.font = `${italic}${weight} ${fontSize}px -apple-system, "Segoe UI", Roboto, sans-serif`;
    ctx.fillStyle = color;

    const align = style?.horizontalAlign ?? (typeof value === "number" ? "right" : "left");
    const textY = y + h / 2;
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
    if (align === "right") {
      ctx.textAlign = "right";
      ctx.fillText(text, x + w - CELL_PAD_X, textY);
    } else if (align === "center") {
      ctx.textAlign = "center";
      ctx.fillText(text, x + w / 2, textY);
    } else {
      ctx.textAlign = "left";
      ctx.fillText(text, x + CELL_PAD_X, textY);
    }
    ctx.restore();
  }

  private paintHeaders(layout: ViewportLayout): void {
    const ctx = this.contentCtx;
    const theme = this.theme;
    const selection = this.selection.state.range;

    ctx.save();
    ctx.font = theme.headerFont;
    ctx.textBaseline = "middle";

    const paintColHeaders = (colStart: number, colEnd: number, originX: number, baseCol: number, clipX: number, clipW: number) => {
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, 0, clipW, this.headerHeight);
      ctx.clip();
      ctx.fillStyle = theme.headerBackground;
      ctx.fillRect(clipX, 0, clipW, this.headerHeight);
      for (let col = colStart; col <= colEnd; col++) {
        const x = originX + (this.cols.positionOf(col) - this.cols.positionOf(baseCol));
        const w = this.cols.sizeOf(col);
        const highlighted = col >= selection.startCol && col <= selection.endCol;
        if (highlighted) {
          ctx.fillStyle = theme.headerHighlight;
          ctx.fillRect(x, 0, w, this.headerHeight);
        }
        ctx.fillStyle = highlighted ? theme.headerHighlightText : theme.headerText;
        ctx.textAlign = "center";
        ctx.fillText(colName(col), x + w / 2, this.headerHeight / 2);
        ctx.strokeStyle = theme.gridLine;
        ctx.beginPath();
        ctx.moveTo(Math.round(x) + 0.5, 0);
        ctx.lineTo(Math.round(x) + 0.5, this.headerHeight);
        ctx.stroke();
      }
      ctx.restore();
    };

    const paintRowHeaders = (rowStart: number, rowEnd: number, originY: number, baseRow: number, clipY: number, clipH: number) => {
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, clipY, this.headerWidth, clipH);
      ctx.clip();
      ctx.fillStyle = theme.headerBackground;
      ctx.fillRect(0, clipY, this.headerWidth, clipH);
      // rowStart..rowEnd are VISUAL rows. Headers display the PHYSICAL row
      // number (filtered sheets keep their original numbering), and the
      // selection highlight compares physical rows (SelectionModel is
      // physical) — never the compacted visual index.
      for (let visualRow = rowStart; visualRow <= rowEnd; visualRow++) {
        const physicalRow = this.projection.visualToPhysical(visualRow);
        const y = originY + (this.rows.positionOf(visualRow) - this.rows.positionOf(baseRow));
        const h = this.rows.sizeOf(visualRow);
        const highlighted = physicalRow >= selection.startRow && physicalRow <= selection.endRow;
        if (highlighted) {
          ctx.fillStyle = theme.headerHighlight;
          ctx.fillRect(0, y, this.headerWidth, h);
        }
        ctx.fillStyle = highlighted ? theme.headerHighlightText : theme.headerText;
        ctx.textAlign = "center";
        ctx.fillText(String(physicalRow + 1), this.headerWidth / 2, y + h / 2);
        ctx.strokeStyle = theme.gridLine;
        ctx.beginPath();
        ctx.moveTo(0, Math.round(y) + 0.5);
        ctx.lineTo(this.headerWidth, Math.round(y) + 0.5);
        ctx.stroke();
      }
      ctx.restore();
    };

    // Column headers: frozen strip + main strip.
    if (layout.left !== null || layout.corner !== null) {
      const frozenCols = this.worksheet.frozenColumns;
      if (frozenCols > 0) {
        paintColHeaders(0, frozenCols - 1, this.headerWidth, 0, this.headerWidth, layout.frozenWidth);
      }
    }
    paintColHeaders(
      layout.main.colStart,
      layout.main.colEnd,
      layout.main.originX,
      layout.main.colStart,
      layout.mainX,
      layout.mainWidth,
    );

    // Row headers: frozen strip + main strip. The frozen strip spans the
    // VISIBLE frozen rows (hidden rows inside the physical frozen zone
    // occupy no header pixels either).
    const visibleFrozenRows = this.projection.visibleCountBefore(this.worksheet.frozenRows);
    if (visibleFrozenRows > 0) {
      paintRowHeaders(0, visibleFrozenRows - 1, this.headerHeight, 0, this.headerHeight, layout.frozenHeight);
    }
    paintRowHeaders(
      layout.main.rowStart,
      layout.main.rowEnd,
      layout.main.originY,
      layout.main.rowStart,
      layout.mainY,
      layout.mainHeight,
    );

    // Header corner box.
    ctx.fillStyle = theme.headerBackground;
    ctx.fillRect(0, 0, this.headerWidth, this.headerHeight);
    ctx.strokeStyle = theme.gridLine;
    ctx.strokeRect(0.5, 0.5, this.headerWidth, this.headerHeight);

    // Header separator lines.
    ctx.beginPath();
    ctx.moveTo(0, this.headerHeight + 0.5);
    ctx.lineTo(this.contentWidth(), this.headerHeight + 0.5);
    ctx.moveTo(this.headerWidth + 0.5, 0);
    ctx.lineTo(this.headerWidth + 0.5, this.contentHeight());
    ctx.stroke();
    ctx.restore();
  }

  // --- overlay painting -----------------------------------------------------

  private paintOverlay(layout: ViewportLayout): void {
    const ctx = this.overlayCtx;
    const theme = this.theme;
    ctx.save();
    ctx.clearRect(0, 0, this.contentWidth(), this.contentHeight());

    // Selection fill + border. SelectionModel is PHYSICAL; the overlay maps
    // the range onto the visual axis (hidden rows inside occupy no pixels,
    // so the fill stays continuous — standard spreadsheet behavior).
    const selection = this.selection.state;
    const visualSelection = physicalRangeToVisualRange(selection.range, this.projection);
    const rects = visualSelection === null
      ? []
      : rangeToCanvasRects(visualSelection, layout, this.rows, this.cols);
    for (const rect of rects) {
      ctx.fillStyle = theme.selectionFill;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.strokeStyle = theme.selectionBorder;
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
    }

    // Active cell border + fill handle (visible by policy; if it ever ends
    // up hidden — e.g. projection swapped under a stale selection — skip
    // rather than paint at a wrong visual slot).
    const activeVisualRow = this.projection.physicalToVisual(selection.active.row);
    if (activeVisualRow !== undefined) {
      const activeRects = rangeToCanvasRects(
        {
          startRow: activeVisualRow,
          startCol: selection.active.col,
          endRow: activeVisualRow,
          endCol: selection.active.col,
        },
        layout,
        this.rows,
        this.cols,
      );
      for (const rect of activeRects) {
        ctx.strokeStyle = theme.activeCellBorder;
        ctx.lineWidth = 2;
        ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
        ctx.lineWidth = 1;
        ctx.fillStyle = theme.activeCellBorder;
        ctx.fillRect(rect.x + rect.width - 4, rect.y + rect.height - 4, 5, 5);
      }
    }

    // Frozen dividers (drawn when any VISIBLE frozen strip exists).
    ctx.strokeStyle = theme.frozenDivider;
    ctx.beginPath();
    if (this.worksheet.frozenColumns > 0) {
      const x = layout.mainX + 0.5;
      ctx.moveTo(x, this.headerHeight);
      ctx.lineTo(x, this.contentHeight());
    }
    if (this.projection.visibleCountBefore(this.worksheet.frozenRows) > 0) {
      const y = layout.mainY + 0.5;
      ctx.moveTo(this.headerWidth, y);
      ctx.lineTo(this.contentWidth(), y);
    }
    ctx.stroke();

    this.paintScrollbars();
    ctx.restore();
  }

  private paintScrollbars(): void {
    const ctx = this.overlayCtx;
    const theme = this.theme;
    const geometry = this.scrollbarGeometry();

    if (geometry.vertical !== null) {
      const v = geometry.vertical;
      ctx.fillStyle = theme.scrollbarTrack;
      ctx.fillRect(v.x, v.trackStart, SCROLLBAR, v.trackSize);
      ctx.fillStyle = theme.scrollbarThumb;
      ctx.fillRect(v.x + 2, v.y, SCROLLBAR - 4, v.height);
    }
    if (geometry.horizontal !== null) {
      const hh = geometry.horizontal;
      ctx.fillStyle = theme.scrollbarTrack;
      ctx.fillRect(hh.trackStart, hh.y, hh.trackSize, SCROLLBAR);
      ctx.fillStyle = theme.scrollbarThumb;
      ctx.fillRect(hh.x, hh.y + 2, hh.width, SCROLLBAR - 4);
    }
  }

  private scrollbarGeometry(): ScrollbarGeometry {
    const layout = this.layout ?? this.computeLayout();
    return computeScrollbarGeometry({
      layout,
      rows: this.rows,
      cols: this.cols,
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      width: this.contentWidth(),
      height: this.contentHeight(),
      headerWidth: this.headerWidth,
      headerHeight: this.headerHeight,
      scrollbarSize: SCROLLBAR,
    });
  }

  // --- input ----------------------------------------------------------------

  private attachInputHandlers(): void {
    this.container.addEventListener("wheel", this.handleWheelBound, { passive: false });
    this.overlayCanvas.addEventListener("mousedown", this.handleMouseDownBound);
    this.overlayCanvas.addEventListener("dblclick", this.handleDblClickBound);
    globalThis.addEventListener("mousemove", this.handleMouseMoveBound);
    globalThis.addEventListener("mouseup", this.handleMouseUpBound);
    this.container.addEventListener("keydown", this.handleKeyDownBound);
  }

  private onWheel(e: WheelEvent): void {
    e.preventDefault();
    const unit = e.deltaMode === 1 ? 16 : 1;
    const dx = (e.shiftKey ? e.deltaY : e.deltaX) * unit;
    const dy = (e.shiftKey ? 0 : e.deltaY) * unit;
    this.setScroll(this.scrollX + dx, this.scrollY + dy);
  }

  private canvasPoint(e: MouseEvent): { x: number; y: number } {
    const rect = this.overlayCanvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }

  private onMouseDown(e: MouseEvent): void {
    if (e.button !== 0) return;
    this.container.focus();
    const { x, y } = this.canvasPoint(e);

    // Scrollbar hit test.
    const geometry = this.scrollbarGeometry();
    if (geometry.vertical !== null && x >= geometry.vertical.x && y >= geometry.vertical.trackStart) {
      this.drag = { kind: "v-scroll", grabOffset: y - geometry.vertical.y };
      return;
    }
    if (geometry.horizontal !== null && y >= geometry.horizontal.y && x >= geometry.horizontal.trackStart) {
      this.drag = { kind: "h-scroll", grabOffset: x - geometry.horizontal.x };
      return;
    }

    const hit = this.hitTest(x, y);
    if (hit.zone !== "cell") return;
    if (e.shiftKey) {
      this.selection.extendTo(hit);
    } else {
      this.selection.setActive(hit);
    }
    this.drag = { kind: "select" };
    this.notifySelection();
    this.scheduleOverlay();
  }

  private onMouseMove(e: MouseEvent): void {
    if (this.drag === null) return;
    const { x, y } = this.canvasPoint(e);
    const geometry = this.scrollbarGeometry();
    if (this.drag.kind === "v-scroll" && geometry.vertical !== null) {
      const v = geometry.vertical;
      const thumbTop = y - this.drag.grabOffset;
      const ratio = (thumbTop - v.trackStart) / Math.max(1, v.trackSize - v.height);
      this.setScroll(this.scrollX, ratio * v.maxScroll);
      return;
    }
    if (this.drag.kind === "h-scroll" && geometry.horizontal !== null) {
      const hh = geometry.horizontal;
      const thumbLeft = x - this.drag.grabOffset;
      const ratio = (thumbLeft - hh.trackStart) / Math.max(1, hh.trackSize - hh.width);
      this.setScroll(ratio * hh.maxScroll, this.scrollY);
      return;
    }
    // Drag-select with edge auto-scroll.
    const EDGE = 20;
    const step = 24;
    let dx = 0;
    let dy = 0;
    if (x > this.contentWidth() - EDGE) dx = step;
    else if (x < this.headerWidth + EDGE) dx = -step;
    if (y > this.contentHeight() - EDGE) dy = step;
    else if (y < this.headerHeight + EDGE) dy = -step;
    if (dx !== 0 || dy !== 0) this.setScroll(this.scrollX + dx, this.scrollY + dy);
    const hit = this.hitTest(x, y);
    if (hit.zone === "cell") {
      this.selection.extendTo(hit);
      this.notifySelection();
      this.scheduleOverlay();
    }
  }

  private onKeyDown(e: KeyboardEvent): void {
    // While the inline editor is open it owns ALL keys (Enter/Tab/Escape/IME
    // live in the textarea; composition keydowns bubble here and must not
    // trigger grid navigation). Blur-commit happens before this returns.
    if (this.editor.isActive) return;

    const shift = e.shiftKey;
    const meta = e.metaKey || e.ctrlKey;

    // Clipboard: Cmd/Ctrl+C and Cmd/Ctrl+V (also works mid-edit? No — the
    // textarea handles its own copy/paste; grid handles selection copy/paste).
    if (meta && !shift && !e.altKey) {
      if (e.key.toLowerCase() === "c") {
        e.preventDefault();
        this.copySelection();
        return;
      }
      if (e.key.toLowerCase() === "v") {
        e.preventDefault();
        this.pasteIntoSelection();
        return;
      }
    }

    // Inline editing start: F2 (edit current value) or a printable char
    // (replace cell content, Excel-style). Keep grid navigation keys out.
    const action = decideKeyInPhase("idle", {
      key: e.key,
      shift,
      ctrl: meta,
      meta,
      alt: e.altKey,
    });
    if (action.kind === "start-editing") {
      e.preventDefault();
      const active = this.selection.state.active;
      const initial = isPrintableKey(e.key, { ctrl: meta, meta, alt: e.altKey })
        ? e.key
        : cellDisplayText(this.worksheet.getCell(active.row, active.col)?.formula, this.worksheet.getCell(active.row, active.col)?.value);
      this.startEditing(active, initial);
      return;
    }

    const pageRows = Math.max(
      1,
      this.rows.indexAt(this.scrollY + (this.layout?.mainHeight ?? 400)) - this.rows.indexAt(this.scrollY) - 1,
    );
    let handled = true;
    switch (e.key) {
      case "ArrowUp":
        this.moveActiveRow(-1, shift);
        break;
      case "ArrowDown":
        this.moveActiveRow(1, shift);
        break;
      case "ArrowLeft":
        this.selection.moveBy(0, -1, shift);
        break;
      case "ArrowRight":
        this.selection.moveBy(0, 1, shift);
        break;
      case "Tab":
        this.selection.moveBy(0, shift ? -1 : 1, false);
        break;
      case "Enter":
        this.moveActiveRow(shift ? -1 : 1, false);
        break;
      case "PageUp":
        this.pageActiveRow(-1, shift, pageRows);
        break;
      case "PageDown":
        this.pageActiveRow(1, shift, pageRows);
        break;
      case "Home":
        if (meta) {
          // Ctrl+Home: first VISIBLE physical row (identity → row 0).
          if (this.projection.visualRowCount > 0) {
            this.selection.jumpTo({ row: this.projection.visualToPhysical(0), col: 0 }, shift);
          }
        } else this.selection.jumpTo({ row: this.selection.state.active.row, col: 0 }, shift);
        break;
      case "End":
        if (meta) {
          // Ctrl+End: last VISIBLE physical row, not the sheet's last row.
          if (this.projection.visualRowCount > 0) {
            this.selection.jumpTo(
              { row: lastVisiblePhysicalRow(this.projection), col: this.worksheet.columnCount - 1 },
              shift,
            );
          }
        } else
          this.selection.jumpTo(
            { row: this.selection.state.active.row, col: this.worksheet.columnCount - 1 },
            shift,
          );
        break;
      default:
        handled = false;
    }
    if (!handled) return;
    e.preventDefault();
    this.scrollActiveIntoView();
    this.notifySelection();
    this.scheduleOverlay();
  }

  /**
   * Vertical keyboard move over VISIBLE rows (M4.1): hidden rows are skipped
   * via the projection. Without shift the move starts from the range edge in
   * the direction of travel and collapses the selection (desktop behavior);
   * with shift it extends from the focus cell. SelectionModel stays physical.
   */
  private moveActiveRow(direction: 1 | -1, extend: boolean): void {
    const state = this.selection.state;
    const baseRow = extend
      ? state.active.row
      : direction > 0
        ? state.range.endRow
        : state.range.startRow;
    const nextRow = this.projection.nextVisible(baseRow, direction) ?? baseRow;
    if (extend) {
      this.selection.extendTo({ row: nextRow, col: state.active.col });
    } else {
      this.selection.setActive({ row: nextRow, col: state.active.col });
    }
  }

  /** Page moves use visual arithmetic: ±pageRows on the visual axis. */
  private pageActiveRow(direction: 1 | -1, extend: boolean, pageRows: number): void {
    const state = this.selection.state;
    const baseRow = extend
      ? state.active.row
      : direction > 0
        ? state.range.endRow
        : state.range.startRow;
    const baseVisual =
      this.projection.physicalToVisual(baseRow) ??
      (() => {
        const relocated = relocateToVisibleRow(this.projection, baseRow);
        return relocated === undefined ? undefined : this.projection.physicalToVisual(relocated);
      })();
    if (baseVisual === undefined) return; // nothing visible — degenerate
    const targetVisual = Math.min(
      Math.max(0, baseVisual + direction * pageRows),
      this.projection.visualRowCount - 1,
    );
    const targetRow = this.projection.visualToPhysical(targetVisual);
    if (extend) {
      this.selection.extendTo({ row: targetRow, col: state.active.col });
    } else {
      this.selection.setActive({ row: targetRow, col: state.active.col });
    }
  }

  // --- inline editing --------------------------------------------------------

  private onDblClick(e: MouseEvent): void {
    if (e.button !== 0) return;
    const { x, y } = this.canvasPoint(e);
    const hit = this.hitTest(x, y);
    if (hit.zone !== "cell") return;
    const cell = this.worksheet.getCell(hit.row, hit.col);
    this.startEditing(hit, cellDisplayText(cell?.formula, cell?.value));
  }

  private startEditing(cell: { row: number; col: number }, initialText: string): void {
    const layout = this.layout ?? this.computeLayout();
    // cellRectInCanvas positions over the VISUAL axis; a hidden cell has no
    // visual slot, so it cannot be edited (kept unreachable by the
    // hidden-active-cell policy — this is the defensive guard).
    const visualRow = this.projection.physicalToVisual(cell.row);
    if (visualRow === undefined) return;
    const rect = cellRectInCanvas({ row: visualRow, col: cell.col }, layout, this.rows, this.cols);
    this.editingCell = { row: cell.row, col: cell.col };
    this.editor.open(rect, initialText);
  }

  private commitEditor(text: string, move: { row: number; col: number } | null): void {
    const cell = this.editingCell;
    this.editingCell = null;
    if (cell === null) return;
    // Editor only READS; the host performs the write through applyOperations.
    this.onCommitCell?.({ row: cell.row, col: cell.col, text });
    if (move !== null) {
      if (move.row !== 0) {
        this.moveActiveRow(move.row > 0 ? 1 : -1, false);
      } else if (move.col !== 0) {
        this.selection.moveBy(0, move.col, false);
      }
      this.scrollActiveIntoView();
      this.notifySelection();
      this.scheduleOverlay();
    }
    this.container.focus();
  }

  // --- clipboard (selection copy / TSV paste) --------------------------------
  // The grid only READS the selection and hands raw cells to the host; the
  // host owns clipboard I/O and the atomic write (M2 semantic: one paste =
  // one history entry).

  private copySelection(): void {
    const range = this.selection.state.range;
    const rows: CellPrimitive[][] = [];
    for (let row = range.startRow; row <= range.endRow; row++) {
      const line: CellPrimitive[] = [];
      for (let col = range.startCol; col <= range.endCol; col++) {
        line.push(cellPrimitiveOf(this.worksheet.getCell(row, col)?.value));
      }
      rows.push(line);
    }
    this.onCopyCells?.(rows);
  }

  private pasteIntoSelection(): void {
    const active = this.selection.state.active;
    this.onPasteRequest?.({ row: active.row, col: active.col });
  }

  private scrollActiveIntoView(): void {
    const active = this.selection.state.active;
    // computeScrollToCell works on the VISUAL axis (positions come from the
    // row AxisMetrics); a hidden active cell has no visual slot to reveal.
    const activeVisualRow = this.projection.physicalToVisual(active.row);
    if (activeVisualRow === undefined) return;
    const next = computeScrollToCell({ row: activeVisualRow, col: active.col }, { scrollX: this.scrollX, scrollY: this.scrollY }, {
      viewportWidth: this.contentWidth() - this.headerWidth,
      viewportHeight: this.contentHeight() - this.headerHeight,
      rows: this.rows,
      cols: this.cols,
      frozenRowCount: this.projection.visibleCountBefore(this.worksheet.frozenRows),
      frozenColCount: this.worksheet.frozenColumns,
    });
    this.setScroll(next.scrollX, next.scrollY);
  }

  private hitTest(x: number, y: number): import("./coordinate-mapper.js").CellHit {
    const layout = this.layout ?? this.computeLayout();
    const hit = hitTestCell({
      x,
      y,
      layout,
      rows: this.rows,
      cols: this.cols,
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      headerWidth: this.headerWidth,
      headerHeight: this.headerHeight,
      scrollbarSize: SCROLLBAR,
      // Bounds and returned rows are VISUAL (they index the row axis)…
      rowCount: this.rows.length,
      colCount: this.worksheet.columnCount,
    });
    // …but SelectionModel and the host are PHYSICAL — convert exactly once.
    if (hit.zone === "cell" || hit.zone === "rowHeader") {
      return { ...hit, row: this.projection.visualToPhysical(hit.row) };
    }
    return hit;
  }

  private scheduleOverlay(): void {
    // Selection moved: row/col header highlights live on the Content canvas.
    this.headerDirty = true;
    this.scheduleFrame();
  }

  private notifySelection(): void {
    const active = this.selection.state.active;
    this.onSelectionChange?.({ activeRow: active.row, activeCol: active.col });
  }

  // --- scrollbar geometry ---------------------------------------------------


}

/** Cell value → clipboard primitive (errors copy as their display text). */
function cellPrimitiveOf(value: import("@injoysai/opensheet-shared").CellValue | undefined): CellPrimitive {
  if (value === undefined || value === null) return null;
  if (typeof value === "object") {
    // CellError
    return value.message !== undefined ? `${value.type}: ${value.message}` : value.type;
  }
  return value;
}

function colName(col: number): string {
  let name = "";
  let n = col + 1;
  while (n > 0) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}
