// SheetGrid: dual-canvas grid renderer.
//
// Boundary (ADR-0001, M1 guardrail): the renderer receives ONLY a
// WorksheetView — it cannot mutate workbook data. All it emits back is
// selection/scroll state via callbacks.

import type { WorksheetView } from "@opensheet/core";
import type {
  CellStyle,
  ChangeEvent,
  ChangeListener,
  Unsubscribe,
} from "@opensheet/shared";
import { isCellError } from "@opensheet/shared";
import { AxisMetrics } from "./axis-metrics.js";
import { DirtyRegionTracker, rangeToCanvasRects, type PixelRect } from "./dirty-region.js";
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
  private paintedCells = 0;

  private readonly contentCanvas: HTMLCanvasElement;
  private readonly overlayCanvas: HTMLCanvasElement;
  private readonly contentCtx: CanvasRenderingContext2D;
  private readonly overlayCtx: CanvasRenderingContext2D;

  private rows: AxisMetrics;
  private cols: AxisMetrics;
  private scrollX = 0;
  private scrollY = 0;
  private layout: ViewportLayout | null = null;

  private readonly selection: SelectionModel;
  private readonly dirty = new DirtyRegionTracker();
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

    const position = getComputedStyle(this.container).position;
    if (position === "static") this.container.style.position = "relative";
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

    const defaultRowHeight = options.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;
    const defaultColWidth = options.defaultColumnWidth ?? DEFAULT_COL_WIDTH;
    this.rows = new AxisMetrics(this.worksheet.rowCount, defaultRowHeight, (i) =>
      this.worksheet.getRowHeight(i),
    );
    this.cols = new AxisMetrics(this.worksheet.columnCount, defaultColWidth, (i) =>
      this.worksheet.getColumnWidth(i),
    );
    this.selection = new SelectionModel(
      () => this.worksheet.rowCount,
      () => this.worksheet.columnCount,
    );

    this.unsubscribe = options.onChange((event) => this.handleChangeEvent(event));
    this.resizeObserver = new ResizeObserver(() => this.resize());
    this.resizeObserver.observe(this.container);
    this.attachInputHandlers();
    this.resize();
  }

  destroy(): void {
    this.destroyed = true;
    if (this.rafId !== 0) cancelAnimationFrame(this.rafId);
    this.unsubscribe();
    this.resizeObserver.disconnect();
    this.contentCanvas.remove();
    this.overlayCanvas.remove();
  }

  // --- change events --------------------------------------------------------

  private handleChangeEvent(event: ChangeEvent): void {
    if (event.sheetId !== this.worksheet.id) return;
    this.dirty.pushEvent(event);
    if (this.dirty.needsStructureRebuild) {
      this.rebuildMetrics();
    }
    this.scheduleFrame();
  }

  private rebuildMetrics(): void {
    const defaultRowHeight = this.rowsSizeDefault(26);
    const defaultColWidth = this.colsSizeDefault(100);
    this.rows.rebuild(this.worksheet.rowCount, defaultRowHeight, (i) =>
      this.worksheet.getRowHeight(i),
    );
    this.cols.rebuild(this.worksheet.columnCount, defaultColWidth, (i) =>
      this.worksheet.getColumnWidth(i),
    );
  }

  private rowsSizeDefault(fallback: number): number {
    return fallback; // default row height is renderer config; overrides come from the view
  }

  private colsSizeDefault(fallback: number): number {
    return fallback;
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
    this.dirty.markFullRedraw();
    this.scheduleFrame();
  }

  // --- scrolling ------------------------------------------------------------

  private setScroll(scrollX: number, scrollY: number): void {
    const frozenWidth = this.cols.positionOf(this.worksheet.frozenColumns);
    const frozenHeight = this.rows.positionOf(this.worksheet.frozenRows);
    const clamped = clampScroll(
      { scrollX, scrollY },
      this.rows,
      this.cols,
      this.contentWidth() - this.headerWidth,
      this.contentHeight() - this.headerHeight,
      frozenWidth,
      frozenHeight,
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
      frozenRowCount: this.worksheet.frozenRows,
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
    const { full, rects } = this.dirty.consume(this.layout, this.rows, this.cols);
    if (full) {
      this.paintAll(this.layout);
    } else if (rects.length > 0) {
      this.paintDirtyRects(this.layout, rects);
    }
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

    // 1. backgrounds
    for (let row = rowStart; row <= rowEnd; row++) {
      const y = q.originY + (this.rows.positionOf(row) - this.rows.positionOf(q.rowStart));
      const h = this.rows.sizeOf(row);
      for (let col = colStart; col <= colEnd; col++) {
        const cell = sheet.getCell(row, col);
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
    for (let row = rowStart; row <= rowEnd + 1; row++) {
      const y = Math.round(q.originY + (this.rows.positionOf(row) - this.rows.positionOf(q.rowStart))) + 0.5;
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
    for (let row = rowStart; row <= rowEnd; row++) {
      const y = q.originY + (this.rows.positionOf(row) - this.rows.positionOf(q.rowStart));
      const h = this.rows.sizeOf(row);
      for (let col = colStart; col <= colEnd; col++) {
        const cell = sheet.getCell(row, col);
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
    value: import("@opensheet/shared").CellValue,
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
      for (let row = rowStart; row <= rowEnd; row++) {
        const y = originY + (this.rows.positionOf(row) - this.rows.positionOf(baseRow));
        const h = this.rows.sizeOf(row);
        const highlighted = row >= selection.startRow && row <= selection.endRow;
        if (highlighted) {
          ctx.fillStyle = theme.headerHighlight;
          ctx.fillRect(0, y, this.headerWidth, h);
        }
        ctx.fillStyle = highlighted ? theme.headerHighlightText : theme.headerText;
        ctx.textAlign = "center";
        ctx.fillText(String(row + 1), this.headerWidth / 2, y + h / 2);
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

    // Row headers: frozen strip + main strip.
    const frozenRows = this.worksheet.frozenRows;
    if (frozenRows > 0) {
      paintRowHeaders(0, frozenRows - 1, this.headerHeight, 0, this.headerHeight, layout.frozenHeight);
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

    // Selection fill + border.
    const selection = this.selection.state;
    const rects = rangeToCanvasRects(selection.range, layout, this.rows, this.cols);
    for (const rect of rects) {
      ctx.fillStyle = theme.selectionFill;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.strokeStyle = theme.selectionBorder;
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
    }

    // Active cell border + fill handle.
    const activeRects = rangeToCanvasRects(
      {
        startRow: selection.active.row,
        startCol: selection.active.col,
        endRow: selection.active.row,
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

    // Frozen dividers.
    ctx.strokeStyle = theme.frozenDivider;
    ctx.beginPath();
    if (this.worksheet.frozenColumns > 0) {
      const x = layout.mainX + 0.5;
      ctx.moveTo(x, this.headerHeight);
      ctx.lineTo(x, this.contentHeight());
    }
    if (this.worksheet.frozenRows > 0) {
      const y = layout.mainY + 0.5;
      ctx.moveTo(this.headerWidth, y);
      ctx.lineTo(this.contentWidth(), y);
    }
    ctx.stroke();

    this.paintScrollbars(layout);
    ctx.restore();
  }

  private paintScrollbars(layout: ViewportLayout): void {
    const ctx = this.overlayCtx;
    const theme = this.theme;
    const w = this.contentWidth();
    const h = this.contentHeight();
    const frozenW = layout.frozenWidth;
    const frozenH = layout.frozenHeight;

    // Vertical scrollbar.
    const trackX = w - SCROLLBAR;
    const trackY = this.headerHeight;
    const trackH = h - this.headerHeight - SCROLLBAR;
    const totalH = this.rows.totalSize + frozenH;
    const viewH = h - this.headerHeight - frozenH;
    if (trackH > 0 && totalH > viewH) {
      const thumbH = Math.max(24, (viewH / totalH) * trackH);
      const maxScroll = totalH - viewH;
      const thumbY = trackY + (this.scrollY / maxScroll) * (trackH - thumbH);
      ctx.fillStyle = theme.scrollbarTrack;
      ctx.fillRect(trackX, trackY, SCROLLBAR, trackH);
      ctx.fillStyle = theme.scrollbarThumb;
      ctx.fillRect(trackX + 2, thumbY, SCROLLBAR - 4, thumbH);
    }

    // Horizontal scrollbar.
    const trackHX = this.headerWidth;
    const trackHW = w - this.headerWidth - SCROLLBAR;
    const totalW = this.cols.totalSize + frozenW;
    const viewW = w - this.headerWidth - frozenW;
    if (trackHW > 0 && totalW > viewW) {
      const thumbW = Math.max(24, (viewW / totalW) * trackHW);
      const maxScroll = totalW - viewW;
      const thumbX = trackHX + (this.scrollX / maxScroll) * (trackHW - thumbW);
      ctx.fillStyle = theme.scrollbarTrack;
      ctx.fillRect(trackHX, h - SCROLLBAR, trackHW, SCROLLBAR);
      ctx.fillStyle = theme.scrollbarThumb;
      ctx.fillRect(thumbX, h - SCROLLBAR + 2, thumbW, SCROLLBAR - 4);
    }
  }

  // --- input ----------------------------------------------------------------

  private attachInputHandlers(): void {
    this.container.addEventListener("wheel", (e) => this.onWheel(e), { passive: false });
    this.overlayCanvas.addEventListener("mousedown", (e) => this.onMouseDown(e));
    globalThis.addEventListener("mousemove", (e) => this.onMouseMove(e));
    globalThis.addEventListener("mouseup", () => {
      this.drag = null;
    });
    this.container.addEventListener("keydown", (e) => this.onKeyDown(e));
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
    const w = this.contentWidth();
    const h = this.contentHeight();
    if (x >= w - SCROLLBAR && y >= this.headerHeight && y <= h - SCROLLBAR) {
      this.drag = { kind: "v-scroll", grabOffset: y - this.vThumbTop() };
      return;
    }
    if (y >= h - SCROLLBAR && x >= this.headerWidth && x <= w - SCROLLBAR) {
      this.drag = { kind: "h-scroll", grabOffset: x - this.hThumbLeft() };
      return;
    }

    const cell = this.hitTestCell(x, y);
    if (cell === null) return;
    if (e.shiftKey) {
      this.selection.extendTo(cell);
    } else {
      this.selection.setActive(cell);
    }
    this.drag = { kind: "select" };
    this.notifySelection();
    this.scheduleOverlay();
  }

  private onMouseMove(e: MouseEvent): void {
    if (this.drag === null) return;
    const { x, y } = this.canvasPoint(e);
    if (this.drag.kind === "v-scroll") {
      const thumbTop = y - this.drag.grabOffset;
      const trackH = this.contentHeight() - this.headerHeight - SCROLLBAR;
      const thumbH = this.vThumbHeight();
      const maxScroll = this.maxScrollY();
      const ratio = (thumbTop - this.headerHeight) / Math.max(1, trackH - thumbH);
      this.setScroll(this.scrollX, ratio * maxScroll);
      return;
    }
    if (this.drag.kind === "h-scroll") {
      const thumbLeft = x - this.drag.grabOffset;
      const trackW = this.contentWidth() - this.headerWidth - SCROLLBAR;
      const thumbW = this.hThumbWidth();
      const ratio = (thumbLeft - this.headerWidth) / Math.max(1, trackW - thumbW);
      this.setScroll(ratio * this.maxScrollX(), this.scrollY);
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
    const cell = this.hitTestCell(x, y);
    if (cell !== null) {
      this.selection.extendTo(cell);
      this.notifySelection();
      this.scheduleOverlay();
    }
  }

  private onKeyDown(e: KeyboardEvent): void {
    const shift = e.shiftKey;
    const meta = e.metaKey || e.ctrlKey;
    const pageRows = Math.max(
      1,
      this.rows.indexAt(this.scrollY + (this.layout?.mainHeight ?? 400)) - this.rows.indexAt(this.scrollY) - 1,
    );
    let handled = true;
    switch (e.key) {
      case "ArrowUp":
        this.selection.moveBy(-1, 0, shift);
        break;
      case "ArrowDown":
        this.selection.moveBy(1, 0, shift);
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
        this.selection.moveBy(shift ? -1 : 1, 0, false);
        break;
      case "PageUp":
        this.selection.moveBy(-pageRows, 0, shift);
        break;
      case "PageDown":
        this.selection.moveBy(pageRows, 0, shift);
        break;
      case "Home":
        if (meta) this.selection.jumpTo({ row: 0, col: 0 }, shift);
        else this.selection.jumpTo({ row: this.selection.state.active.row, col: 0 }, shift);
        break;
      case "End":
        if (meta) this.selection.jumpTo(this.selection.lastCell(), shift);
        else
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

  private scrollActiveIntoView(): void {
    const active = this.selection.state.active;
    const next = computeScrollToCell(active, { scrollX: this.scrollX, scrollY: this.scrollY }, {
      viewportWidth: this.contentWidth() - this.headerWidth,
      viewportHeight: this.contentHeight() - this.headerHeight,
      rows: this.rows,
      cols: this.cols,
      frozenRowCount: this.worksheet.frozenRows,
      frozenColCount: this.worksheet.frozenColumns,
    });
    this.setScroll(next.scrollX, next.scrollY);
  }

  private hitTestCell(x: number, y: number): { row: number; col: number } | null {
    if (x < this.headerWidth || y < this.headerHeight) return null;
    const frozenCols = this.worksheet.frozenColumns;
    const frozenRows = this.worksheet.frozenRows;
    const frozenW = this.cols.positionOf(frozenCols);
    const frozenH = this.rows.positionOf(frozenRows);

    const inFrozenCol = x < this.headerWidth + frozenW;
    const inFrozenRow = y < this.headerHeight + frozenH;
    const col = inFrozenCol
      ? this.cols.indexAt(x - this.headerWidth)
      : this.cols.indexAt(this.scrollX + (x - this.headerWidth - frozenW));
    const row = inFrozenRow
      ? this.rows.indexAt(y - this.headerHeight)
      : this.rows.indexAt(this.scrollY + (y - this.headerHeight - frozenH));
    if (row >= this.worksheet.rowCount || col >= this.worksheet.columnCount) return null;
    return { row, col };
  }

  private scheduleOverlay(): void {
    this.scheduleFrame();
  }

  private notifySelection(): void {
    const active = this.selection.state.active;
    this.onSelectionChange?.({ activeRow: active.row, activeCol: active.col });
  }

  // --- scrollbar geometry ---------------------------------------------------

  private maxScrollY(): number {
    const frozenH = this.rows.positionOf(this.worksheet.frozenRows);
    return Math.max(0, this.rows.totalSize + frozenH - (this.contentHeight() - this.headerHeight - frozenH));
  }

  private maxScrollX(): number {
    const frozenW = this.cols.positionOf(this.worksheet.frozenColumns);
    return Math.max(0, this.cols.totalSize + frozenW - (this.contentWidth() - this.headerWidth - frozenW));
  }

  private vThumbHeight(): number {
    const trackH = this.contentHeight() - this.headerHeight - SCROLLBAR;
    const totalH = this.rows.totalSize + this.rows.positionOf(this.worksheet.frozenRows);
    const viewH = this.contentHeight() - this.headerHeight - this.rows.positionOf(this.worksheet.frozenRows);
    return Math.max(24, (viewH / Math.max(viewH + 1, totalH)) * trackH);
  }

  private hThumbWidth(): number {
    const trackW = this.contentWidth() - this.headerWidth - SCROLLBAR;
    const totalW = this.cols.totalSize + this.cols.positionOf(this.worksheet.frozenColumns);
    const viewW = this.contentWidth() - this.headerWidth - this.cols.positionOf(this.worksheet.frozenColumns);
    return Math.max(24, (viewW / Math.max(viewW + 1, totalW)) * trackW);
  }

  private vThumbTop(): number {
    const trackH = this.contentHeight() - this.headerHeight - SCROLLBAR;
    return this.headerHeight + (this.scrollY / Math.max(1, this.maxScrollY())) * (trackH - this.vThumbHeight());
  }

  private hThumbLeft(): number {
    const trackW = this.contentWidth() - this.headerWidth - SCROLLBAR;
    return this.headerWidth + (this.scrollX / Math.max(1, this.maxScrollX())) * (trackW - this.hThumbWidth());
  }
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
