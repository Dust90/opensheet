// @opensheet/renderer-canvas — dual-canvas grid renderer (Content + Overlay).

export { SheetGrid } from "./grid.js";
export type { SheetGridOptions } from "./grid.js";
export { AxisMetrics } from "./axis-metrics.js";
export type { SizeAccessor } from "./axis-metrics.js";
export {
  clampScroll,
  computeScrollToCell,
  computeViewport,
} from "./viewport.js";
export type { Quadrant, ScrollPosition, ViewportInput, ViewportLayout } from "./viewport.js";
export { SelectionModel } from "./selection.js";
export type { SelectionState } from "./selection.js";
export {
  FilteredRowProjection,
  IdentityRowProjection,
  lastVisiblePhysicalRow,
  physicalRangeToVisualRange,
  relocateToVisibleRow,
} from "./row-projection.js";
export type { RowProjection } from "./row-projection.js";
export {
  DirtyRegionTracker,
  mergeRects,
  rangeToCanvasRects,
} from "./dirty-region.js";
export type { PixelRect } from "./dirty-region.js";
export { darkTheme, lightTheme } from "./theme.js";
export type { GridTheme } from "./theme.js";
export { CellEditor } from "./editor/cell-editor.js";
export type { CellEditorCallbacks, CellEditorRect } from "./editor/cell-editor.js";
export {
  cellDisplayText,
  decideKeyInPhase,
  inferPrimitive,
  isPrintableKey,
} from "./editor/editor-state.js";
export type { EditorAction, EditorKeyInfo, EditorPhase } from "./editor/editor-state.js";
export { cellRectInCanvas, hitTestCell } from "./coordinate-mapper.js";
export type { CellHit, HitZone } from "./coordinate-mapper.js";
