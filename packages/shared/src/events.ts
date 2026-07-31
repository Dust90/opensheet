// @opensheet/shared — change event contracts

import type { Range } from "./range.js";

/**
 * Who caused a change.
 * - "user": interactive editing
 * - "api": external SDK / applyOperations
 * - "undo" / "redo": history replay
 * - "derived": formula recalculation output. Never enters Undo history.
 */
export type ChangeSource = "user" | "api" | "undo" | "redo" | "derived";

export type ChangeKind =
  | "cells" // cell values/formulas changed inside the range
  | "style" // only presentation changed
  | "rows" // row structure (insert/delete/resize)
  | "columns" // column structure
  | "structure" // sheet-level structure (rename, freeze, sheet add/remove)
  | "metadata"; // row heights / column widths

export interface CellChange {
  range: Range;
  kind: ChangeKind;
}

/**
 * A single atomic notification. During a transaction all intermediate
 * emissions are buffered; on commit, observers receive at most ONE merged
 * event per (workbook + sheet + source) — e.g. a user write plus a derived
 * recalculation yields two events, one per source. Observers never see
 * intermediate transaction state.
 */
export interface ChangeEvent {
  workbookId: string;
  sheetId: string;
  changes: CellChange[];
  source: ChangeSource;
  /** True when this event is the merged result of a multi-command transaction. */
  batch: boolean;
}

export type ChangeListener = (event: ChangeEvent) => void;

export type Unsubscribe = () => void;
