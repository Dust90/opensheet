// Row/column structure commands with COMPLETE inverse journals.
//
// M2 semantics — a delete journal must be able to restore:
//   1. deleted cells (value + formula + styleId + numberFormat)
//   2. deleted row heights / column widths
//   3. style references (captured as part of CellData.styleId)
//   4. formula source (captured as part of CellData.formula)
//   5. the freeze state affected by the deletion
//   6. selection adjustments (UI layer clamps on the structure event)
//
// Insert journals are trivial inverses: inserting empty rows/cols is undone
// by deleting exactly those rows/cols (cells were shifted, not copied).

import { MAX_COLS, MAX_ROWS, SheetError, type CellData, type FilterSpec } from "@opensheet/shared";
import { rewriteFormulaReferences, type StructureChange } from "@opensheet/formula-engine";
import type { Worksheet } from "@opensheet/core";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

export interface StructurePayload {
  at: number;
  count?: number;
}

type Axis = "rows" | "columns";

function makeStructureCommand(axis: Axis): SheetCommand<StructurePayload> {
  const kind = axis; // "rows" | "columns"
  return {
    id: axis === "rows" ? "row.insert" : "column.insert",
    validate(payload) {
      if (!Number.isInteger(payload.at) || payload.at < 0) {
        throw new SheetError("E_VALIDATION", `${this.id} requires a non-negative integer "at"`);
      }
      const count = payload.count ?? 1;
      if (!Number.isInteger(count) || count < 1) {
        throw new SheetError("E_VALIDATION", `${this.id} count must be a positive integer`);
      }
    },
    execute(ctx, payload): CommandOutcome {
      const sheet = ctx.workbook.getSheet(ctx.sheetId);
      const previousFilter = cloneFilter(sheet.filter);
      const filterRange = previousFilter === null ? undefined : { ...previousFilter.range };
      const filterBytes = previousFilter === null ? 0 : JSON.stringify(previousFilter).length;
      const count = payload.count ?? 1;
      if (payload.at > (axis === "rows" ? sheet.rowCount : sheet.columnCount)) {
        throw new SheetError(
          "E_VALIDATION",
          `${this.id} position ${payload.at} exceeds sheet ${axis} (${axis === "rows" ? sheet.rowCount : sheet.columnCount})`,
        );
      }
      const total = axis === "rows" ? sheet.rowCount : sheet.columnCount;
      const max = axis === "rows" ? MAX_ROWS : MAX_COLS;
      if (total + count > max) {
        throw new SheetError(
          "E_VALIDATION",
          `${this.id} would grow the sheet past ${axis === "rows" ? "MAX_ROWS" : "MAX_COLS"} (${max})`,
        );
      }
      const apply = () => {
        if (axis === "rows") sheet.insertRows(payload.at, count);
        else sheet.insertColumns(payload.at, count);
      };
      const undoSingle = () => {
        // Inserted empty rows/cols are removed; shifted content returns.
        if (axis === "rows") sheet.deleteRows(payload.at, count);
        else sheet.deleteColumns(payload.at, count);
      };
      // M3: formula references shift with the structure change (relative
      // refs follow; absolute refs stay put). Record original↔rewritten so
      // undo can restore the exact source text.
      const change: StructureChange =
        axis === "rows"
          ? { type: "insertRows", at: payload.at, count }
          : { type: "insertColumns", at: payload.at, count };
      const rewrites: Array<{ row: number; col: number; original: string }> = [];
      const rewriteFormulas = () => {
        for (const [row, col, data] of [...sheet.cellEntries()]) {
          if (data.formula === undefined) continue;
          const rewritten = rewriteFormulaReferences(data.formula, change);
          if (rewritten === data.formula) continue;
          rewrites.push({ row, col, original: data.formula });
          sheet.setCell(row, col, { ...data, formula: rewritten });
        }
      };
      const undoRewrites = () => {
        for (const r of rewrites) {
          const cell = sheet.getCell(r.row, r.col);
          if (cell !== undefined) sheet.setCell(r.row, r.col, { ...cell, formula: r.original });
        }
      };
      const emit = (workbook: typeof ctx.workbook, source: typeof ctx.source) => {
        workbook.emit({
          workbookId: workbook.id,
          sheetId: sheet.id,
          changes: filterRange === undefined
            ? [{ range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 }, kind }]
            : [
                { range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 }, kind },
                { range: filterRange, kind: "filter" },
              ],
          source,
          batch: false,
        });
      };
      // M4.2-D: any structural coordinate change invalidates a filter. Clear
      // before resizing so a formerly valid range can never become invalid.
      if (previousFilter !== null) sheet.setFilter(null);
      apply();
      rewriteFormulas();
      // Freeze stays numerically identical on insert (the frozen pane grows
      // to include the inserted empty rows/cols when inserting above it), so
      // undo does not touch freeze either.
      emit(ctx.workbook, ctx.source);
      const journal: JournalEntry = {
        label: this.id,
        affected: [
          { sheetId: sheet.id, range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 }, kind },
          ...(filterRange === undefined ? [] : [{ sheetId: sheet.id, range: filterRange, kind: "filter" as const }]),
        ],
        approxBytes: 256 + rewrites.length * 96 + filterBytes,
        undo: (rctx) => {
          // Cells are back at their original positions after undoSingle;
          // restore the original formula sources first.
          undoRewrites();
          undoSingle();
          // Restore only after the original dimensions exist again.
          if (previousFilter !== null) sheet.setFilter(previousFilter);
          emit(rctx.workbook, rctx.source);
        },
        redo: (rctx) => {
          if (previousFilter !== null) sheet.setFilter(null);
          apply();
          rewriteFormulas();
          emit(rctx.workbook, rctx.source);
        },
      };
      return { result: undefined, journal };
    },
  };
}

function captureDeletedCells(sheet: Worksheet, at: number, count: number, axis: Axis): Map<string, CellData> {
  const captured = new Map<string, CellData>();
  for (const [row, col, data] of sheet.cellEntries()) {
    const hit = axis === "rows" ? row >= at && row < at + count : col >= at && col < at + count;
    if (hit) captured.set(`${row}:${col}`, { ...data });
  }
  return captured;
}

function makeDeleteCommand(axis: Axis): SheetCommand<StructurePayload> {
  const kind = axis;
  return {
    id: axis === "rows" ? "row.delete" : "column.delete",
    validate(payload) {
      if (!Number.isInteger(payload.at) || payload.at < 0) {
        throw new SheetError("E_VALIDATION", `${this.id} requires a non-negative integer "at"`);
      }
      const count = payload.count ?? 1;
      if (!Number.isInteger(count) || count < 1) {
        throw new SheetError("E_VALIDATION", `${this.id} count must be a positive integer`);
      }
    },
    execute(ctx, payload): CommandOutcome {
      const sheet = ctx.workbook.getSheet(ctx.sheetId);
      const previousFilter = cloneFilter(sheet.filter);
      const filterRange = previousFilter === null ? undefined : { ...previousFilter.range };
      const filterBytes = previousFilter === null ? 0 : JSON.stringify(previousFilter).length;
      const count = payload.count ?? 1;
      const total = axis === "rows" ? sheet.rowCount : sheet.columnCount;
      if (payload.at + count > total) {
        throw new SheetError(
          "E_VALIDATION",
          `${this.id} range [${payload.at}, ${payload.at + count}) exceeds sheet ${axis} (${total})`,
        );
      }
      if (total - count < 1) {
        throw new SheetError(
          "E_VALIDATION",
          `A worksheet must retain at least one ${axis === "rows" ? "row" : "column"}`,
        );
      }
      const prevFreeze = axis === "rows" ? sheet.frozenRows : sheet.frozenColumns;
      // INVERSE DATA (complete): deleted cells, sizes, freeze adjustment.
      const deletedCells = captureDeletedCells(sheet, payload.at, count, axis);
      const deletedSizes = new Map<number, number>();
      if (axis === "rows") {
        for (let i = payload.at; i < payload.at + count; i++) {
          const h = sheet.rowHeights.get(i);
          if (h !== undefined) deletedSizes.set(i, h);
        }
      } else {
        for (let i = payload.at; i < payload.at + count; i++) {
          const w = sheet.columnWidths.get(i);
          if (w !== undefined) deletedSizes.set(i, w);
        }
      }
      const apply = () => {
        if (axis === "rows") sheet.deleteRows(payload.at, count);
        else sheet.deleteColumns(payload.at, count);
        // Freeze: rows/cols deleted inside the frozen pane shrink it; then
        // clamp to the (possibly smaller) sheet.
        const removed = payload.at < prevFreeze ? Math.min(count, prevFreeze - payload.at) : 0;
        const nextFreeze = Math.max(0, prevFreeze - removed);
        if (axis === "rows") sheet.frozenRows = Math.min(nextFreeze, sheet.rowCount);
        else sheet.frozenColumns = Math.min(nextFreeze, sheet.columnCount);
      };
      const restoreInverse = () => {
        if (axis === "rows") {
          sheet.insertRows(payload.at, count);
          for (const [key, data] of deletedCells) {
            const [r, c] = key.split(":").map(Number);
            sheet.setCell(r!, c!, { ...data });
          }
          for (const [index, size] of deletedSizes) sheet.rowHeights.set(index, size);
          sheet.frozenRows = prevFreeze;
        } else {
          sheet.insertColumns(payload.at, count);
          for (const [key, data] of deletedCells) {
            const [r, c] = key.split(":").map(Number);
            sheet.setCell(r!, c!, { ...data });
          }
          for (const [index, size] of deletedSizes) sheet.columnWidths.set(index, size);
          sheet.frozenColumns = prevFreeze;
        }
      };
      // M3: rewrite formula references for the deletion (refs into deleted
      // rows/cols become #REF!); undo restores the original source text.
      const change: StructureChange =
        axis === "rows"
          ? { type: "deleteRows", at: payload.at, count }
          : { type: "deleteColumns", at: payload.at, count };
      const rewrites: Array<{ row: number; col: number; original: string }> = [];
      const rewriteFormulas = () => {
        for (const [row, col, data] of [...sheet.cellEntries()]) {
          if (data.formula === undefined) continue;
          const rewritten = rewriteFormulaReferences(data.formula, change);
          if (rewritten === data.formula) continue;
          rewrites.push({ row, col, original: data.formula });
          sheet.setCell(row, col, { ...data, formula: rewritten });
        }
      };
      const undoRewrites = () => {
        for (const r of rewrites) {
          const cell = sheet.getCell(r.row, r.col);
          if (cell !== undefined) sheet.setCell(r.row, r.col, { ...cell, formula: r.original });
        }
      };
      const emit = (workbook: typeof ctx.workbook, source: typeof ctx.source) => {
        workbook.emit({
          workbookId: workbook.id,
          sheetId: sheet.id,
          changes: filterRange === undefined
            ? [{ range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 }, kind }]
            : [
                { range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 }, kind },
                { range: filterRange, kind: "filter" },
              ],
          source,
          batch: false,
        });
      };
      // Clear before a shrinking mutation so setFilter bounds stay valid.
      if (previousFilter !== null) sheet.setFilter(null);
      apply();
      rewriteFormulas();
      emit(ctx.workbook, ctx.source);
      const journal: JournalEntry = {
        label: this.id,
        affected: [
          { sheetId: sheet.id, range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 }, kind },
          ...(filterRange === undefined ? [] : [{ sheetId: sheet.id, range: filterRange, kind: "filter" as const }]),
        ],
        approxBytes: 512 + deletedCells.size * 160 + deletedSizes.size * 32 + rewrites.length * 96 + filterBytes,
        undo: (rctx) => {
          undoRewrites();
          restoreInverse();
          // Deleted dimensions must exist before this validated FilterSpec can
          // be restored (it may extend past the shrunken sheet).
          if (previousFilter !== null) sheet.setFilter(previousFilter);
          emit(rctx.workbook, rctx.source);
        },
        redo: (rctx) => {
          if (previousFilter !== null) sheet.setFilter(null);
          apply();
          rewriteFormulas();
          emit(rctx.workbook, rctx.source);
        },
      };
      return { result: undefined, journal };
    },
  };
}

export const rowInsertCommand = makeStructureCommand("rows");
export const columnInsertCommand = makeStructureCommand("columns");
export const rowDeleteCommand = makeDeleteCommand("rows");
export const columnDeleteCommand = makeDeleteCommand("columns");

function cloneFilter(filter: Readonly<FilterSpec> | null): FilterSpec | null {
  return filter === null
    ? null
    : {
        range: { ...filter.range },
        hasHeader: filter.hasHeader,
        conditions: filter.conditions.map((condition) => ({ ...condition })),
      };
}
