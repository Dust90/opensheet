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

import { MAX_COLS, MAX_ROWS, SheetError, type CellData } from "@opensheet/shared";
import type { Worksheet } from "@opensheet/core";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

export interface StructurePayload {
  at: number;
  count?: number;
}

type Axis = "rows" | "columns";

const FULL_RANGE = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };

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
      const emit = (workbook: typeof ctx.workbook, source: typeof ctx.source) => {
        workbook.emit({
          workbookId: workbook.id,
          sheetId: sheet.id,
          changes: [
            {
              range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 },
              kind,
            },
          ],
          source,
          batch: false,
        });
      };
      apply();
      // Freeze stays numerically identical on insert (the frozen pane grows
      // to include the inserted empty rows/cols when inserting above it), so
      // undo does not touch freeze either.
      emit(ctx.workbook, ctx.source);
      const journal: JournalEntry = {
        label: this.id,
        affected: [{ sheetId: sheet.id, range: FULL_RANGE, kind }],
        approxBytes: 256,
        undo: (rctx) => {
          undoSingle();
          emit(rctx.workbook, rctx.source);
        },
        redo: (rctx) => {
          apply();
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
      const emit = (workbook: typeof ctx.workbook, source: typeof ctx.source) => {
        workbook.emit({
          workbookId: workbook.id,
          sheetId: sheet.id,
          changes: [
            {
              range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 },
              kind,
            },
          ],
          source,
          batch: false,
        });
      };
      apply();
      emit(ctx.workbook, ctx.source);
      const journal: JournalEntry = {
        label: this.id,
        affected: [{ sheetId: sheet.id, range: FULL_RANGE, kind }],
        approxBytes: 512 + deletedCells.size * 160 + deletedSizes.size * 32,
        undo: (rctx) => {
          restoreInverse();
          emit(rctx.workbook, rctx.source);
        },
        redo: (rctx) => {
          apply();
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
