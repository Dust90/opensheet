// sheet.freeze: set frozen row/column counts.

import { SheetError } from "@injoysai/opensheet-shared";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

export interface SheetFreezePayload {
  frozenRows: number;
  frozenColumns: number;
}

export const sheetFreezeCommand: SheetCommand<SheetFreezePayload> = {
  id: "sheet.freeze",
  validate(payload, ctx) {
    if (
      !Number.isInteger(payload.frozenRows) ||
      payload.frozenRows < 0 ||
      !Number.isInteger(payload.frozenColumns) ||
      payload.frozenColumns < 0
    ) {
      throw new SheetError("E_VALIDATION", "sheet.freeze counts must be non-negative integers");
    }
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    if (payload.frozenRows > sheet.rowCount || payload.frozenColumns > sheet.columnCount) {
      throw new SheetError(
        "E_VALIDATION",
        `sheet.freeze exceeds sheet bounds: sheet is ${sheet.rowCount}x${sheet.columnCount}, ` +
          `requested ${payload.frozenRows}x${payload.frozenColumns}`,
      );
    }
  },
  execute(ctx, payload): CommandOutcome {
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    const previous = { frozenRows: sheet.frozenRows, frozenColumns: sheet.frozenColumns };
    const next = { frozenRows: payload.frozenRows, frozenColumns: payload.frozenColumns };
    const range = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
    const emitTo = (workbook: typeof ctx.workbook, source: typeof ctx.source) => {
      workbook.emit({
        workbookId: workbook.id,
        sheetId: sheet.id,
        changes: [{ range, kind: "structure" }],
        source,
        batch: false,
      });
    };
    sheet.frozenRows = next.frozenRows;
    sheet.frozenColumns = next.frozenColumns;
    emitTo(ctx.workbook, ctx.source);
    const journal: JournalEntry = {
      label: "sheet.freeze",
      affected: [{ sheetId: sheet.id, range, kind: "structure" }],
      approxBytes: 128,
      undo: (rctx) => {
        sheet.frozenRows = previous.frozenRows;
        sheet.frozenColumns = previous.frozenColumns;
        emitTo(rctx.workbook, rctx.source);
      },
      redo: (rctx) => {
        sheet.frozenRows = next.frozenRows;
        sheet.frozenColumns = next.frozenColumns;
        emitTo(rctx.workbook, rctx.source);
      },
    };
    return { result: undefined, journal };
  },
};
