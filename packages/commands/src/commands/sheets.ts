// Sheet lifecycle commands (M0: sheet.create only).

import { Worksheet } from "@opensheet/core";
import { SheetError } from "@opensheet/shared";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

export interface SheetCreatePayload {
  name: string;
  rows?: number;
  columns?: number;
}

export interface SheetCreateResult {
  sheetId: string;
  name: string;
  rowCount: number;
  columnCount: number;
}

const DEFAULT_ROWS = 1000;
const DEFAULT_COLUMNS = 26;

export const sheetCreateCommand: SheetCommand<SheetCreatePayload, SheetCreateResult> = {
  id: "sheet.create",
  validate(payload) {
    if (typeof payload.name !== "string" || payload.name.trim().length === 0) {
      throw new SheetError("E_VALIDATION", "sheet.create requires a non-empty name");
    }
    if ((payload.rows !== undefined && payload.rows < 1) || (payload.columns !== undefined && payload.columns < 1)) {
      throw new SheetError("E_VALIDATION", "sheet.create rows/columns must be >= 1");
    }
  },
  execute(ctx, payload): CommandOutcome<SheetCreateResult> {
    if (ctx.workbook.listSheets().some((s) => s.name === payload.name)) {
      throw new SheetError("E_VALIDATION", `Sheet name already exists: "${payload.name}"`);
    }
    const sheet = new Worksheet({
      id: crypto.randomUUID(),
      name: payload.name,
      rowCount: payload.rows ?? DEFAULT_ROWS,
      columnCount: payload.columns ?? DEFAULT_COLUMNS,
    });
    const index = ctx.workbook.listSheets().length;
    const apply = () => ctx.workbook.restoreSheet(sheet, index);
    apply();
    ctx.workbook.setActiveSheet(sheet.id);
    ctx.workbook.emit({
      workbookId: ctx.workbook.id,
      sheetId: sheet.id,
      changes: [
        {
          range: { startRow: 0, startCol: 0, endRow: sheet.rowCount - 1, endCol: sheet.columnCount - 1 },
          kind: "structure",
        },
      ],
      source: ctx.source,
      batch: false,
    });
    const journal: JournalEntry = {
      label: "sheet.create",
      affected: [
        {
          sheetId: sheet.id,
          range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
          kind: "structure",
        },
      ],
      approxBytes: 512,
      undo: (rctx) => {
        rctx.workbook.removeSheet(sheet.id);
        rctx.workbook.emit({
          workbookId: rctx.workbook.id,
          sheetId: sheet.id,
          changes: [{ range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, kind: "structure" }],
          source: rctx.source,
          batch: false,
        });
      },
      redo: (rctx) => {
        rctx.workbook.restoreSheet(sheet, index);
        rctx.workbook.emit({
          workbookId: rctx.workbook.id,
          sheetId: sheet.id,
          changes: [{ range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, kind: "structure" }],
          source: rctx.source,
          batch: false,
        });
      },
    };
    return {
      result: {
        sheetId: sheet.id,
        name: sheet.name,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
      },
      journal,
    };
  },
};
