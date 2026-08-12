// Sheet lifecycle commands (M0: sheet.create only).

import { Worksheet } from "@injoysai/opensheet-core";
import { SheetError } from "@injoysai/opensheet-shared";
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

/**
 * Internal payload used by import/export adapters.  The Worksheet is built
 * off-workbook first, so an import never exposes partially-written cells.
 */
export interface SheetImportPayload {
  sheet: Worksheet;
}

export interface SheetImportResult {
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
    const previousActiveSheetId = ctx.workbook.activeSheetId;
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
        // Restore the sheet that was active before sheet.create ran.
        if (rctx.workbook.listSheets().some((s) => s.id === previousActiveSheetId)) {
          rctx.workbook.setActiveSheet(previousActiveSheetId);
        }
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
        rctx.workbook.setActiveSheet(sheet.id);
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

/**
 * Attach a fully-populated staging worksheet as one reversible operation.
 * This is deliberately not exposed as a public SheetOperation: callers must
 * construct the sheet away from the live Workbook before executing it.
 */
export const sheetImportCommand: SheetCommand<SheetImportPayload, SheetImportResult> = {
  id: "sheet.import",
  validate(payload) {
    if (typeof payload !== "object" || payload === null || !(payload.sheet instanceof Worksheet)) {
      throw new SheetError("E_VALIDATION", "sheet.import requires a staging Worksheet");
    }
  },
  execute(ctx, payload): CommandOutcome<SheetImportResult> {
    const sheet = payload.sheet;
    if (ctx.workbook.listSheets().some((candidate) => candidate.id === sheet.id || candidate.name === sheet.name)) {
      throw new SheetError("E_VALIDATION", `Imported sheet already exists: "${sheet.name}"`);
    }
    const index = ctx.workbook.listSheets().length;
    const previousActiveSheetId = ctx.workbook.activeSheetId;
    const range = {
      startRow: 0,
      startCol: 0,
      endRow: Math.max(0, sheet.rowCount - 1),
      endCol: Math.max(0, sheet.columnCount - 1),
    };
    const emit = (workbook: typeof ctx.workbook, source: typeof ctx.source) => {
      workbook.emit({
        workbookId: workbook.id,
        sheetId: sheet.id,
        changes: [{ range, kind: "structure" }],
        source,
        batch: false,
      });
    };

    ctx.workbook.restoreSheet(sheet, index);
    ctx.workbook.setActiveSheet(sheet.id);
    emit(ctx.workbook, ctx.source);

    return {
      result: { sheetId: sheet.id, name: sheet.name, rowCount: sheet.rowCount, columnCount: sheet.columnCount },
      journal: {
        label: "sheet.import",
        affected: [{ sheetId: sheet.id, range, kind: "structure" }],
        // The journal retains the imported sparse worksheet for redo.  This
        // estimate keeps the existing History memory policy meaningful.
        approxBytes: 512 + sheet.cellCount * 96,
        undo: (replay) => {
          replay.workbook.removeSheet(sheet.id);
          if (replay.workbook.listSheets().some((candidate) => candidate.id === previousActiveSheetId)) {
            replay.workbook.setActiveSheet(previousActiveSheetId);
          }
          emit(replay.workbook, replay.source);
        },
        redo: (replay) => {
          replay.workbook.restoreSheet(sheet, index);
          replay.workbook.setActiveSheet(sheet.id);
          emit(replay.workbook, replay.source);
        },
      },
    };
  },
};
