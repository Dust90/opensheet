// Cell mutation commands: cell.set / cell.clear / range.write.

import type { CellData, CellPrimitive, Range } from "@opensheet/shared";
import { parseRange, SheetError } from "@opensheet/shared";
import type { Worksheet } from "@opensheet/core";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

type CapturedCell = { row: number; col: number; previous: CellData | undefined };

function captureCells(sheet: Worksheet, range: Range): CapturedCell[] {
  const captured: CapturedCell[] = [];
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      const previous = sheet.getCell(row, col);
      captured.push({ row, col, previous: previous === undefined ? undefined : { ...previous } });
    }
  }
  return captured;
}

function restoreCells(sheet: Worksheet, captured: readonly CapturedCell[]): void {
  for (const { row, col, previous } of captured) {
    if (previous === undefined) sheet.deleteCell(row, col);
    else sheet.setCell(row, col, { ...previous });
  }
}

function makeJournal(init: {
  label: string;
  sheetId: string;
  range: Range;
  apply: () => void;
  captured: CapturedCell[];
  sheet: Worksheet;
}): JournalEntry {
  const { label, sheetId, range, apply, captured, sheet } = init;
  return {
    label,
    affected: [{ sheetId, range, kind: "cells" }],
    approxBytes: 256 + captured.length * 160,
    undo: (rctx) => {
      restoreCells(sheet, captured);
      rctx.workbook.emit({
        workbookId: rctx.workbook.id,
        sheetId,
        changes: [{ range, kind: "cells" }],
        source: rctx.source,
        batch: false,
      });
    },
    redo: (rctx) => {
      apply();
      rctx.workbook.emit({
        workbookId: rctx.workbook.id,
        sheetId,
        changes: [{ range, kind: "cells" }],
        source: rctx.source,
        batch: false,
      });
    },
  };
}

function assertRangeInSheet(range: Range, sheet: Worksheet): void {
  if (range.endRow >= sheet.rowCount || range.endCol >= sheet.columnCount) {
    throw new SheetError(
      "E_INVALID_RANGE",
      `Range exceeds sheet bounds (${sheet.rowCount} rows x ${sheet.columnCount} cols)`,
    );
  }
}

export const cellSetCommand: SheetCommand<{ range: string; value: CellPrimitive }> = {
  id: "cell.set",
  validate(payload) {
    parseRange(payload.range); // throws on malformed input
    const t = typeof payload.value;
    if (payload.value !== null && t !== "string" && t !== "number" && t !== "boolean") {
      throw new SheetError("E_VALIDATION", `cell.set value must be a primitive, got ${t}`);
    }
  },
  execute(ctx, payload): CommandOutcome {
    const range = parseRange(payload.range);
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    assertRangeInSheet(range, sheet);
    const captured = captureCells(sheet, range);
    const apply = () => {
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startCol; col <= range.endCol; col++) {
          sheet.setCell(row, col, { value: payload.value });
        }
      }
    };
    apply();
    ctx.workbook.emit({
      workbookId: ctx.workbook.id,
      sheetId: sheet.id,
      changes: [{ range, kind: "cells" }],
      source: ctx.source,
      batch: false,
    });
    return {
      result: undefined,
      journal: makeJournal({ label: "cell.set", sheetId: sheet.id, range, apply, captured, sheet }),
    };
  },
};

export const cellClearCommand: SheetCommand<{ range: string }> = {
  id: "cell.clear",
  validate(payload) {
    parseRange(payload.range);
  },
  execute(ctx, payload): CommandOutcome {
    const range = parseRange(payload.range);
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    assertRangeInSheet(range, sheet);
    const captured = captureCells(sheet, range);
    const apply = () => {
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startCol; col <= range.endCol; col++) {
          sheet.deleteCell(row, col);
        }
      }
    };
    apply();
    ctx.workbook.emit({
      workbookId: ctx.workbook.id,
      sheetId: sheet.id,
      changes: [{ range, kind: "cells" }],
      source: ctx.source,
      batch: false,
    });
    return {
      result: undefined,
      journal: makeJournal({
        label: "cell.clear",
        sheetId: sheet.id,
        range,
        apply,
        captured,
        sheet,
      }),
    };
  },
};

export const rangeWriteCommand: SheetCommand<{ range: string; values: CellPrimitive[][] }> = {
  id: "range.write",
  validate(payload) {
    const range = parseRange(payload.range);
    const rows = payload.values.length;
    const cols = payload.values[0]?.length ?? 0;
    if (rows === 0 || cols === 0) {
      throw new SheetError("E_VALIDATION", "range.write values must be a non-empty matrix");
    }
    if (payload.values.some((r) => r.length !== cols)) {
      throw new SheetError("E_VALIDATION", "range.write values must be rectangular");
    }
    if (rows !== range.endRow - range.startRow + 1 || cols !== range.endCol - range.startCol + 1) {
      throw new SheetError(
        "E_VALIDATION",
        `range.write values (${rows}x${cols}) do not match range dimensions`,
      );
    }
  },
  execute(ctx, payload): CommandOutcome {
    const range = parseRange(payload.range);
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    assertRangeInSheet(range, sheet);
    const captured = captureCells(sheet, range);
    const apply = () => {
      payload.values.forEach((rowValues, r) => {
        rowValues.forEach((value, c) => {
          sheet.setCell(range.startRow + r, range.startCol + c, { value });
        });
      });
    };
    apply();
    ctx.workbook.emit({
      workbookId: ctx.workbook.id,
      sheetId: sheet.id,
      changes: [{ range, kind: "cells" }],
      source: ctx.source,
      batch: false,
    });
    return {
      result: undefined,
      journal: makeJournal({
        label: "range.write",
        sheetId: sheet.id,
        range,
        apply,
        captured,
        sheet,
      }),
    };
  },
};
