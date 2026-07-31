// Cell mutation commands: cell.set / cell.clear / range.write.
//
// WRITE SEMANTICS (M2.8): writing a value NEVER touches presentation
// metadata. styleId / numberFormat are preserved; the old formula is
// superseded by the new literal value. An empty write to a cell without any
// style removes the sparse entry; a cell carrying ONLY style survives as
// { value: null, styleId }.

import type { CellData, CellPrimitive, Range } from "@opensheet/shared";
import { parseRange, SheetError } from "@opensheet/shared";
import type { Worksheet } from "@opensheet/core";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

type CapturedCell = { row: number; col: number; previous: CellData | undefined };

/** Merge a new literal value into the previous cell (metadata-preserving). */
function writeValuePreservingStyle(sheet: Worksheet, row: number, col: number, value: CellPrimitive): void {
  const previous = sheet.getCell(row, col);
  if (previous === undefined) {
    if (value === null) return; // nothing to store
    sheet.setCell(row, col, { value });
    return;
  }
  const next: CellData = { ...previous, value };
  delete next.formula; // literal value supersedes any stored formula
  const hasMetadata = next.styleId !== undefined || next.numberFormat !== undefined;
  if (value === null && !hasMetadata) {
    sheet.deleteCell(row, col); // truly empty: reclaim the sparse slot
    return;
  }
  sheet.setCell(row, col, next);
}

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
          writeValuePreservingStyle(sheet, row, col, payload.value);
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
    // Normalize/clone rule: journal closures must never reference caller-owned
    // payloads — later mutations by the caller would corrupt redo.
    const values = payload.values.map((row) => [...row]);
    const captured = captureCells(sheet, range);
    const apply = () => {
      values.forEach((rowValues, r) => {
        rowValues.forEach((value, c) => {
          writeValuePreservingStyle(sheet, range.startRow + r, range.startCol + c, value);
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
