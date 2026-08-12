// formula.set: store a formula source on a single cell (M3).
//
// Semantics:
//   - Single-cell range only (no range fill in phase 1).
//   - The formula must PARSE; a syntax error rejects the whole transaction
//     and leaves the old cell untouched (no invalid formula source stored).
//   - The cell keeps its metadata (styleId/numberFormat); the previous
//     computed value becomes the NEW cached value if the old cell already
//     had a matching formula, otherwise null — the beforeCommit recalculation
//     hook overwrites it within the same transaction.
//   - Journal captures the complete previous AND next CellData for
//     undo/redo, so the dependency graph can be reconciled by the runtime.

import { parseRange, SheetError, type CellData } from "@injoysai/opensheet-shared";
import { parseFormula } from "@injoysai/opensheet-formula-engine";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

export interface FormulaSetPayload {
  range: string;
  formula: string;
}

export const formulaSetCommand: SheetCommand<FormulaSetPayload> = {
  id: "formula.set",
  validate(payload) {
    const range = parseRange(payload.range);
    if (range.startRow !== range.endRow || range.startCol !== range.endCol) {
      throw new SheetError("E_VALIDATION", "formula.set supports a single cell only (phase 1)");
    }
    if (typeof payload.formula !== "string" || !payload.formula.trim().startsWith("=")) {
      throw new SheetError("E_VALIDATION", "formula.set requires a formula string starting with '='");
    }
    // Syntax errors reject the command — nothing is written.
    parseFormula(payload.formula);
  },
  execute(ctx, payload): CommandOutcome {
    const range = parseRange(payload.range);
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    if (range.endRow >= sheet.rowCount || range.endCol >= sheet.columnCount) {
      throw new SheetError(
        "E_INVALID_RANGE",
        `Range exceeds sheet bounds (${sheet.rowCount} rows x ${sheet.columnCount} cols)`,
      );
    }
    const row = range.startRow;
    const col = range.startCol;
    const previous = sheet.getCell(row, col);
    const previousClone = previous === undefined ? undefined : { ...previous };
    // Metadata preserved; value resets to null (recalc fills it this commit).
    const next: CellData = {
      ...(previousClone ?? {}),
      formula: payload.formula,
      value: null,
    };
    const apply = () => sheet.setCell(row, col, { ...next });
    apply();
    ctx.workbook.emit({
      workbookId: ctx.workbook.id,
      sheetId: sheet.id,
      changes: [{ range, kind: "cells" }],
      source: ctx.source,
      batch: false,
    });
    const journal: JournalEntry = {
      label: "formula.set",
      affected: [{ sheetId: sheet.id, range, kind: "cells" }],
      approxBytes: 512 + payload.formula.length,
      undo: (rctx) => {
        if (previousClone === undefined) sheet.deleteCell(row, col);
        else sheet.setCell(row, col, { ...previousClone });
        rctx.workbook.emit({
          workbookId: rctx.workbook.id,
          sheetId: sheet.id,
          changes: [{ range, kind: "cells" }],
          source: rctx.source,
          batch: false,
        });
      },
      redo: (rctx) => {
        apply();
        rctx.workbook.emit({
          workbookId: rctx.workbook.id,
          sheetId: sheet.id,
          changes: [{ range, kind: "cells" }],
          source: rctx.source,
          batch: false,
        });
      },
    };
    return { result: undefined, journal };
  },
};
