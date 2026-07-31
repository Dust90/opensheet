// range.style: apply presentation attributes to a range. Style objects are
// deduplicated by the workbook StyleTable; cells reference them by id. The
// journal captures each cell's PREVIOUS CellData so undo restores values,
// formulas AND style references exactly (a styled-then-undone cell that had
// a value but no style must keep its value).

import type { CellData, CellStyle } from "@opensheet/shared";
import { parseRange, SheetError } from "@opensheet/shared";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";

export interface RangeStylePayload {
  range: string;
  style: Partial<CellStyle>;
}

export const rangeStyleCommand: SheetCommand<RangeStylePayload> = {
  id: "range.style",
  validate(payload) {
    parseRange(payload.range);
    const style = payload.style ?? {};
    const keys = Object.keys(style);
    if (keys.length === 0) {
      throw new SheetError("E_VALIDATION", "range.style requires at least one style attribute");
    }
    if (style.bold !== undefined && typeof style.bold !== "boolean") {
      throw new SheetError("E_VALIDATION", "range.style bold must be a boolean");
    }
    if (style.italic !== undefined && typeof style.italic !== "boolean") {
      throw new SheetError("E_VALIDATION", "range.style italic must be a boolean");
    }
    if (style.fontSize !== undefined && (!Number.isFinite(style.fontSize) || style.fontSize <= 0)) {
      throw new SheetError("E_VALIDATION", "range.style fontSize must be a positive number");
    }
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
    // Merge semantics: keep existing attributes, override only the provided
    // ones. Empty diff (identical style) is a no-op but still journaled.
    const patch = { ...payload.style };
    const apply = () => {
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startCol; col <= range.endCol; col++) {
          const cell = sheet.getCell(row, col);
          const previousId = cell?.styleId;
          const merged: CellStyle = previousId !== undefined ? { ...ctx.workbook.styles.get(previousId) } : {};
          Object.assign(merged, patch);
          const nextId = ctx.workbook.styles.register(merged);
          if (previousId !== nextId) {
            if (cell === undefined) sheet.setCell(row, col, { value: null, styleId: nextId });
            else sheet.setCell(row, col, { ...cell, styleId: nextId });
          }
        }
      }
    };
    // Capture previous cells BEFORE applying (inverse patch). Recording only
    // the styleId is NOT enough: a cell that had a value but no style must
    // survive an undo, while a cell that did not exist must be removed.
    const previous = new Map<string, CellData | undefined>();
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = sheet.getCell(row, col);
        previous.set(`${row}:${col}`, cell === undefined ? undefined : { ...cell });
      }
    }
    apply();
    ctx.workbook.emit({
      workbookId: ctx.workbook.id,
      sheetId: sheet.id,
      changes: [{ range, kind: "style" }],
      source: ctx.source,
      batch: false,
    });
    const journal: JournalEntry = {
      label: "range.style",
      affected: [{ sheetId: sheet.id, range, kind: "style" }],
      approxBytes: 512 + previous.size * 160,
      undo: (rctx) => {
        for (let row = range.startRow; row <= range.endRow; row++) {
          for (let col = range.startCol; col <= range.endCol; col++) {
            const data = previous.get(`${row}:${col}`);
            if (data === undefined) {
              sheet.deleteCell(row, col);
            } else {
              sheet.setCell(row, col, { ...data });
            }
          }
        }
        rctx.workbook.emit({
          workbookId: rctx.workbook.id,
          sheetId: sheet.id,
          changes: [{ range, kind: "style" }],
          source: rctx.source,
          batch: false,
        });
      },
      redo: (rctx) => {
        apply();
        rctx.workbook.emit({
          workbookId: rctx.workbook.id,
          sheetId: sheet.id,
          changes: [{ range, kind: "style" }],
          source: rctx.source,
          batch: false,
        });
      },
    };
    return { result: undefined, journal };
  },
};
