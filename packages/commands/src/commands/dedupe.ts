import { SheetError, type CellData, type DedupeSpec, type Range, validateDedupeSpec } from "@injoysai/opensheet-shared";
import { translateFormulaReferences } from "@injoysai/opensheet-formula-engine";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";
import { conflictsWithFilter } from "../data-operation-conflicts.js";
import { buildDedupePlan, type DedupePlan } from "../dedupe-plan.js";

export interface RangeDedupePayload { spec: DedupeSpec; }

interface FormulaMove { row: number; col: number; translated: string; }
interface SavedCell { rowOffset: number; colOffset: number; data: CellData; }

export const rangeDedupeCommand: SheetCommand<RangeDedupePayload> = {
  id: "range.dedupe",
  requiresFreshDerivedState: true,
  validate(payload, ctx) {
    if (typeof payload !== "object" || payload === null || !("spec" in payload)) {
      throw new SheetError("E_VALIDATION", "range.dedupe requires a DedupeSpec");
    }
    validateDedupeSpec(payload.spec);
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    if (payload.spec.range.endRow >= sheet.rowCount || payload.spec.range.endCol >= sheet.columnCount) {
      throw new SheetError("E_INVALID_RANGE", "DedupeSpec range exceeds worksheet bounds");
    }
    if (conflictsWithFilter(payload.spec.range, sheet.filter)) {
      throw new SheetError("E_VALIDATION", "Dedupe range intersects active filter rows");
    }
  },
  execute(ctx, payload): CommandOutcome {
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    const spec = cloneDedupeSpec(payload.spec);
    const plan = buildDedupePlan(sheet, spec);
    if (plan.removedRows === 0) return { result: undefined, journal: null };

    const firstChangedOffset = plan.removedSourceOffsets[0]!;
    // Parsing/translation can throw; finish it before changing any CellStore data.
    const formulas = collectTranslations(sheet, spec.range, plan);
    const before = captureSuffix(sheet, spec.range, plan.bodyStartRow, firstChangedOffset);
    applyCompaction(sheet, spec.range, plan, formulas);
    emit(ctx.workbook, sheet.id, spec.range, ctx.source);
    return { result: undefined, journal: journal(sheet.id, spec.range, plan, firstChangedOffset, before, formulas) };
  },
};

function cloneDedupeSpec(spec: DedupeSpec): DedupeSpec {
  return { range: { ...spec.range }, hasHeader: spec.hasHeader, keyColumnOffsets: [...spec.keyColumnOffsets], keep: "first" };
}

function collectTranslations(sheet: import("@injoysai/opensheet-core").Worksheet, range: Range, plan: DedupePlan): FormulaMove[] {
  const moves: FormulaMove[] = [];
  for (let destinationOffset = 0; destinationOffset < plan.keptRowCount; destinationOffset += 1) {
    const sourceOffset = plan.keptSourceOffsets[destinationOffset]!;
    if (sourceOffset === destinationOffset) continue;
    const sourceRow = plan.bodyStartRow + sourceOffset;
    const destinationRow = plan.bodyStartRow + destinationOffset;
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const formula = sheet.getCell(sourceRow, col)?.formula;
      if (formula !== undefined) moves.push({
        row: destinationRow,
        col,
        translated: translateFormulaReferences(formula, destinationRow - sourceRow, 0, {
          rowCount: sheet.rowCount,
          columnCount: sheet.columnCount,
        }),
      });
    }
  }
  return moves;
}

function captureSuffix(
  sheet: import("@injoysai/opensheet-core").Worksheet,
  range: Range,
  bodyStartRow: number,
  firstChangedOffset: number,
): SavedCell[] {
  const firstRow = bodyStartRow + firstChangedOffset;
  const cells: SavedCell[] = [];
  for (const [row, col, data] of sheet.cellEntries()) {
    if (row >= firstRow && row <= range.endRow && col >= range.startCol && col <= range.endCol) {
      cells.push({ rowOffset: row - bodyStartRow, colOffset: col - range.startCol, data: { ...data } });
    }
  }
  return cells;
}

function clearSuffix(sheet: import("@injoysai/opensheet-core").Worksheet, range: Range, bodyStartRow: number, firstChangedOffset: number): void {
  const firstRow = bodyStartRow + firstChangedOffset;
  for (let row = firstRow; row <= range.endRow; row += 1) {
    for (let col = range.startCol; col <= range.endCol; col += 1) sheet.deleteCell(row, col);
  }
}

function applyCompaction(
  sheet: import("@injoysai/opensheet-core").Worksheet,
  range: Range,
  plan: DedupePlan,
  formulas: readonly FormulaMove[],
): void {
  for (let col = range.startCol; col <= range.endCol; col += 1) {
    const values = new Array<CellData | undefined>(plan.bodyRowCount);
    for (let sourceOffset = 0; sourceOffset < plan.bodyRowCount; sourceOffset += 1) {
      values[sourceOffset] = sheet.getCell(plan.bodyStartRow + sourceOffset, col);
    }
    for (let destinationOffset = 0; destinationOffset < plan.bodyRowCount; destinationOffset += 1) {
      const data = destinationOffset < plan.keptRowCount
        ? values[plan.keptSourceOffsets[destinationOffset]!]
        : undefined;
      const row = plan.bodyStartRow + destinationOffset;
      if (data === undefined) sheet.deleteCell(row, col);
      else sheet.setCell(row, col, { ...data });
    }
  }
  for (const move of formulas) {
    const data = sheet.getCell(move.row, move.col);
    if (data !== undefined) sheet.setCell(move.row, move.col, { ...data, formula: move.translated });
  }
}

function restoreSuffix(
  sheet: import("@injoysai/opensheet-core").Worksheet,
  range: Range,
  bodyStartRow: number,
  firstChangedOffset: number,
  before: readonly SavedCell[],
): void {
  clearSuffix(sheet, range, bodyStartRow, firstChangedOffset);
  for (const cell of before) {
    sheet.setCell(bodyStartRow + cell.rowOffset, range.startCol + cell.colOffset, { ...cell.data });
  }
}

function emit(workbook: import("@injoysai/opensheet-core").Workbook, sheetId: string, range: Range, source: import("@injoysai/opensheet-shared").ChangeSource): void {
  workbook.emit({ workbookId: workbook.id, sheetId, changes: [{ range, kind: "reorder" }], source, batch: false });
}

function journal(
  sheetId: string,
  range: Range,
  plan: DedupePlan,
  firstChangedOffset: number,
  before: SavedCell[],
  formulas: FormulaMove[],
): JournalEntry {
  const approxBytes = 256 + before.length * 160 + formulas.reduce((total, move) => total + move.translated.length, 0);
  return {
    label: "range.dedupe",
    affected: [{ sheetId, range, kind: "reorder" }],
    approxBytes,
    undo: (ctx) => {
      const sheet = ctx.workbook.getSheet(sheetId);
      restoreSuffix(sheet, range, plan.bodyStartRow, firstChangedOffset, before);
      emit(ctx.workbook, sheetId, range, ctx.source);
    },
    redo: (ctx) => {
      const sheet = ctx.workbook.getSheet(sheetId);
      applyCompaction(sheet, range, plan, formulas);
      emit(ctx.workbook, sheetId, range, ctx.source);
    },
  };
}
