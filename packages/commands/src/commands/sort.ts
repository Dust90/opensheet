import { SheetError, type CellData, type Range, type SortSpec, validateSortSpec } from "@opensheet/shared";
import { translateFormulaReferences } from "@opensheet/formula-engine";
import type { CommandOutcome, JournalEntry, SheetCommand } from "../types.js";
import { buildSortPlan, conflictsWithFilter, type SortPlan } from "../sort-plan.js";

export interface RangeSortPayload { spec: SortSpec; }

export const rangeSortCommand: SheetCommand<RangeSortPayload> = {
  id: "range.sort",
  requiresFreshDerivedState: true,
  validate(payload, ctx) {
    if (typeof payload !== "object" || payload === null || !("spec" in payload)) throw new SheetError("E_VALIDATION", "range.sort requires a SortSpec");
    validateSortSpec(payload.spec);
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    if (payload.spec.range.endRow >= sheet.rowCount || payload.spec.range.endCol >= sheet.columnCount) throw new SheetError("E_INVALID_RANGE", "SortSpec range exceeds worksheet bounds");
    if (conflictsWithFilter(payload.spec.range, sheet.filter)) throw new SheetError("E_VALIDATION", "Sort range intersects active filter rows");
  },
  execute(ctx, payload): CommandOutcome {
    const sheet = ctx.workbook.getSheet(ctx.sheetId);
    const spec = cloneSortSpec(payload.spec);
    const plan = buildSortPlan(sheet, spec);
    if (plan.movedRows === 0) return { result: undefined, journal: null };
    const formulas = collectTranslations(sheet, spec.range, plan);
    applyPermutation(sheet, spec.range, plan, formulas, false);
    emit(ctx.workbook, sheet.id, spec.range, ctx.source);
    return { result: undefined, journal: journal(ctx.workbook, sheet.id, spec.range, plan, formulas) };
  },
};

function cloneSortSpec(spec: SortSpec): SortSpec {
  return { range: { ...spec.range }, hasHeader: spec.hasHeader, keys: spec.keys.map(key => ({ ...key })), ...(spec.locale === undefined ? {} : { locale: spec.locale }) };
}

interface FormulaMove { sourceRow: number; row: number; col: number; original: string; translated: string; }
function collectTranslations(sheet: import("@opensheet/core").Worksheet, range: Range, plan: SortPlan): FormulaMove[] {
  const moves: FormulaMove[] = [];
  for (let source = 0; source < plan.bodyRowCount; source += 1) {
    const dest = plan.sourceToDestination[source]!;
    if (source === dest) continue;
    const sourceRow = plan.bodyStartRow + source;
    const destRow = plan.bodyStartRow + dest;
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const formula = sheet.getCell(sourceRow, col)?.formula;
      if (formula !== undefined) moves.push({ sourceRow, row: destRow, col, original: formula, translated: translateFormulaReferences(formula, destRow - sourceRow, 0, { rowCount: sheet.rowCount, columnCount: sheet.columnCount }) });
    }
  }
  return moves;
}
function applyPermutation(sheet: import("@opensheet/core").Worksheet, range: Range, plan: SortPlan, formulas: readonly FormulaMove[], inverse: boolean): void {
  // One dense column buffer avoids Map hashing while retaining O(rows)
  // temporary space instead of duplicating the whole sorted rectangle.
  const sourceForDestination = inverse ? plan.sourceToDestination : plan.destinationToSource;
  for (let col = range.startCol; col <= range.endCol; col += 1) {
    const values = new Array<CellData | undefined>(plan.bodyRowCount);
    for (let source = 0; source < plan.bodyRowCount; source += 1) values[source] = sheet.getCell(plan.bodyStartRow + source, col);
    for (let destination = 0; destination < plan.bodyRowCount; destination += 1) {
      const data = values[sourceForDestination[destination]!];
      const row = plan.bodyStartRow + destination;
      if (data === undefined) sheet.deleteCell(row, col);
      else sheet.setCell(row, col, data);
    }
  }
  for (const move of formulas) {
    const row = inverse ? move.sourceRow : move.row;
    const data = sheet.getCell(row, move.col);
    if (data !== undefined) sheet.setCell(row, move.col, { ...data, formula: inverse ? move.original : move.translated });
  }
}
function emit(workbook: import("@opensheet/core").Workbook, sheetId: string, range: Range, source: import("@opensheet/shared").ChangeSource): void { workbook.emit({ workbookId: workbook.id, sheetId, changes: [{ range, kind: "reorder" }], source, batch: false }); }
function journal(workbook: import("@opensheet/core").Workbook, sheetId: string, range: Range, plan: SortPlan, formulas: FormulaMove[]): JournalEntry { return { label: "range.sort", affected: [{ sheetId, range, kind: "reorder" }], approxBytes: 256 + plan.destinationToSource.byteLength + plan.sourceToDestination.byteLength + formulas.reduce((n, f) => n + f.original.length + f.translated.length, 0), undo: c => { const s=c.workbook.getSheet(sheetId); applyPermutation(s, range, plan, formulas, true); emit(c.workbook,sheetId,range,c.source); }, redo: c => { const s=c.workbook.getSheet(sheetId); applyPermutation(s, range, plan, formulas, false); emit(c.workbook,sheetId,range,c.source); } }; }
