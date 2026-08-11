import type { WorksheetView } from "@opensheet/core";
import type { CellAddress, CellValue, FindOptions } from "@opensheet/shared";

/** Deterministic sparse find: matching cells are ordered row-major or reverse row-major. */
export function findCells(sheet: WorksheetView, options: FindOptions): CellAddress[] {
  const needle = normalize(options.query, options.matchCase);
  const matches: CellAddress[] = [];
  for (const [row, col, data] of sheet.cellEntries()) {
    const value = options.searchIn === "formulas" ? data.formula : data.value;
    if (value === undefined) continue;
    const text = normalize(displayText(value), options.matchCase);
    if (options.wholeCell ? text === needle : text.includes(needle)) matches.push({ row, col });
  }
  matches.sort((a, b) => options.direction === "forward" ? a.row - b.row || a.col - b.col : b.row - a.row || b.col - a.col);
  return matches;
}

function normalize(value: string, matchCase: boolean): string { return matchCase ? value : value.toLowerCase(); }
function displayText(value: CellValue | string): string {
  if (typeof value === "string") return value;
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return String(value);
  return value.type;
}
