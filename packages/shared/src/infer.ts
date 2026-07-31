// @opensheet/shared — primitive inference (editor + clipboard share this).

import type { CellPrimitive } from "./cell.js";

/**
 * Infer a typed primitive from raw input text (editor commit / TSV paste).
 * Numbers and booleans become typed values; everything else stays a string.
 * Empty → null. Leading "=" stays literal text until the formula engine (M3).
 */
export function inferPrimitive(text: string): CellPrimitive {
  if (text === "") return null;
  if (text === "TRUE") return true;
  if (text === "FALSE") return false;
  // Excel-like numeric inference: optional sign, decimals, exponent.
  if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(text)) {
    const num = Number(text);
    if (Number.isFinite(num)) return num;
  }
  return text;
}

/** Render a stored value for editing (formula source wins over cached value). */
export function cellDisplayText(formula: string | undefined, value: unknown): string {
  if (formula !== undefined) return formula;
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}
