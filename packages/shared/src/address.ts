// @injoysai/opensheet-shared — A1 address parsing/formatting

import { SheetError } from "./errors.js";

/** 0-based cell address. */
export interface CellAddress {
  row: number;
  col: number;
}

/** A1-style reference with absolute markers ($A$1). */
export interface CellRef extends CellAddress {
  rowAbs: boolean;
  colAbs: boolean;
}

/** Hard bounds shared by parser and stores (matches common spreadsheet limits). */
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384; // XFD

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/;

/** Convert a 0-based column index to letters: 0 → "A", 26 → "AA". */
export function colToName(col: number): string {
  if (!Number.isInteger(col) || col < 0 || col >= MAX_COLS) {
    throw new SheetError("E_INVALID_ADDRESS", `Column index out of bounds: ${col}`);
  }
  let name = "";
  let n = col + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

/** Convert column letters to a 0-based index: "A" → 0, "AA" → 26. */
export function colFromName(name: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(name)) {
    throw new SheetError("E_INVALID_ADDRESS", `Invalid column name: "${name}"`);
  }
  let col = 0;
  const upper = name.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
  }
  const index = col - 1;
  if (index >= MAX_COLS) {
    throw new SheetError("E_INVALID_ADDRESS", `Column name out of bounds: "${name}"`);
  }
  return index;
}

/** Parse "A1" / "$A$1" / "a$1" into a CellRef (0-based). */
export function parseCellRef(ref: string): CellRef {
  const match = CELL_RE.exec(ref.trim());
  if (!match) {
    throw new SheetError("E_INVALID_ADDRESS", `Invalid cell reference: "${ref}"`);
  }
  const [, colAbsMark, colName, rowAbsMark, rowDigits] = match;
  const row = Number(rowDigits) - 1;
  if (!Number.isInteger(row) || row < 0 || row >= MAX_ROWS) {
    throw new SheetError("E_INVALID_ADDRESS", `Row out of bounds in reference: "${ref}"`);
  }
  return {
    row,
    col: colFromName(colName!),
    rowAbs: rowAbsMark === "$",
    colAbs: colAbsMark === "$",
  };
}

/** Parse "A1" into a 0-based address. Absolute markers are accepted and dropped. */
export function parseAddress(a1: string): CellAddress {
  const ref = parseCellRef(a1);
  return { row: ref.row, col: ref.col };
}

/** Format a 0-based address as A1 notation, optionally with absolute markers. */
export function formatAddress(
  addr: CellAddress,
  abs?: { rowAbs?: boolean; colAbs?: boolean },
): string {
  if (
    !Number.isInteger(addr.row) ||
    !Number.isInteger(addr.col) ||
    addr.row < 0 ||
    addr.row >= MAX_ROWS ||
    addr.col < 0 ||
    addr.col >= MAX_COLS
  ) {
    throw new SheetError(
      "E_INVALID_ADDRESS",
      `Address out of bounds: row=${addr.row} col=${addr.col}`,
    );
  }
  const colPart = `${abs?.colAbs ? "$" : ""}${colToName(addr.col)}`;
  const rowPart = `${abs?.rowAbs ? "$" : ""}${addr.row + 1}`;
  return `${colPart}${rowPart}`;
}
