// Formula AST (M3.2).

import type { CellAddress, CellValue } from "@injoysai/opensheet-shared";

export interface CellRef {
  row: number; // 0-based
  col: number; // 0-based
  rowAbs: boolean;
  colAbs: boolean;
}

/** A range reference as written in the formula (NOT expanded to cells). */
export interface CellRangeRef {
  start: CellRef;
  end: CellRef;
}

export type Expr =
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "bool"; value: boolean }
  | { kind: "null" }
  | { kind: "error"; error: CellValue } // literal error like #REF!
  | { kind: "cell"; ref: CellRef }
  | { kind: "range"; start: CellRef; end: CellRef }
  | { kind: "function"; name: string; args: Expr[] }
  | { kind: "unary"; op: "-" | "+" | "%" | "!"; operand: Expr }
  | { kind: "binary"; op: string; left: Expr; right: Expr };

export interface FormulaDependencies {
  /** Individual cell refs (deduplicated). */
  cells: CellRef[];
  /** Range refs as written — consumed lazily, never expanded. */
  ranges: CellRangeRef[];
}

export interface FormulaParseResult {
  ast: Expr;
  dependencies: FormulaDependencies;
}

/** Walk an expression tree, calling `visit` for every node. */
export function walkExpr(node: Expr, visit: (node: Expr) => void): void {
  visit(node);
  switch (node.kind) {
    case "binary":
      walkExpr(node.left, visit);
      walkExpr(node.right, visit);
      break;
    case "unary":
      walkExpr(node.operand, visit);
      break;
    case "function":
      for (const arg of node.args) walkExpr(arg, visit);
      break;
    default:
      break;
  }
}

/**
 * Collect the dependency surface of an AST WITHOUT expanding ranges:
 * individual cells are deduplicated; range refs are kept as intervals so
 * huge ranges (A1:A1000000) never allocate per-cell objects.
 */
export function collectDependencies(node: Expr): FormulaDependencies {
  const cells: CellRef[] = [];
  const seen = new Set<string>();
  const ranges: CellRangeRef[] = [];
  walkExpr(node, (n) => {
    if (n.kind === "cell") {
      const key = `${n.ref.row}:${n.ref.col}`;
      if (!seen.has(key)) {
        seen.add(key);
        cells.push({ ...n.ref });
      }
    } else if (n.kind === "range") {
      ranges.push({ start: { ...n.start }, end: { ...n.end } });
    }
  });
  return { cells, ranges };
}

/** Convert a range ref to its normalized bounds (min corner → max corner). */
export function rangeBounds(range: CellRangeRef): { row1: number; col1: number; row2: number; col2: number } {
  return {
    row1: Math.min(range.start.row, range.end.row),
    col1: Math.min(range.start.col, range.end.col),
    row2: Math.max(range.start.row, range.end.row),
    col2: Math.max(range.start.col, range.end.col),
  };
}

/** Iterate every cell address inside a range (lazy generator). */
export function* iterateRange(range: CellRangeRef): Iterable<CellAddress> {
  const { row1, col1, row2, col2 } = rangeBounds(range);
  for (let row = row1; row <= row2; row++) {
    for (let col = col1; col <= col2; col++) {
      yield { row, col };
    }
  }
}
