// Formula AST (M3.2).

import type { CellValue } from "@opensheet/shared";

export interface CellRef {
  row: number; // 0-based
  col: number; // 0-based
  rowAbs: boolean;
  colAbs: boolean;
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

export interface FormulaParseResult {
  ast: Expr;
  /** All cell references this formula touches (ranges expanded), for the dependency graph. */
  dependencies: CellRef[];
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

/** Collect every cell ref (ranges expanded cell-by-cell, deduplicated). */
export function collectDependencies(node: Expr): CellRef[] {
  const refs: CellRef[] = [];
  const seen = new Set<string>();
  const push = (ref: CellRef) => {
    const key = `${ref.row}:${ref.col}`;
    if (seen.has(key)) return;
    seen.add(key);
    refs.push({ ...ref });
  };
  walkExpr(node, (n) => {
    if (n.kind === "cell") push(n.ref);
    else if (n.kind === "range") {
      const r1 = Math.min(n.start.row, n.end.row);
      const r2 = Math.max(n.start.row, n.end.row);
      const c1 = Math.min(n.start.col, n.end.col);
      const c2 = Math.max(n.start.col, n.end.col);
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) push({ row: r, col: c, rowAbs: false, colAbs: false });
      }
    }
  });
  return refs;
}
