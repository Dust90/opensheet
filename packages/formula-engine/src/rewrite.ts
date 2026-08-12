// Formula reference rewriting (M3.5): after row/column insert/delete, cell
// references inside stored formulas must be adjusted the way Excel does:
//   - Relative references shift with the structure change.
//   - Absolute references ($A$1 / $A1 / A$1) do NOT shift on their
//     absolute axis; the relative axis follows the change.
//   - References that point INTO deleted rows/cols become #REF!.
//
// Implementation: parse → rewrite AST refs → serialize back to text, so the
// stored formula source always matches what the evaluator parses.

import { colToName } from "@injoysai/opensheet-shared";
import type { CellRef, Expr } from "./ast.js";
import { parseFormula } from "./parser.js";

export type StructureChange =
  | { type: "insertRows"; at: number; count: number }
  | { type: "deleteRows"; at: number; count: number }
  | { type: "insertColumns"; at: number; count: number }
  | { type: "deleteColumns"; at: number; count: number };

/** Rewrite a stored formula body's references for a structure change. */
export function rewriteFormulaReferences(formula: string, change: StructureChange): string {
  const { ast } = parseFormula(formula);
  const rewritten = rewriteExpr(ast, change);
  return "=" + exprToString(rewritten);
}

export interface FormulaReferenceBounds {
  rowCount: number;
  columnCount: number;
}

/**
 * Translate a formula together with the cell that owns it (for range sort).
 * Unlike structure rewriting, every relative axis moves by the supplied
 * delta; absolute axes remain fixed. A translated reference outside the
 * destination worksheet becomes #REF!.
 */
export function translateFormulaReferences(
  formula: string,
  deltaRow: number,
  deltaCol: number,
  bounds: FormulaReferenceBounds,
): string {
  const { ast } = parseFormula(formula);
  const translated = translateExpr(ast, deltaRow, deltaCol, bounds);
  return "=" + exprToString(translated);
}

function translateExpr(
  node: Expr,
  deltaRow: number,
  deltaCol: number,
  bounds: FormulaReferenceBounds,
): Expr {
  switch (node.kind) {
    case "cell": {
      const ref = translateCellRef(node.ref, deltaRow, deltaCol, bounds);
      return ref === null ? { kind: "error", error: { type: "#REF!" } } : { ...node, ref };
    }
    case "range": {
      const start = translateCellRef(node.start, deltaRow, deltaCol, bounds);
      const end = translateCellRef(node.end, deltaRow, deltaCol, bounds);
      return start === null || end === null ? { kind: "error", error: { type: "#REF!" } } : { ...node, start, end };
    }
    case "binary":
      return { ...node, left: translateExpr(node.left, deltaRow, deltaCol, bounds), right: translateExpr(node.right, deltaRow, deltaCol, bounds) };
    case "unary":
      return { ...node, operand: translateExpr(node.operand, deltaRow, deltaCol, bounds) };
    case "function":
      return { ...node, args: node.args.map((arg) => translateExpr(arg, deltaRow, deltaCol, bounds)) };
    default:
      return node;
  }
}

function translateCellRef(
  ref: CellRef,
  deltaRow: number,
  deltaCol: number,
  bounds: FormulaReferenceBounds,
): CellRef | null {
  const row = ref.rowAbs ? ref.row : ref.row + deltaRow;
  const col = ref.colAbs ? ref.col : ref.col + deltaCol;
  if (row < 0 || col < 0 || row >= bounds.rowCount || col >= bounds.columnCount) return null;
  return { row, col, rowAbs: ref.rowAbs, colAbs: ref.colAbs };
}

function rewriteExpr(node: Expr, change: StructureChange): Expr {
  switch (node.kind) {
    case "cell": {
      const ref = rewriteCellRef(node.ref, change);
      if (ref === null) return { kind: "error", error: { type: "#REF!" } };
      return { ...node, ref };
    }
    case "range": {
      const start = rewriteCellRef(node.start, change);
      const end = rewriteCellRef(node.end, change);
      if (start === null || end === null) return { kind: "error", error: { type: "#REF!" } };
      return { ...node, start, end };
    }
    case "binary":
      return { ...node, left: rewriteExpr(node.left, change), right: rewriteExpr(node.right, change) };
    case "unary":
      return { ...node, operand: rewriteExpr(node.operand, change) };
    case "function":
      return { ...node, args: node.args.map((arg) => rewriteExpr(arg, change)) };
    default:
      return node;
  }
}

/** Rewrite one ref; null means the ref is now broken (→ #REF!). */
function rewriteCellRef(ref: CellRef, change: StructureChange): CellRef | null {
  let row = ref.row;
  let col = ref.col;
  let rowBroken = false;
  let colBroken = false;

  switch (change.type) {
    case "insertRows": {
      // Rows at/after the insertion point shift down — relative refs only.
      if (!ref.rowAbs && row >= change.at) row += change.count;
      break;
    }
    case "deleteRows": {
      if (!ref.rowAbs) {
        if (row >= change.at + change.count) row -= change.count;
        else if (row >= change.at) rowBroken = true; // pointed into deleted rows
      } else if (row >= change.at && row < change.at + change.count) {
        rowBroken = true; // absolute ref into deleted rows
      }
      break;
    }
    case "insertColumns": {
      if (!ref.colAbs && col >= change.at) col += change.count;
      break;
    }
    case "deleteColumns": {
      if (!ref.colAbs) {
        if (col >= change.at + change.count) col -= change.count;
        else if (col >= change.at) colBroken = true;
      } else if (col >= change.at && col < change.at + change.count) {
        colBroken = true;
      }
      break;
    }
  }

  if (rowBroken || colBroken) return null;
  return { row, col, rowAbs: ref.rowAbs, colAbs: ref.colAbs };
}

/** Serialize an AST back to formula text with precedence-aware parens. */
export function exprToString(node: Expr): string {
  return exprNodeToString(node, 0, "left");
}

const PRECEDENCE: Record<string, number> = {
  "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
  "&": 2,
  "+": 3, "-": 3,
  "*": 4, "/": 4,
  "^": 5,
};

/**
 * True for ALL left-associative binary operators. In IEEE 754 floating-point,
 * even addition and multiplication are NOT truly associative
 * (e.g. 1e16+(-1e16+1) ≠ 1e16+(-1e16)+1), so the right child at EQUAL
 * precedence must always keep its parentheses to guarantee AST round-trip
 * fidelity after any structural rewrite.
 *
 * `^` is right-associative and is handled separately via the side=="left" check.
 */
function isLeftAssoc(op: string): boolean {
  // Every binary operator in PRECEDENCE is left-associative except `^`.
  return op !== "^" && op in PRECEDENCE;
}

function precOf(node: Expr): number {
  return node.kind === "binary" ? (PRECEDENCE[node.op] ?? 0) : node.kind === "unary" ? 6 : 7;
}

/**
 * Serialize a node with awareness of its parent's precedence and which side
 * (left/right) it occupies.
 *
 * Rules:
 *  - `ownPrec < parentPrec`  → always needs parens.
 *  - `^` left child at prec 5 → needs parens (right-associativity).
 *  - Left-assoc non-commutative operators: right child at EQUAL precedence
 *    needs parens to preserve semantics (parentPrec += 0.5 trick).
 */
function exprNodeToString(node: Expr, parentPrec: number, side: "left" | "right"): string {
  const inner = rawToString(node);
  const ownPrec = precOf(node);
  const needParens =
    ownPrec < parentPrec ||
    (node.kind === "binary" && node.op === "^" && side === "left" && parentPrec === 5);
  return needParens ? `(${inner})` : inner;
}

function rawToString(node: Expr): string {
  switch (node.kind) {
    case "number":
      return String(node.value);
    case "string":
      return `"${node.value.replace(/"/g, '""')}"`;
    case "bool":
      return node.value ? "TRUE" : "FALSE";
    case "null":
      return "NULL";
    case "error":
      return (node.error as { type: string }).type;
    case "cell":
      return refToString(node.ref);
    case "range":
      return `${refToString(node.start)}:${refToString(node.end)}`;
    case "function":
      return `${node.name}(${node.args.map((arg) => exprNodeToString(arg, 0, "left")).join(",")})`;
    case "unary":
      if (node.op === "%") return `${exprNodeToString(node.operand, 6, "left")}%`;
      return `${node.op}${exprNodeToString(node.operand, 6, "left")}`;
    case "binary": {
      const prec = PRECEDENCE[node.op] ?? 0;
      const left = exprNodeToString(node.left, prec, "left");
      // For ALL left-associative operators the right child at equal precedence
      // must keep its parentheses, even for + and * — IEEE 754 floating-point
      // is NOT truly associative (1e16+(-1e16+1) ≠ 1e16+(-1e16)+1). The 0.5
      // bump is smaller than any integer step in PRECEDENCE, so it only
      // triggers when ownPrec === prec (equal-precedence same-level grouping).
      const rightPrec = isLeftAssoc(node.op) ? prec + 0.5 : prec;
      const right = exprNodeToString(node.right, rightPrec, "right");
      return `${left}${node.op}${right}`;
    }
  }
}

function refToString(ref: CellRef): string {
  const col = colToName(ref.col);
  const row = ref.row + 1;
  return `${ref.colAbs ? "$" : ""}${col}${ref.rowAbs ? "$" : ""}${row}`;
}

/** Cell refs this formula references AFTER the rewrite (for graph rebuild). */
export function rewrittenDependencies(formula: string, change: StructureChange): {
  cells: CellRef[];
  ranges: { start: CellRef; end: CellRef }[];
} {
  const { ast } = parseFormula(formula);
  const rewritten = rewriteExpr(ast, change);
  const cells: CellRef[] = [];
  const ranges: { start: CellRef; end: CellRef }[] = [];
  const seen = new Set<string>();
  const collect = (node: Expr): void => {
    switch (node.kind) {
      case "cell": {
        const key = `${node.ref.row}:${node.ref.col}`;
        if (!seen.has(key)) {
          seen.add(key);
          cells.push({ ...node.ref });
        }
        break;
      }
      case "range":
        ranges.push({ start: { ...node.start }, end: { ...node.end } });
        break;
      case "binary":
        collect(node.left);
        collect(node.right);
        break;
      case "unary":
        collect(node.operand);
        break;
      case "function":
        for (const arg of node.args) collect(arg);
        break;
      default:
        break;
    }
  };
  collect(rewritten);
  return { cells, ranges };
}
