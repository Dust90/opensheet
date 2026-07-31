// Evaluator (M3.3): AST → CellValue, with a pluggable cell resolver.
// Pure: no Worksheet access — the host provides values via FormulaContext.

import type { CellRef, Expr } from "./ast.js";
import type { CellValue } from "@opensheet/shared";
import { isCellError, type CellError } from "@opensheet/shared";
import { FunctionRegistry, type ScalarArg } from "./functions.js";

export interface FormulaContext {
  /** Read a single cell's value (errors are values, not exceptions). */
  getCellValue(ref: CellRef): CellValue;
}

function isTruthy(value: CellValue): boolean {
  if (isCellError(value)) throw value;
  if (value === null) return false;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  return value.length > 0;
}

/** Excel-like equality: numbers vs numeric strings compare numerically. */
function eq(a: CellValue, b: CellValue): boolean {
  if (isCellError(a)) throw a;
  if (isCellError(b)) throw b;
  if (typeof a === "number" && typeof b === "number") return a === b;
  const an = typeof a === "string" ? Number(a) : a;
  const bn = typeof b === "string" ? Number(b) : b;
  if (typeof an === "number" && typeof bn === "number" && !Number.isNaN(an) && !Number.isNaN(bn)) {
    return an === bn;
  }
  return a === b;
}

function compare(a: CellValue, b: CellValue): number {
  if (isCellError(a)) throw a;
  if (isCellError(b)) throw b;
  if (typeof a === "number" && typeof b === "number") return a - b;
  return String(a ?? "").localeCompare(String(b ?? ""));
}

/** Number coercion used by arithmetic (null→0, bool→1/0, numeric strings parse). */
function toNum(value: CellValue): number {
  if (isCellError(value)) throw value;
  if (value === null) return 0;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  const n = Number(value);
  if (Number.isNaN(n)) {
    throw { type: "#VALUE!", message: `"${value}" is not numeric` } satisfies CellError;
  }
  return n;
}

function toText(value: CellValue): string {
  if (isCellError(value)) throw value;
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

/**
 * Evaluate a parsed AST. Errors surface as CellError VALUES (except inside
 * control-flow functions like IF where the unevaluated branch never runs).
 */
export function evaluateExpr(expr: Expr, ctx: FormulaContext, registry: FunctionRegistry): CellValue {
  switch (expr.kind) {
    case "number":
      return expr.value;
    case "string":
      return expr.value;
    case "bool":
      return expr.value;
    case "null":
      return null;
    case "error":
      return expr.error;
    case "cell":
      return ctx.getCellValue(expr.ref);
    case "range": {
      // A bare range in a scalar position: Excel returns the top-left cell.
      const topLeft: CellRef = {
        row: Math.min(expr.start.row, expr.end.row),
        col: Math.min(expr.start.col, expr.end.col),
        rowAbs: false,
        colAbs: false,
      };
      return ctx.getCellValue(topLeft);
    }
    case "unary": {
      const operand = evaluateExpr(expr.operand, ctx, registry);
      switch (expr.op) {
        case "-":
          return -toNum(operand);
        case "+":
          return toNum(operand);
        case "%":
          return toNum(operand) / 100;
        case "!":
          return !isTruthy(operand);
        default:
          return operand;
      }
    }
    case "binary":
      return evaluateBinary(expr.op, expr.left, expr.right, ctx, registry);
    case "function": {
      if (registry.has(expr.name)) {
        const impl = registry.get(expr.name);
        const args: ScalarArg[][] = expr.args.map((arg) => {
          if (arg.kind === "range") {
            return expandRange(arg.start, arg.end, ctx);
          }
          return [evaluateExpr(arg, ctx, registry)];
        });
        return impl(args);
      }
      return { type: "#NAME?", message: `Unknown function: ${expr.name}` };
    }
  }
}

function expandRange(start: CellRef, end: CellRef, ctx: FormulaContext): ScalarArg[] {
  const r1 = Math.min(start.row, end.row);
  const r2 = Math.max(start.row, end.row);
  const c1 = Math.min(start.col, end.col);
  const c2 = Math.max(start.col, end.col);
  const out: ScalarArg[] = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      out.push(ctx.getCellValue({ row: r, col: c, rowAbs: false, colAbs: false }));
    }
  }
  return out;
}

function evaluateBinary(
  op: string,
  leftExpr: Expr,
  rightExpr: Expr,
  ctx: FormulaContext,
  registry: FunctionRegistry,
): CellValue {
  // Short-circuit boolean logic: errors in the untaken branch must not leak.
  if (op === "&&") return isTruthy(evaluateExpr(leftExpr, ctx, registry)) && isTruthy(evaluateExpr(rightExpr, ctx, registry));
  if (op === "||") return isTruthy(evaluateExpr(leftExpr, ctx, registry)) || isTruthy(evaluateExpr(rightExpr, ctx, registry));

  const left = evaluateExpr(leftExpr, ctx, registry);
  const right = evaluateExpr(rightExpr, ctx, registry);

  switch (op) {
    case "=":
      return eq(left, right);
    case "<>":
      return !eq(left, right);
    case "<":
      return compare(left, right) < 0;
    case ">":
      return compare(left, right) > 0;
    case "<=":
      return compare(left, right) <= 0;
    case ">=":
      return compare(left, right) >= 0;
    case "&":
      return toText(left) + toText(right);
    case "+":
      return toNum(left) + toNum(right);
    case "-":
      return toNum(left) - toNum(right);
    case "*":
      return toNum(left) * toNum(right);
    case "/": {
      const divisor = toNum(right);
      if (divisor === 0) return { type: "#DIV/0!", message: "Division by zero" };
      return toNum(left) / divisor;
    }
    case "^":
      return Math.pow(toNum(left), toNum(right));
    default:
      return { type: "#VALUE!", message: `Unsupported operator ${op}` };
  }
}
