// Evaluator (M3.3/M3.4): AST → CellValue, with a pluggable cell resolver.
// Pure: no Worksheet access — the host provides values via FormulaContext.
//
// M3.4 semantics:
//   - Errors are VALUES: the public evaluateExpr never throws a CellError;
//     it returns it. (An uncaught programming error still propagates.)
//   - IF/AND/OR are SPECIAL FORMS evaluated lazily: the untaken branch is
//     never evaluated at all.
//   - Ranges are LAZY: range arguments are passed as iterable views, never
//     expanded to per-cell arrays.

import type { CellRangeRef, CellRef, Expr } from "./ast.js";
import { iterateRange, rangeBounds } from "./ast.js";
import type { CellAddress, CellValue } from "@opensheet/shared";
import { isCellError, type CellError } from "@opensheet/shared";
import { FunctionRegistry } from "./functions.js";

export interface FormulaContext {
  /** Read a single cell's value (errors are values, not exceptions). */
  getCellValue(ref: CellRef): CellValue;
}

/** Lazy iterable view over a range of cell values (no array materialization). */
export interface CellRangeValue {
  readonly kind: "range";
  values(): Iterable<CellValue>;
  addresses(): Iterable<CellAddress>;
}

export type FormulaArgument = CellValue | CellRangeValue;

function makeRangeValue(range: CellRangeRef, ctx: FormulaContext): CellRangeValue {
  return {
    kind: "range",
    *values() {
      for (const address of iterateRange(range)) {
        yield ctx.getCellValue({ row: address.row, col: address.col, rowAbs: false, colAbs: false });
      }
    },
    addresses: () => iterateRange(range),
  };
}

function isErrorLike(value: unknown): value is CellError {
  return (
    typeof value === "object" &&
    value !== null &&
    "type" in value &&
    typeof (value as { type: unknown }).type === "string" &&
    ((value as { type: string }).type.startsWith("#") || (value as { type: string }).type.endsWith("!"))
  );
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

// --- internal evaluator (may throw CellError for propagation) --------------

function evaluateInternal(expr: Expr, ctx: FormulaContext, registry: FunctionRegistry): CellValue {
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
      const { row1, col1 } = rangeBounds({ start: expr.start, end: expr.end });
      return ctx.getCellValue({ row: row1, col: col1, rowAbs: false, colAbs: false });
    }
    case "unary": {
      const operand = evaluateInternal(expr.operand, ctx, registry);
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
    case "function":
      return evaluateFunction(expr, ctx, registry);
  }
}

function evaluateFunction(expr: Extract<Expr, { kind: "function" }>, ctx: FormulaContext, registry: FunctionRegistry): CellValue {
  const name = expr.name;
  // Special forms — lazy evaluation, args evaluated ONLY on the taken path.
  if (name === "IF") {
    if (expr.args.length < 2 || expr.args.length > 3) {
      return { type: "#VALUE!", message: "IF needs 2-3 arguments" };
    }
    const condition = isTruthy(evaluateInternal(expr.args[0]!, ctx, registry));
    if (condition) return evaluateInternal(expr.args[1]!, ctx, registry);
    if (expr.args.length >= 3) return evaluateInternal(expr.args[2]!, ctx, registry);
    return false;
  }
  if (name === "AND") {
    for (const arg of expr.args) {
      if (!isTruthy(evaluateInternal(arg, ctx, registry))) return false;
    }
    return true;
  }
  if (name === "OR") {
    for (const arg of expr.args) {
      if (isTruthy(evaluateInternal(arg, ctx, registry))) return true;
    }
    return false;
  }
  // Ordinary functions: evaluate every argument; ranges stay lazy.
  if (!registry.has(name)) {
    return { type: "#NAME?", message: `Unknown function: ${name}` };
  }
  const impl = registry.get(name);
  const args: FormulaArgument[] = expr.args.map((arg) => {
    if (arg.kind === "range") return makeRangeValue({ start: arg.start, end: arg.end }, ctx);
    return evaluateInternal(arg, ctx, registry);
  });
  return impl(args);
}

function evaluateBinary(
  op: string,
  leftExpr: Expr,
  rightExpr: Expr,
  ctx: FormulaContext,
  registry: FunctionRegistry,
): CellValue {
  const left = evaluateInternal(leftExpr, ctx, registry);
  const right = evaluateInternal(rightExpr, ctx, registry);

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

// --- public boundary -------------------------------------------------------

/**
 * Evaluate a parsed AST. NEVER throws a CellError: propagated errors are
 * returned as values. Programming bugs and SheetErrors still propagate.
 */
export function evaluateExpr(expr: Expr, ctx: FormulaContext, registry: FunctionRegistry): CellValue {
  try {
    return evaluateInternal(expr, ctx, registry);
  } catch (error) {
    if (isErrorLike(error) && isCellError(error as CellError)) {
      return error as CellError;
    }
    throw error;
  }
}
