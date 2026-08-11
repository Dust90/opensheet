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
import { FunctionRegistry, SPECIAL_FORM_NAMES } from "./functions.js";
import { finiteNumber } from "./numeric.js";

const [IF_NAME, AND_NAME, OR_NAME] = SPECIAL_FORM_NAMES;

export interface FormulaContext {
  /** Read a single cell's value (errors are values, not exceptions). */
  getCellValue(ref: CellRef): CellValue;
}

/**
 * Deterministic evaluation budget (M3.5 guardrail): a single formula may
 * consume at most `maxCellReads` evaluation budget units; the whole
 * transaction budget is enforced by the host sharing one budget across
 * formulas. Exceeding the limit yields #VALUE! instead of blocking the main
 * thread for millions of iterations.
 *
 * NOTE on naming: one "unit" is not always one cell read — range nodes and
 * other AST overhead also consume units (e.g. SUM(A1:A3) costs 4, not 3).
 * A future split into `maxCellReads` + `maxOperations` may refine this; the
 * current semantics safely prevent runaway evaluation.
 */
export interface EvaluationBudget {
  maxCellReads: number;
  remaining: number;
  /** Consume `n` reads; returns false when the budget is exhausted. */
  consume(n?: number): boolean;
}

export function makeBudget(maxCellReads: number): EvaluationBudget {
  const state = { remaining: maxCellReads };
  return {
    maxCellReads,
    get remaining() {
      return state.remaining;
    },
    consume: (n = 1) => {
      if (state.remaining < n) return false;
      state.remaining -= n;
      return true;
    },
  };
}

/** Lazy iterable view over a range of cell values (no array materialization). */
export interface CellRangeValue {
  readonly kind: "range";
  values(): Iterable<CellValue>;
  addresses(): Iterable<CellAddress>;
}

export type FormulaArgument = CellValue | CellRangeValue;

function makeRangeValue(range: CellRangeRef, ctx: FormulaContext, budget: EvaluationBudget | undefined): CellRangeValue {
  return {
    kind: "range",
    *values() {
      // Do not route value iteration through iterateRange(): that generator
      // creates one CellAddress for every cell, followed by another CellRef
      // for the resolver. Large aggregate ranges are intentionally streamed,
      // so avoid those short-lived intermediate address objects too.
      const { row1, col1, row2, col2 } = rangeBounds(range);
      for (let row = row1; row <= row2; row += 1) {
        for (let col = col1; col <= col2; col += 1) {
          if (budget !== undefined && !budget.consume(1)) {
            throw { type: "#VALUE!", message: "Formula evaluation limit exceeded" };
          }
          yield ctx.getCellValue({ row, col, rowAbs: false, colAbs: false });
        }
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

function evaluateInternal(
  expr: Expr,
  ctx: FormulaContext,
  registry: FunctionRegistry,
  budget: EvaluationBudget | undefined,
): CellValue {
  if (budget !== undefined && !budget.consume(1)) {
    throw { type: "#VALUE!", message: "Formula evaluation limit exceeded" };
  }
  switch (expr.kind) {
    case "number":
      return finiteNumber(expr.value);
    case "string":
      return expr.value;
    case "bool":
      return expr.value;
    case "null":
      return null;
    case "error":
      return expr.error;
    case "cell":
      if (budget !== undefined && !budget.consume(1)) {
        throw { type: "#VALUE!", message: "Formula evaluation limit exceeded" };
      }
      return ctx.getCellValue(expr.ref);
    case "range": {
      // A bare range in a scalar position: Excel returns the top-left cell.
      const { row1, col1 } = rangeBounds({ start: expr.start, end: expr.end });
      if (budget !== undefined && !budget.consume(1)) {
        throw { type: "#VALUE!", message: "Formula evaluation limit exceeded" };
      }
      return ctx.getCellValue({ row: row1, col: col1, rowAbs: false, colAbs: false });
    }
    case "unary": {
      const operand = evaluateInternal(expr.operand, ctx, registry, budget);
      switch (expr.op) {
        case "-":
          return finiteNumber(-toNum(operand));
        case "+":
          return finiteNumber(toNum(operand));
        case "%":
          return finiteNumber(toNum(operand) / 100);
        case "!":
          return !isTruthy(operand);
        default:
          return operand;
      }
    }
    case "binary":
      return evaluateBinary(expr.op, expr.left, expr.right, ctx, registry, budget);
    case "function":
      return evaluateFunction(expr, ctx, registry, budget);
  }
}

function evaluateFunction(
  expr: Extract<Expr, { kind: "function" }>,
  ctx: FormulaContext,
  registry: FunctionRegistry,
  budget: EvaluationBudget | undefined,
): CellValue {
  const name = expr.name;
  // Special forms — lazy evaluation, args evaluated ONLY on the taken path.
  if (name === IF_NAME) {
    if (expr.args.length < 2 || expr.args.length > 3) {
      return { type: "#VALUE!", message: "IF needs 2-3 arguments" };
    }
    const condition = isTruthy(evaluateInternal(expr.args[0]!, ctx, registry, budget));
    if (condition) return evaluateInternal(expr.args[1]!, ctx, registry, budget);
    if (expr.args.length >= 3) return evaluateInternal(expr.args[2]!, ctx, registry, budget);
    return false;
  }
  if (name === AND_NAME) {
    for (const arg of expr.args) {
      if (!isTruthy(evaluateInternal(arg, ctx, registry, budget))) return false;
    }
    return true;
  }
  if (name === OR_NAME) {
    for (const arg of expr.args) {
      if (isTruthy(evaluateInternal(arg, ctx, registry, budget))) return true;
    }
    return false;
  }
  // Ordinary functions: evaluate every argument; ranges stay lazy.
  if (!registry.has(name)) {
    return { type: "#NAME?", message: `Unknown function: ${name}` };
  }
  const impl = registry.get(name);
  const args: FormulaArgument[] = expr.args.map((arg) => {
    if (arg.kind === "range") return makeRangeValue({ start: arg.start, end: arg.end }, ctx, budget);
    return evaluateInternal(arg, ctx, registry, budget);
  });
  return impl(args);
}

function evaluateBinary(
  op: string,
  leftExpr: Expr,
  rightExpr: Expr,
  ctx: FormulaContext,
  registry: FunctionRegistry,
  budget: EvaluationBudget | undefined,
): CellValue {
  const left = evaluateInternal(leftExpr, ctx, registry, budget);
  const right = evaluateInternal(rightExpr, ctx, registry, budget);

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
      return finiteNumber(toNum(left) + toNum(right));
    case "-":
      return finiteNumber(toNum(left) - toNum(right));
    case "*":
      return finiteNumber(toNum(left) * toNum(right));
    case "/": {
      const divisor = toNum(right);
      if (divisor === 0) return { type: "#DIV/0!", message: "Division by zero" };
      return finiteNumber(toNum(left) / divisor);
    }
    case "^":
      return finiteNumber(Math.pow(toNum(left), toNum(right)));
    default:
      return { type: "#VALUE!", message: `Unsupported operator ${op}` };
  }
}

// --- public boundary -------------------------------------------------------

/**
 * Evaluate a parsed AST. NEVER throws a CellError: propagated errors are
 * returned as values. Programming bugs and SheetErrors still propagate.
 * `budget` optionally caps cell reads per formula (M3.5 guardrail).
 */
export function evaluateExpr(
  expr: Expr,
  ctx: FormulaContext,
  registry: FunctionRegistry,
  budget?: EvaluationBudget,
): CellValue {
  try {
    return evaluateInternal(expr, ctx, registry, budget);
  } catch (error) {
    if (isErrorLike(error) && isCellError(error as CellError)) {
      return error as CellError;
    }
    throw error;
  }
}
