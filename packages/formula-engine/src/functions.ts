// Built-in function registry (M3.3/M3.4). Pure value functions over lazy
// FormulaArgument inputs — ranges are consumed streaming, never materialized.
//
// Control-flow forms (IF/AND/OR) are NOT registered here: the evaluator
// handles them as special forms with lazy evaluation.

import type { CellValue } from "@opensheet/shared";
import { isCellError } from "@opensheet/shared";
import type { CellRangeValue, FormulaArgument } from "./evaluate.js";
import { finiteNumber } from "./numeric.js";

/** Iterate every value in an argument (scalar → one value; range → lazy). */
export function* iterateValues(arg: FormulaArgument): Iterable<CellValue> {
  if (isRangeArg(arg)) {
    yield* arg.values();
  } else {
    yield arg;
  }
}

function isRangeArg(arg: FormulaArgument): arg is CellRangeValue {
  return typeof arg === "object" && arg !== null && (arg as CellRangeValue).kind === "range";
}

/** First value of an argument (scalar → itself; range → top-left). */
function scalarOf(arg: FormulaArgument): CellValue {
  if (isRangeArg(arg)) {
    for (const v of arg.values()) return v; // first only
    return null;
  }
  return arg;
}

function toNumber(value: CellValue): number {
  if (isCellError(value)) throw value;
  if (value === null) return 0;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  const n = Number(value);
  if (Number.isNaN(n)) {
    throw { type: "#VALUE!", message: `"${value}" is not numeric` };
  }
  return n;
}

function toText(value: CellValue): string {
  if (isCellError(value)) throw value;
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

export type FunctionImpl = (args: FormulaArgument[]) => CellValue;

export class FunctionRegistry {
  private readonly functions = new Map<string, FunctionImpl>();

  register(name: string, impl: FunctionImpl): void {
    const key = name.toUpperCase();
    if (this.functions.has(key)) {
      throw new Error(`Function already registered: ${key}`);
    }
    this.functions.set(key, impl);
  }

  has(name: string): boolean {
    return this.functions.has(name.toUpperCase());
  }

  get(name: string): FunctionImpl {
    const impl = this.functions.get(name.toUpperCase());
    if (impl === undefined) {
      throw { type: "#NAME?" as const, message: `Unknown function: ${name.toUpperCase()}` };
    }
    return impl;
  }
}

export function createDefaultFunctions(): FunctionRegistry {
  const registry = new FunctionRegistry();

  registry.register("SUM", (args) => {
    let total = 0;
    for (const arg of args) {
      for (const v of iterateValues(arg)) {
        if (isCellError(v)) return v;
        if (v === null || typeof v === "boolean") continue;
        if (typeof v === "string") {
          const n = Number(v);
          if (!Number.isNaN(n)) total += n;
          continue;
        }
        total += v;
      }
    }
    return total;
  });

  registry.register("AVERAGE", (args) => {
    let total = 0;
    let count = 0;
    for (const arg of args) {
      for (const v of iterateValues(arg)) {
        if (isCellError(v)) return v;
        if (v === null || typeof v === "boolean") continue;
        if (typeof v === "string") {
          const n = Number(v);
          if (!Number.isNaN(n)) {
            total += n;
            count++;
          }
          continue;
        }
        total += v;
        count++;
      }
    }
    return count === 0 ? { type: "#DIV/0!", message: "AVERAGE of no values" } : total / count;
  });

  registry.register("MIN", (args) => {
    let min: number | null = null;
    for (const arg of args) {
      for (const v of iterateValues(arg)) {
        if (isCellError(v)) return v;
        if (v === null || typeof v === "boolean") continue;
        const n = typeof v === "string" ? Number(v) : v;
        if (Number.isNaN(n)) continue;
        if (min === null || n < min) min = n;
      }
    }
    return min === null ? 0 : min;
  });

  registry.register("MAX", (args) => {
    let max: number | null = null;
    for (const arg of args) {
      for (const v of iterateValues(arg)) {
        if (isCellError(v)) return v;
        if (v === null || typeof v === "boolean") continue;
        const n = typeof v === "string" ? Number(v) : v;
        if (Number.isNaN(n)) continue;
        if (max === null || n > max) max = n;
      }
    }
    return max === null ? 0 : max;
  });

  registry.register("COUNT", (args) => {
    let count = 0;
    for (const arg of args) {
      for (const v of iterateValues(arg)) {
        if (isCellError(v)) return v;
        if (typeof v === "number") count++;
      }
    }
    return count;
  });

  registry.register("COUNTA", (args) => {
    let count = 0;
    for (const arg of args) {
      for (const v of iterateValues(arg)) {
        if (isCellError(v)) return v;
        if (v !== null) count++;
      }
    }
    return count;
  });

  registry.register("ABS", (args) => {
    return finiteNumber(Math.abs(toNumber(scalarOf(args[0]!))));
  });

  registry.register("ROUND", (args) => {
    const a = toNumber(scalarOf(args[0]!));
    const b = args.length > 1 ? toNumber(scalarOf(args[1]!)) : 0;
    const digits = Number.isInteger(b) ? b : 0;
    const factor = 10 ** digits;
    return finiteNumber(Math.round(a * factor) / factor);
  });

  registry.register("ROUNDUP", (args) => {
    const a = toNumber(scalarOf(args[0]!));
    const b = args.length > 1 ? toNumber(scalarOf(args[1]!)) : 0;
    const digits = Number.isInteger(b) ? b : 0;
    const factor = 10 ** digits;
    const scaled = a * factor;
    return finiteNumber((scaled < 0 ? -Math.ceil(-scaled) : Math.ceil(scaled)) / factor);
  });

  registry.register("ROUNDDOWN", (args) => {
    const a = toNumber(scalarOf(args[0]!));
    const b = args.length > 1 ? toNumber(scalarOf(args[1]!)) : 0;
    const digits = Number.isInteger(b) ? b : 0;
    const factor = 10 ** digits;
    const scaled = a * factor;
    return finiteNumber((scaled < 0 ? -Math.floor(-scaled) : Math.floor(scaled)) / factor);
  });

  registry.register("INT", (args) => {
    return finiteNumber(Math.floor(toNumber(scalarOf(args[0]!))));
  });

  registry.register("MOD", (args) => {
    const a = toNumber(scalarOf(args[0]!));
    const b = args.length > 1 ? toNumber(scalarOf(args[1]!)) : 0;
    if (b === 0) return { type: "#DIV/0!", message: "MOD by zero" };
    return finiteNumber(((a % b) + b) % b);
  });

  registry.register("POWER", (args) => {
    const a = toNumber(scalarOf(args[0]!));
    const b = args.length > 1 ? toNumber(scalarOf(args[1]!)) : 0;
    // Negative base with fractional exponent → NaN → #NUM!.
    return finiteNumber(Math.pow(a, b));
  });

  registry.register("SQRT", (args) => {
    const n = toNumber(scalarOf(args[0]!));
    if (n < 0) return { type: "#NUM!", message: "SQRT of negative number" };
    return Math.sqrt(n);
  });

  registry.register("CONCAT", (args) => {
    let out = "";
    for (const arg of args) {
      for (const v of iterateValues(arg)) {
        if (isCellError(v)) return v;
        out += toText(v);
      }
    }
    return out;
  });

  registry.register("CONCATENATE", registry.get("CONCAT"));

  registry.register("LEN", (args) => {
    const v = scalarOf(args[0]!);
    if (isCellError(v)) return v;
    return toText(v).length;
  });

  registry.register("UPPER", (args) => {
    const v = scalarOf(args[0]!);
    if (isCellError(v)) return v;
    return toText(v).toUpperCase();
  });

  registry.register("LOWER", (args) => {
    const v = scalarOf(args[0]!);
    if (isCellError(v)) return v;
    return toText(v).toLowerCase();
  });

  registry.register("TRIM", (args) => {
    const v = scalarOf(args[0]!);
    if (isCellError(v)) return v;
    return toText(v).trim().replace(/\s+/g, " ");
  });

  registry.register("LEFT", (args) => {
    const v = scalarOf(args[0]!);
    if (isCellError(v)) return v;
    const n = args.length > 1 ? Math.max(0, toNumber(scalarOf(args[1]!))) : 1;
    return toText(v).slice(0, n);
  });

  registry.register("RIGHT", (args) => {
    const v = scalarOf(args[0]!);
    if (isCellError(v)) return v;
    const n = args.length > 1 ? Math.max(0, toNumber(scalarOf(args[1]!))) : 1;
    const text = toText(v);
    return text.slice(Math.max(0, text.length - n));
  });

  registry.register("MID", (args) => {
    const v = scalarOf(args[0]!);
    if (isCellError(v)) return v;
    const text = toText(v);
    const start = args.length > 1 ? toNumber(scalarOf(args[1]!)) : 1;
    const len = args.length > 2 ? toNumber(scalarOf(args[2]!)) : text.length;
    const from = Math.max(0, start - 1);
    return text.slice(from, from + Math.max(0, len));
  });

  registry.register("SUMIF", (args) => {
    if (args.length < 2) throw { type: "#VALUE!" as const, message: "SUMIF needs 2-3 arguments" };
    const criteria = scalarOf(args[1]!);
    const valuesIt = iterateValues(args[0]!)[Symbol.iterator]();
    const sumIt = args.length > 2 ? iterateValues(args[2]!)[Symbol.iterator]() : null;
    let total = 0;
    for (;;) {
      const v = valuesIt.next();
      if (v.done) break;
      // Both iterators advance in lockstep even when the row is skipped.
      const s = sumIt !== null ? sumIt.next() : null;
      const value = v.value;
      if (isCellError(value)) return value;
      const match = typeof criteria === "number" || typeof criteria === "boolean"
        ? toNumber(value) === toNumber(criteria)
        : toText(value) === toText(criteria ?? null);
      if (!match) continue;
      const src = sumIt !== null ? (s === null || s.done ? null : s.value) : value;
      if (isCellError(src)) return src;
      const n = typeof src === "string" ? Number(src) : toNumber(src);
      if (!Number.isNaN(n)) total += n;
    }
    return total;
  });

  registry.register("ISBLANK", (args) => {
    const v = scalarOf(args[0]!);
    if (isCellError(v)) return v;
    return v === null;
  });

  registry.register("ISERROR", (args) => {
    const v = scalarOf(args[0]!);
    return isCellError(v);
  });

  return registry;
}
