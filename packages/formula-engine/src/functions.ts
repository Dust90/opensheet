// Built-in function registry (M3.3). Pure value functions over CellValue.

import type { CellValue } from "@opensheet/shared";
import { isCellError, type CellError } from "@opensheet/shared";

export type ScalarArg = string | number | boolean | null | CellError;

/** Convert a value for arithmetic (numbers stay, booleans→1/0, strings→NaN). */
function toNumber(value: ScalarArg): number {
  if (isCellError(value)) throw value;
  if (value === null) return 0;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  const n = Number(value);
  return Number.isNaN(n) ? NaN : n;
}

function toText(value: ScalarArg): string {
  if (isCellError(value)) throw value;
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

export type FunctionImpl = (args: ScalarArg[][]) => CellValue;

/**
 * A spreadsheet function receives ARRAY arguments (each arg may be a scalar
 * or a range expansion). The registry maps names to implementations; the
 * evaluator flattens cell/range args into value matrices.
 */
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

/** Number-cast an argument matrix; throws on CellError cells. */
function numericMatrix(arg: ScalarArg[]): number[] {
  return arg.map((v) => {
    if (isCellError(v)) throw v;
    const n = toNumber(v);
    if (Number.isNaN(n) && typeof v === "string") {
      throw { type: "#VALUE!" as const, message: `Cannot coerce "${v}" to a number` };
    }
    return n;
  });
}

/** Flatten a single scalar-or-array argument into values. */
function values(arg: ScalarArg[]): ScalarArg[] {
  return arg;
}

function firstError(args: ScalarArg[][]): CellError | null {
  for (const arg of args) {
    for (const v of arg) {
      if (isCellError(v)) return v;
    }
  }
  return null;
}

export function createDefaultFunctions(): FunctionRegistry {
  const registry = new FunctionRegistry();

  registry.register("SUM", (args) => {
    let total = 0;
    for (const arg of args) {
      for (const v of arg) {
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
      for (const v of arg) {
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
      for (const v of arg) {
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
      for (const v of arg) {
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
      for (const v of arg) {
        if (isCellError(v)) return v;
        if (typeof v === "number") count++;
      }
    }
    return count;
  });

  registry.register("COUNTA", (args) => {
    let count = 0;
    for (const arg of args) {
      for (const v of arg) {
        if (isCellError(v)) return v;
        if (v !== null) count++;
      }
    }
    return count;
  });

  registry.register("IF", (args) => {
    if (args.length < 2) throw { type: "#VALUE!" as const, message: "IF needs 2-3 arguments" };
    const cond = values(args[0]!);
    const condition = cond.length === 0 ? false : toNumber(cond[0]!) !== 0;
    // Short-circuit: only the selected branch's errors surface.
    if (condition) {
      const yes = values(args[1]!);
      return yes[0] ?? null;
    }
    if (args.length >= 3) {
      const no = values(args[2]!);
      return no[0] ?? null;
    }
    return false;
  });

  registry.register("AND", (args) => {
    for (const arg of args) {
      for (const v of arg) {
        if (isCellError(v)) return v;
        // FALSE, 0 and blank are falsy (Excel: AND(TRUE,0) = FALSE).
        if (v === null || v === false || v === 0) return false;
      }
    }
    return true;
  });

  registry.register("OR", (args) => {
    for (const arg of args) {
      for (const v of arg) {
        if (isCellError(v)) return v;
        if (v === true || v === 1) return true;
      }
    }
    return false;
  });

  registry.register("NOT", (args) => {
    const v = values(args[0] ?? []);
    if (v.length === 0 || v[0] === undefined || isCellError(v[0])) return v[0] ?? null;
    return toNumber(v[0]!) === 0;
  });

  registry.register("ABS", (args) => {
    const nums = numericMatrix(values(args[0] ?? []));
    return Math.abs(nums[0] ?? 0);
  });

  registry.register("ROUND", (args) => {
    const [a, b] = numericMatrix(values(args[0] ?? [])).concat(numericMatrix(values(args[1] ?? [])));
    const bNum = b ?? 0;
    const digits = Number.isInteger(bNum) ? bNum : 0;
    const factor = 10 ** digits;
    return Math.round((a ?? 0) * factor) / factor;
  });

  registry.register("ROUNDUP", (args) => {
    const [a, b] = numericMatrix(values(args[0] ?? [])).concat(numericMatrix(values(args[1] ?? [])));
    const bNum = b ?? 0;
    const digits = Number.isInteger(bNum) ? bNum : 0;
    const factor = 10 ** digits;
    const scaled = (a ?? 0) * factor;
    return (scaled < 0 ? -Math.ceil(-scaled) : Math.ceil(scaled)) / factor;
  });

  registry.register("ROUNDDOWN", (args) => {
    const [a, b] = numericMatrix(values(args[0] ?? [])).concat(numericMatrix(values(args[1] ?? [])));
    const bNum = b ?? 0;
    const digits = Number.isInteger(bNum) ? bNum : 0;
    const factor = 10 ** digits;
    const scaled = (a ?? 0) * factor;
    return (scaled < 0 ? -Math.floor(-scaled) : Math.floor(scaled)) / factor;
  });

  registry.register("INT", (args) => {
    const nums = numericMatrix(values(args[0] ?? []));
    return Math.floor(nums[0] ?? 0);
  });

  registry.register("MOD", (args) => {
    const [a, b] = numericMatrix(values(args[0] ?? [])).concat(numericMatrix(values(args[1] ?? [])));
    if ((b ?? 0) === 0) return { type: "#DIV/0!", message: "MOD by zero" };
    const bNum = b ?? 0;
    return ((a ?? 0) % bNum + bNum) % bNum;
  });

  registry.register("POWER", (args) => {
    const [a, b] = numericMatrix(values(args[0] ?? [])).concat(numericMatrix(values(args[1] ?? [])));
    return Math.pow(a ?? 0, b ?? 0);
  });

  registry.register("SQRT", (args) => {
    const nums = numericMatrix(values(args[0] ?? []));
    const n = nums[0] ?? 0;
    if (n < 0) return { type: "#NUM!", message: "SQRT of negative number" };
    return Math.sqrt(n);
  });

  registry.register("CONCAT", (args) => {
    let out = "";
    for (const arg of args) {
      for (const v of arg) {
        if (isCellError(v)) return v;
        out += toText(v);
      }
    }
    return out;
  });

  registry.register("CONCATENATE", registry.get("CONCAT"));

  registry.register("LEN", (args) => {
    const v = values(args[0] ?? []);
    if (v.length === 0) return 0;
    if (v[0] === undefined || isCellError(v[0])) return v[0] ?? null;
    return toText(v[0]!).length;
  });

  registry.register("UPPER", (args) => {
    const v = values(args[0] ?? []);
    if (v.length === 0) return "";
    if (v[0] === undefined || isCellError(v[0])) return v[0] ?? null;
    return toText(v[0]!).toUpperCase();
  });

  registry.register("LOWER", (args) => {
    const v = values(args[0] ?? []);
    if (v.length === 0) return "";
    if (v[0] === undefined || isCellError(v[0])) return v[0] ?? null;
    return toText(v[0]!).toLowerCase();
  });

  registry.register("TRIM", (args) => {
    const v = values(args[0] ?? []);
    if (v.length === 0) return "";
    if (v[0] === undefined || isCellError(v[0])) return v[0] ?? null;
    return toText(v[0]!).trim().replace(/\s+/g, " ");
  });

  registry.register("LEFT", (args) => {
    const v = values(args[0] ?? []);
    const nums = numericMatrix(values(args[1] ?? []));
    if (v.length === 0 || v[0] === undefined || isCellError(v[0])) return v[0] ?? "";
    const text = toText(v[0]!);
    return text.slice(0, Math.max(0, nums[0] ?? 1));
  });

  registry.register("RIGHT", (args) => {
    const v = values(args[0] ?? []);
    const nums = numericMatrix(values(args[1] ?? []));
    if (v.length === 0 || v[0] === undefined || isCellError(v[0])) return v[0] ?? "";
    const text = toText(v[0]!);
    const n = Math.max(0, nums[0] ?? 1);
    return text.slice(Math.max(0, text.length - n));
  });

  registry.register("MID", (args) => {
    const v = values(args[0] ?? []);
    const [start, len] = numericMatrix(values(args[1] ?? [])).concat(numericMatrix(values(args[2] ?? [])));
    if (v.length === 0 || v[0] === undefined || isCellError(v[0])) return v[0] ?? "";
    const text = toText(v[0]!);
    const startNum = start ?? 1;
    const lenNum = len ?? text.length;
    const from = Math.max(0, (Number.isFinite(startNum) ? startNum : 1) - 1);
    const count = Number.isFinite(lenNum) ? lenNum : text.length;
    return text.slice(from, from + Math.max(0, count));
  });

  registry.register("SUMIF", (args) => {
    // SUMIF(range, criteria, [sum_range]) — criteria only supports number/string equality at M3.
    if (args.length < 2) throw { type: "#VALUE!" as const, message: "SUMIF needs 2-3 arguments" };
    const range = args[0]!;
    const criteria = values(args[1]!)[0];
    const sumRange = args[2];
    let total = 0;
    for (let i = 0; i < range.length; i++) {
      const v = range[i]!;
      if (isCellError(v)) return v;
      const match = typeof criteria === "number" || typeof criteria === "boolean"
        ? toNumber(v) === toNumber(criteria)
        : toText(v) === toText(criteria ?? null);
      if (match) {
        const src = sumRange !== undefined ? sumRange[i]! : v;
        if (isCellError(src)) return src;
        const n = typeof src === "string" ? Number(src) : toNumber(src);
        if (!Number.isNaN(n)) total += n;
      }
    }
    return total;
  });

  registry.register("ISBLANK", (args) => {
    const v = values(args[0] ?? []);
    if (v.length === 0 || v[0] === undefined || isCellError(v[0])) return v[0] ?? null;
    return v[0] === null;
  });

  registry.register("ISERROR", (args) => {
    const v = values(args[0] ?? []);
    if (v.length === 0 || v[0] === undefined) return false;
    return isCellError(v[0]);
  });

  return registry;
}

export { firstError };
