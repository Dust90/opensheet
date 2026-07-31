// Numeric guards (M3.5): shared by the evaluator and the function registry.

import type { CellValue } from "@opensheet/shared";

/** Unified finite-number gate: any non-finite numeric result → #NUM!. */
export function finiteNumber(value: number): CellValue {
  return Number.isFinite(value) ? value : { type: "#NUM!", message: "Numeric result is not finite" };
}
