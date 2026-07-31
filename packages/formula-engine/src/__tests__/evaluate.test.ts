import { describe, expect, it } from "vitest";
import { parseFormula } from "../parser.js";
import { evaluateExpr, type FormulaContext } from "../evaluate.js";
import { createDefaultFunctions } from "../functions.js";
import type { CellRef, CellValue } from "@opensheet/shared";

/** Test context backed by a plain map. */
function makeContext(values: Record<string, CellValue>): FormulaContext {
  return {
    getCellValue: (ref: CellRef) => values[`${colName(ref.col)}${ref.row + 1}`] ?? null,
  };
}

function colName(col: number): string {
  let name = "";
  let n = col + 1;
  while (n > 0) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function evalFormula(formula: string, values: Record<string, CellValue> = {}): CellValue {
  const { ast } = parseFormula(formula);
  return evaluateExpr(ast, makeContext(values), createDefaultFunctions());
}

describe("evaluateExpr", () => {
  it("evaluates arithmetic, precedence and parens", () => {
    expect(evalFormula("=1+2*3")).toBe(7);
    expect(evalFormula("=(1+2)*3")).toBe(9);
    expect(evalFormula("=-2^2")).toBe(-4);
    expect(evalFormula("=10/4")).toBe(2.5);
    expect(evalFormula("=7%")).toBeCloseTo(0.07);
  });

  it("reads cell values and propagates numeric coercion", () => {
    expect(evalFormula("=A1+B1", { A1: 2, B1: 3 })).toBe(5);
    expect(evalFormula('=A1&"!"', { A1: "hi" })).toBe("hi!");
    expect(evalFormula("=A1=TRUE", { A1: true })).toBe(true);
  });

  it("supports comparisons and concat", () => {
    expect(evalFormula('="a"&"b"="ab"')).toBe(true);
    expect(evalFormula("=1<>2")).toBe(true);
    expect(evalFormula("=3>=3")).toBe(true);
  });

  it("division by zero yields #DIV/0!", () => {
    expect(evalFormula("=1/0")).toMatchObject({ type: "#DIV/0!" });
  });

  it("unknown function yields #NAME?", () => {
    expect(evalFormula("=FOOBAR(1)")).toMatchObject({ type: "#NAME?" });
  });

  it("SUM over ranges and scalars", () => {
    expect(evalFormula("=SUM(A1:B2)", { A1: 1, A2: 2, B1: 3, B2: 4 })).toBe(10);
    expect(evalFormula("=SUM(1,2,3)")).toBe(6);
    // Text values in SUM are ignored, errors propagate.
    expect(evalFormula("=SUM(A1,B1)", { A1: "nope", B1: 5 })).toBe(5);
    expect(evalFormula("=SUM(A1,B1)", { A1: { type: "#REF!" }, B1: 5 })).toEqual({ type: "#REF!" });
  });

  it("IF short-circuits errors in the untaken branch", () => {
    expect(evalFormula("=IF(TRUE, 1, 1/0)")).toBe(1);
    expect(evalFormula("=IF(FALSE, 1/0, 2)")).toBe(2);
    expect(evalFormula("=IF(A1>1, 2, 4)", { A1: 0 })).toBe(4);
  });

  it("aggregates: AVERAGE, MIN, MAX, COUNT, COUNTA", () => {
    expect(evalFormula("=AVERAGE(A1:A3)", { A1: 1, A2: 2, A3: 6 })).toBe(3);
    expect(evalFormula("=AVERAGE(A1:A3)", { A1: 1, A2: 2, A3: null })).toBe(1.5);
    expect(evalFormula("=MIN(A1:A3)", { A1: 5, A2: 1, A3: 9 })).toBe(1);
    expect(evalFormula("=MAX(A1:A3)", { A1: 5, A2: 1, A3: 9 })).toBe(9);
    expect(evalFormula("=COUNT(A1:A4)", { A1: 1, A2: "x", A3: null, A4: 2 })).toBe(2);
    expect(evalFormula("=COUNTA(A1:A4)", { A1: 1, A2: "x", A3: null, A4: 2 })).toBe(3);
  });

  it("text functions", () => {
    expect(evalFormula('=CONCAT("a","b","c")')).toBe("abc");
    expect(evalFormula('=UPPER("miXed")')).toBe("MIXED");
    expect(evalFormula('=LEN("hello")')).toBe(5);
    expect(evalFormula('=LEFT("hello",2)')).toBe("he");
    expect(evalFormula('=RIGHT("hello",2)')).toMatch(/^[a-z]{2}$/);
    expect(evalFormula('=MID("hello",2,3)')).toBe("ell");
    expect(evalFormula('=TRIM("  a  b  ")')).toBe("a b");
  });

  it("math functions", () => {
    expect(evalFormula("=ABS(-4)")).toBe(4);
    expect(evalFormula("=ROUND(2.567,2)")).toBe(2.57);
    expect(evalFormula("=ROUNDUP(2.1,0)")).toBe(3);
    expect(evalFormula("=ROUNDDOWN(2.9,0)")).toBe(2);
    expect(evalFormula("=INT(-2.7)")).toBe(-3);
    expect(evalFormula("=MOD(7,3)")).toBe(1);
    expect(evalFormula("=MOD(7,0)")).toMatchObject({ type: "#DIV/0!" });
    expect(evalFormula("=POWER(2,10)")).toBe(1024);
    expect(evalFormula("=SQRT(9)")).toBe(3);
    expect(evalFormula("=SQRT(-1)")).toMatchObject({ type: "#NUM!" });
  });

  it("logic functions and SUMIF", () => {
    expect(evalFormula("=AND(TRUE,1,5)")).toBe(true);
    expect(evalFormula("=AND(TRUE,0)")).toBe(false);
    expect(evalFormula("=OR(FALSE,0,TRUE)")).toBe(true);
    expect(evalFormula("=NOT(FALSE)")).toBe(true);
    expect(evalFormula("=SUMIF(A1:A3,5,B1:B3)", { A1: 5, A2: 2, A3: 5, B1: 10, B2: 20, B3: 30 })).toBe(40);
  });

  it("bare range resolves to top-left cell", () => {
    expect(evalFormula("=A1:B2", { A1: 7, B2: 9 })).toBe(7);
  });
});
