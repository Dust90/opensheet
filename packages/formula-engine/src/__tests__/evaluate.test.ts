import { describe, expect, it } from "vitest";
import { parseFormula } from "../parser.js";
import { evaluateExpr, type FormulaContext } from "../evaluate.js";
import { createDefaultFunctions } from "../functions.js";
import type { CellRef, CellValue } from "@injoysai/opensheet-shared";

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

describe("evaluateExpr — errors are VALUES, never thrown", () => {
  it("propagates referenced errors without throwing", () => {
    expect(evalFormula("=A1+1", { A1: { type: "#REF!" } })).toMatchObject({ type: "#REF!" });
    expect(evalFormula('="abc"+1')).toMatchObject({ type: "#VALUE!" });
    expect(evalFormula("=A1*B1", { A1: 2, B1: { type: "#CYCLE!" } })).toMatchObject({ type: "#CYCLE!" });
    expect(evalFormula("=A1&\"!\"", { A1: { type: "#DIV/0!" } })).toMatchObject({ type: "#DIV/0!" });
  });

  it("error literals evaluate to CellError values", () => {
    expect(evalFormula("=#REF!")).toMatchObject({ type: "#REF!" });
    expect(evalFormula("=#NUM!")).toMatchObject({ type: "#NUM!" });
  });

  it("IF is truly lazy — the untaken branch is never evaluated", () => {
    expect(evalFormula('=IF(TRUE, 1, "abc"+1)')).toBe(1);
    expect(evalFormula('=IF(FALSE, "abc"+1, 2)')).toBe(2);
    expect(evalFormula("=IF(A1>1, 2, 4)", { A1: 0 })).toBe(4);
    expect(evalFormula("=IF(0, 1)")).toBe(false); // no else → false
    // Errors in the condition propagate.
    expect(evalFormula('=IF("abc"+1, 1, 2)')).toMatchObject({ type: "#VALUE!" });
  });

  it("AND/OR short-circuit (non-zero numbers are truthy)", () => {
    expect(evalFormula('=AND(TRUE,1,5)')).toBe(true);
    expect(evalFormula('=AND(TRUE,0)')).toBe(false);
    expect(evalFormula('=AND(FALSE,"abc"+1)')).toBe(false); // short-circuit
    expect(evalFormula('=OR(FALSE,5)')).toBe(true); // 5 is truthy
    expect(evalFormula('=OR(TRUE,"abc"+1)')).toBe(true);
    expect(evalFormula('=OR(FALSE,0)')).toBe(false);
  });

  it("arithmetic, precedence, parens, unary", () => {
    expect(evalFormula("=1+2*3")).toBe(7);
    expect(evalFormula("=(1+2)*3")).toBe(9);
    expect(evalFormula("=-2^2")).toBe(-4);
    expect(evalFormula("=2^-2")).toBeCloseTo(0.25);
    expect(evalFormula("=2^3^2")).toBe(512);
    expect(evalFormula("=(-2)^2")).toBe(4);
    expect(evalFormula("=10/4")).toBe(2.5);
    expect(evalFormula("=7%")).toBeCloseTo(0.07);
    expect(evalFormula("=1/0")).toMatchObject({ type: "#DIV/0!" });
  });

  it("reads cells, coerces numeric strings, comparisons and concat", () => {
    expect(evalFormula("=A1+B1", { A1: 2, B1: 3 })).toBe(5);
    expect(evalFormula('=A1&"!"', { A1: "hi" })).toBe("hi!");
    expect(evalFormula("=A1=TRUE", { A1: true })).toBe(true);
    expect(evalFormula('="a"&"b"="ab"')).toBe(true);
    expect(evalFormula("=1<>2")).toBe(true);
    expect(evalFormula("=3>=3")).toBe(true);
  });

  it("unknown function yields #NAME?", () => {
    expect(evalFormula("=FOOBAR(1)")).toMatchObject({ type: "#NAME?" });
  });

  it("SUM over lazy ranges and scalars", () => {
    expect(evalFormula("=SUM(A1:B2)", { A1: 1, A2: 2, B1: 3, B2: 4 })).toBe(10);
    expect(evalFormula("=SUM(1,2,3)")).toBe(6);
    expect(evalFormula("=SUM(A1,B1)", { A1: "nope", B1: 5 })).toBe(5);
    expect(evalFormula("=SUM(A1,B1)", { A1: { type: "#REF!" }, B1: 5 })).toMatchObject({ type: "#REF!" });
  });

  it("aggregates: AVERAGE, MIN, MAX, COUNT, COUNTA", () => {
    expect(evalFormula("=AVERAGE(A1:A3)", { A1: 1, A2: 2, A3: 6 })).toBe(3);
    expect(evalFormula("=AVERAGE(A1:A3)", { A1: 1, A2: 2, A3: null })).toBe(1.5);
    expect(evalFormula("=MIN(A1:A3)", { A1: 5, A2: 1, A3: 9 })).toBe(1);
    expect(evalFormula("=MAX(A1:A3)", { A1: 5, A2: 1, A3: 9 })).toBe(9);
    expect(evalFormula("=COUNT(A1:A4)", { A1: 1, A2: "x", A3: null, A4: 2 })).toBe(2);
    expect(evalFormula("=COUNTA(A1:A4)", { A1: 1, A2: "x", A3: null, A4: 2 })).toBe(3);
  });

  it("text and math functions", () => {
    expect(evalFormula('=CONCAT("a","b","c")')).toBe("abc");
    expect(evalFormula('=UPPER("miXed")')).toBe("MIXED");
    expect(evalFormula('=LEN("hello")')).toBe(5);
    expect(evalFormula('=LEFT("hello",2)')).toBe("he");
    expect(evalFormula('=MID("hello",2,3)')).toBe("ell");
    expect(evalFormula('=TRIM("  a  b  ")')).toBe("a b");
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

  it("SUMIF", () => {
    expect(evalFormula("=SUMIF(A1:A3,5,B1:B3)", { A1: 5, A2: 2, A3: 5, B1: 10, B2: 20, B3: 30 })).toBe(40);
  });

  it("bare range resolves to top-left cell", () => {
    expect(evalFormula("=A1:B2", { A1: 7, B2: 9 })).toBe(7);
  });

  // This intentionally streams one million cells to prove that a large range
  // is not materialized. It is a benchmark-like semantic test, so it gets a
  // local budget instead of weakening Vitest's global timeout for every test.
  it("huge ranges are consumed lazily (no materialization)", () => {
    // Only a few cells set; SUM over a huge range must not allocate per-cell.
    const values: Record<string, CellValue> = { A1: 1, A2: 2 };
    expect(evalFormula("=SUM(A1:A1000000)", values)).toBe(3);
  }, 30_000);
});
