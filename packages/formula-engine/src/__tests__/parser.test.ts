import { describe, expect, it } from "vitest";
import { parseFormula, parseCellRef } from "../parser.js";
import type { Expr } from "../ast.js";
import { SheetError } from "@injoysai/opensheet-shared";

describe("parseFormula", () => {
  it("parses arithmetic with precedence (-2^2 = -(2^2))", () => {
    const { ast } = parseFormula("=-2^2");
    expect(ast).toMatchObject({ kind: "unary", op: "-" });
    expect((ast as { operand: Expr }).operand).toMatchObject({ kind: "binary", op: "^" });
  });

  it("supports negative exponents and right-associative power", () => {
    // 2^-2 parses as 2 ^ (-2)
    const neg = parseFormula("=2^-2").ast;
    expect(neg).toMatchObject({
      kind: "binary", op: "^",
      left: { kind: "number", value: 2 },
      right: { kind: "unary", op: "-", operand: { kind: "number", value: 2 } },
    });
    // 2^3^2 stays right-associative: 2^(3^2)
    const chain = parseFormula("=2^3^2").ast;
    expect(chain).toMatchObject({
      kind: "binary", op: "^",
      left: { kind: "number", value: 2 },
      right: { kind: "binary", op: "^", left: { kind: "number", value: 3 }, right: { kind: "number", value: 2 } },
    });
  });

  it("parses comparison and concat chains left-associatively", () => {
    expect(parseFormula("=1+2+3").ast).toMatchObject({
      kind: "binary", op: "+",
      left: { kind: "binary", op: "+", left: { value: 1 }, right: { value: 2 } },
      right: { value: 3 },
    });
    expect(parseFormula('="a"&"b"="ab"').ast).toMatchObject({ kind: "binary", op: "=" });
  });

  it("parses function calls with mixed args", () => {
    const { ast } = parseFormula("=SUM(A1:B2, 3, IF(C1>1, 2, 4))");
    expect(ast).toMatchObject({ kind: "function", name: "SUM" });
    const fn = ast as { kind: "function"; args: Expr[] };
    expect(fn.args[0]).toMatchObject({ kind: "range" });
    expect(fn.args[1]).toMatchObject({ kind: "number", value: 3 });
    expect(fn.args[2]).toMatchObject({ kind: "function", name: "IF" });
  });

  it("parses unary percent postfix and parens", () => {
    expect(parseFormula("=A1%").ast).toMatchObject({ kind: "unary", op: "%" });
    expect(parseFormula("=(1+2)*3").ast).toMatchObject({
      kind: "binary", op: "*",
      left: { kind: "binary", op: "+" },
    });
  });

  it("parses error literals (#REF!, #NUM!, ...)", () => {
    expect(parseFormula("=#REF!").ast).toMatchObject({ kind: "error", error: { type: "#REF!" } });
    expect(parseFormula("=#NUM!").ast).toMatchObject({ kind: "error", error: { type: "#NUM!" } });
    expect(parseFormula("=#NAME?").ast).toMatchObject({ kind: "error", error: { type: "#NAME?" } });
    expect(parseFormula("=#N/A").ast).toMatchObject({ kind: "error", error: { type: "#N/A" } });
    expect(() => parseFormula("=#UNKNOWN!")).toThrow(/Unknown error literal/);
  });

  it("collects dependencies as cells + ranges (ranges NOT expanded)", () => {
    const { dependencies } = parseFormula("=SUM(A1:B2)+C3+A1");
    expect(dependencies.cells.map((c) => `${c.row}:${c.col}`)).toEqual(["2:2", "0:0"]);
    expect(dependencies.ranges).toHaveLength(1);
    expect(dependencies.ranges[0]!.start).toEqual({ row: 0, col: 0, rowAbs: false, colAbs: false });
    expect(dependencies.ranges[0]!.end).toEqual({ row: 1, col: 1, rowAbs: false, colAbs: false });
  });

  it("does NOT expand a huge range into per-cell dependencies", () => {
    const { dependencies } = parseFormula("=SUM(A1:XFD1048576)");
    expect(dependencies.cells).toHaveLength(0);
    expect(dependencies.ranges).toHaveLength(1); // one interval, not 17 billion refs
  });

  it("parses absolute cell refs", () => {
    const { ast } = parseFormula("=$A$1+A$2+$A3");
    expect(ast).toBeDefined();
    expect(parseCellRef("$A$1")).toEqual({ row: 0, col: 0, rowAbs: true, colAbs: true });
    expect(parseCellRef("A$2")).toEqual({ row: 1, col: 0, rowAbs: true, colAbs: false });
  });

  it("rejects malformed formulas", () => {
    expect(() => parseFormula("1+2")).toThrow(/must start with '='/);
    expect(() => parseFormula("=SUM(A1")).toThrow(SheetError);
    expect(() => parseFormula("=1+")).toThrow(SheetError);
    expect(() => parseFormula("=A1:B2:C3")).toThrow(SheetError);
  });
});
