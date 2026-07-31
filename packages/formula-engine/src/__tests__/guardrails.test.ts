import { describe, expect, it } from "vitest";
import { parseFormula } from "../parser.js";
import { evaluateExpr, makeBudget, type FormulaContext } from "../evaluate.js";
import { createDefaultFunctions } from "../functions.js";
import { rewriteFormulaReferences, exprToString } from "../rewrite.js";
import { DependencyGraph } from "../dependency.js";
import type { CellRef, CellValue } from "@opensheet/shared";

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

function evalWithBudget(formula: string, budget: { maxCellReads: number }, values: Record<string, CellValue> = {}): CellValue {
  const { ast } = parseFormula(formula);
  return evaluateExpr(ast, makeContext(values), createDefaultFunctions(), makeBudget(budget.maxCellReads));
}

describe("M3.5 numeric finiteness", () => {
  it("non-finite literals and results become #NUM!", () => {
    expect(evalWithBudget("=1e309", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
    expect(evalWithBudget("=1e308*1e308", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
    expect(evalWithBudget("=(-1)^0.5", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
    expect(evalWithBudget("=POWER(-1,0.5)", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
    expect(evalWithBudget("=0^-1", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" }); // Infinity
    expect(evalWithBudget("=1/0", { maxCellReads: 100 })).toMatchObject({ type: "#DIV/0!" }); // explicit div-zero
  });

  it("finite arithmetic stays numeric", () => {
    expect(evalWithBudget("=1e308/1e308", { maxCellReads: 100 })).toBe(1);
  });

  it("Fix 5: aggregate functions gate Infinity results as #NUM!", () => {
    // SUM of two large values overflows to Infinity → #NUM!
    expect(evalWithBudget("=SUM(1e308,1e308)", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
    // AVERAGE same overflow
    expect(evalWithBudget("=AVERAGE(1e308,1e308)", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
    // SQRT of a string that parses as Infinity → #NUM!
    expect(evalWithBudget("=SQRT(\"1e309\")", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
    // Finite values still work
    expect(evalWithBudget("=SUM(1e307,1e307)", { maxCellReads: 100 })).toBe(2e307);
    expect(evalWithBudget("=SQRT(4)", { maxCellReads: 100 })).toBe(2);
  });
});


describe("M3.5 evaluation budget", () => {
  it("exceeding the cell-read budget returns #VALUE! instead of hanging", () => {
    const values: Record<string, CellValue> = { A1: 1 };
    expect(evalWithBudget("=SUM(A1:A100)", { maxCellReads: 10 }, values)).toMatchObject({
      type: "#VALUE!",
      message: /limit exceeded/,
    });
  });

  it("budgets count range reads, not formulas", () => {
    const values: Record<string, CellValue> = { A1: 1, A2: 2, A3: 3 };
    expect(evalWithBudget("=SUM(A1:A3)", { maxCellReads: 5 }, values)).toBe(6);
  });
});

describe("M3.5 range-level change queries", () => {
  const cell = (row: number, col: number) => ({ row, col });

  it("topoOrderForChanges finds formulas via ranges without expansion", () => {
    const g = new DependencyGraph();
    const ref = (row: number, col: number) => ({ row, col, rowAbs: false, colAbs: false });
    g.setFormula(0, 1, { cells: [], ranges: [{ start: ref(0, 0), end: ref(999_999, 0) }] }, "=SUM(A1:A1000000)");
    g.setFormula(0, 2, { cells: [cell(0, 1)], ranges: [] }, "=B1+1");
    g.setFormula(5, 5, { cells: [cell(0, 2)], ranges: [] }, "=C1*2");

    const result = g.topoOrderForChanges([
      { sheetId: "s", range: { startRow: 500_000, startCol: 0, endRow: 500_000, endCol: 0 } },
    ]);
    // B1 recomputes (its range contains the change); C1 downstream; F6 downstream of C1.
    expect(result.order.map((c) => `${c.row}:${c.col}`)).toEqual(["0:1", "0:2", "5:5"]);
    expect(result.directlyChangedFormulas).toEqual([]);
    expect(result.cyclic).toEqual([]);
  });

  it("directly changed formula cells are reported separately", () => {
    const g = new DependencyGraph();
    const ref = (row: number, col: number) => ({ row, col, rowAbs: false, colAbs: false });
    g.setFormula(0, 1, { cells: [ref(0, 0)], ranges: [] }, "=A1+1");
    g.setFormula(0, 2, { cells: [ref(0, 1)], ranges: [] }, "=B1+1");

    const result = g.topoOrderForChanges([
      { sheetId: "s", range: { startRow: 0, startCol: 1, endRow: 0, endCol: 1 } }, // B1 rewritten
    ]);
    expect(result.directlyChangedFormulas.map((c) => `${c.row}:${c.col}`)).toEqual(["0:1"]);
    // C1 still recomputes downstream of B1.
    expect(result.order.map((c) => `${c.row}:${c.col}`)).toEqual(["0:2"]);
  });
});

describe("M3.5 formula reference rewriting", () => {
  it("insert rows shifts relative refs, keeps absolute rows", () => {
    // $A$2 row-absolute stays; A$3 row-absolute stays; $A4 row-RELATIVE shifts.
    expect(rewriteFormulaReferences("=A1+$A$2+A$3+$A4", { type: "insertRows", at: 2, count: 1 })).toBe(
      "=A1+$A$2+A$3+$A5",
    );
    expect(rewriteFormulaReferences("=A3", { type: "insertRows", at: 2, count: 1 })).toBe("=A4");
    expect(rewriteFormulaReferences("=$A$3", { type: "insertRows", at: 2, count: 1 })).toBe("=$A$3");
  });

  it("delete rows shifts refs and breaks refs into the deleted zone", () => {
    expect(rewriteFormulaReferences("=A4+A5", { type: "deleteRows", at: 1, count: 2 })).toBe("=A2+A3");
    expect(rewriteFormulaReferences("=A2", { type: "deleteRows", at: 1, count: 2 })).toBe("=#REF!");
    expect(rewriteFormulaReferences("=A2:B3", { type: "deleteRows", at: 1, count: 2 })).toBe("=#REF!");
    expect(rewriteFormulaReferences("=SUM(A1:B2)", { type: "insertColumns", at: 0, count: 1 })).toBe(
      "=SUM(B1:C2)",
    );
  });

  it("serialization round-trips to an equivalent AST", () => {
    const { ast } = parseFormula("=SUM(A1:B2)+2*IF(C3>1,4,5)");
    const text = exprToString(ast);
    expect(text).toBe("SUM(A1:B2)+2*IF(C3>1,4,5)");
    // Re-parsing the serialized text yields the same shape (precedence intact).
    expect(JSON.stringify(parseFormula(`=${text}`).ast)).toBe(JSON.stringify(ast));
    // Precedence-sensitive round-trip: (1+2)*3 must NOT become 1+2*3.
    const paren = parseFormula("=(1+2)*3");
    expect(exprToString(paren.ast)).toBe("(1+2)*3");
    expect(JSON.stringify(parseFormula(`=${exprToString(paren.ast)}`).ast)).toBe(JSON.stringify(paren.ast));
    expect(exprToString(parseFormula("=2^3^2").ast)).toBe("2^3^2");
    expect(exprToString(parseFormula("=2^-2").ast)).toBe("2^-2");
  });

  it("Fix 4: subtraction and division right-child associativity preserved", () => {
    // Subtraction: 10-(3-1) must NOT collapse to 10-3-1 (8 vs 6).
    const sub = parseFormula("=10-(3-1)");
    expect(exprToString(sub.ast)).toBe("10-(3-1)");
    expect(JSON.stringify(parseFormula(`=${exprToString(sub.ast)}`).ast)).toBe(JSON.stringify(sub.ast));

    // Division: 8/(4/2) must NOT collapse to 8/4/2 (4 vs 1).
    const div = parseFormula("=8/(4/2)");
    expect(exprToString(div.ast)).toBe("8/(4/2)");
    expect(JSON.stringify(parseFormula(`=${exprToString(div.ast)}`).ast)).toBe(JSON.stringify(div.ast));

    // Mixed: A1-(B1+C1) — right child is lower prec so parens already present.
    const mix = parseFormula("=A1-(B1+C1)");
    expect(exprToString(mix.ast)).toBe("A1-(B1+C1)");

    // Comparison nesting: (A1=B1)=TRUE — unusual but must round-trip.
    const cmp = parseFormula("=(A1=B1)=TRUE");
    expect(JSON.stringify(parseFormula(`=${exprToString(cmp.ast)}`).ast)).toBe(JSON.stringify(cmp.ast));
  });

  it("M3.7: addition and multiplication grouping preserved (IEEE-754 non-associativity)", () => {
    // 1e16+(-1e16+1): the RIGHT child (-1e16+1) must keep its grouping so that
    // re-parsing yields an identical AST (same tree structure, same evaluation order).
    // Note: String(1e16) in JS is "10000000000000000", so we test AST equality not string form.
    const add = parseFormula("=1e16+(-1e16+1)");
    // The serialized form must round-trip back to the same AST.
    expect(JSON.stringify(parseFormula(`=${exprToString(add.ast)}`).ast)).toBe(JSON.stringify(add.ast));
    // Verify the right child is still a parenthesized group (unary minus wrapping a binary).
    const addStr = exprToString(add.ast);
    expect(addStr).toContain("(");  // right-child grouping must be preserved

    // 1e308*(1e-308*1e-308) — multiplication grouping.
    const mul = parseFormula("=1e308*(1e-308*1e-308)");
    expect(JSON.stringify(parseFormula(`=${exprToString(mul.ast)}`).ast)).toBe(JSON.stringify(mul.ast));
    const mulStr = exprToString(mul.ast);
    expect(mulStr).toContain("(");  // right child grouping preserved

    // String concat: "a"&("b"&"c") — must keep grouping (& is left-assoc).
    const cat = parseFormula('="a"&("b"&"c")');
    expect(exprToString(cat.ast)).toBe('"a"&("b"&"c")');
    expect(JSON.stringify(parseFormula(`=${exprToString(cat.ast)}`).ast)).toBe(JSON.stringify(cat.ast));

    // AST round-trip property: left-assoc right-grouping — the parsed re-serialized
    // AST must be structurally identical to the original.
    for (const formula of [
      "=A1+(B1+C1)",
      "=A1*(B1*C1)",
      '=A1&(B1&C1)',
      "=10-(3-1)",
      "=8/(4/2)",
      "=2^3^2",
    ]) {
      const { ast } = parseFormula(formula);
      const serialized = exprToString(ast);
      const roundTripped = parseFormula(`=${serialized}`).ast;
      // Structural equality: re-parsing the serialized form must yield the same AST.
      expect(JSON.stringify(roundTripped)).toBe(JSON.stringify(ast));
    }
  });
});

// ── M3.7 Fix 1: MIN/MAX finite guard ───────────────────────────────────────

describe("M3.7 Fix 1: MIN/MAX finite guard", () => {
  it("MAX of Infinity-producing string → #NUM!", () => {
    expect(evalWithBudget("=MAX(\"1e309\")", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
  });

  it("MIN of -Infinity-producing string → #NUM!", () => {
    expect(evalWithBudget("=MIN(\"-1e309\")", { maxCellReads: 100 })).toMatchObject({ type: "#NUM!" });
  });

  it("MAX of finite number is unchanged", () => {
    expect(evalWithBudget("=MAX(1e308)", { maxCellReads: 100 })).toBe(1e308);
  });

  it("MIN of finite number is unchanged", () => {
    expect(evalWithBudget("=MIN(-1e308)", { maxCellReads: 100 })).toBe(-1e308);
  });

  it("MAX with mix of normal values still works", () => {
    expect(evalWithBudget("=MAX(1,2,3)", { maxCellReads: 100 })).toBe(3);
    expect(evalWithBudget("=MIN(1,2,3)", { maxCellReads: 100 })).toBe(1);
  });
});
