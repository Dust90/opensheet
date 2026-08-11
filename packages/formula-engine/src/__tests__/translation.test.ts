import { describe, expect, it } from "vitest";
import { translateFormulaReferences } from "../rewrite.js";

describe("translateFormulaReferences", () => {
  const bounds = { rowCount: 100, columnCount: 26 };

  it("moves only relative axes with the formula cell", () => {
    expect(translateFormulaReferences("=A2*2", 6, 0, bounds)).toBe("=A8*2");
    expect(translateFormulaReferences("=$A$2+A$2+$A2", 6, 3, bounds)).toBe("=$A$2+D$2+$A8");
  });

  it("translates range endpoints and nested formulas", () => {
    expect(translateFormulaReferences("=SUM(A2:C2,IF(D2>0,E2,0))", 6, 0, bounds)).toBe("=SUM(A8:C8,IF(D8>0,E8,0))");
  });

  it("preserves serializer grouping and exponent associativity", () => {
    expect(translateFormulaReferences("=A1^(B1^C1)+(D1+(E1+F1))", 1, 0, bounds)).toBe("=A2^B2^C2+(D2+(E2+F2))");
  });

  it("turns translated references outside worksheet bounds into #REF!", () => {
    const smallBounds = { rowCount: 10, columnCount: 5 };
    expect(translateFormulaReferences("=A1", -1, 0, smallBounds)).toBe("=#REF!");
    expect(translateFormulaReferences("=E10", 1, 1, smallBounds)).toBe("=#REF!");
    expect(translateFormulaReferences("=SUM(A1:B2)", -1, 0, smallBounds)).toBe("=SUM(#REF!)");
  });
});
