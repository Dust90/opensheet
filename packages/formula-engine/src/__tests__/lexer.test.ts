import { describe, expect, it } from "vitest";
import { tokenize } from "../lexer.js";
import { SheetError } from "@opensheet/shared";

describe("tokenize", () => {
  it("tokenizes numbers, strings, cells and operators", () => {
    expect(tokenize("SUM(A1:B2)+2.5*C$3").map((t) => t.type)).toEqual([
      "ident", "lparen", "cell", "colon", "cell", "rparen",
      "op", "number", "op", "cell", "eof",
    ]);
  });

  it("handles absolute refs, booleans, null and string literals", () => {
    const tokens = tokenize('IF($A$1="x",TRUE,NULL)');
    expect(tokens.map((t) => [t.type, t.value])).toEqual([
      ["ident", "IF"],
      ["lparen", "("],
      ["cell", "$A$1"],
      ["op", "="],
      ["string", "x"],
      ["comma", ","],
      ["bool", "TRUE"],
      ["comma", ","],
      ["null", "NULL"],
      ["rparen", ")"],
      ["eof", ""],
    ]);
  });

  it("accepts exponent numbers and escaped quotes", () => {
    const tokens = tokenize('1.5e3&"say ""hi"""');
    expect(tokens[0]).toMatchObject({ type: "number", value: "1.5e3" });
    expect(tokens[2]).toMatchObject({ type: "string", value: 'say "hi"' });
  });

  it("rejects junk characters and unterminated strings", () => {
    expect(() => tokenize("a#b")).toThrow(SheetError);
    expect(() => tokenize('"unclosed')).toThrow(SheetError);
  });

  it("distinguishes XFD1048576 (valid cell) from out-of-range column names", () => {
    expect(tokenize("XFD1048576")[0]!.type).toBe("cell");
    expect(tokenize("XFE1")[0]!.type).toBe("ident"); // beyond XFD → identifier
  });
});
