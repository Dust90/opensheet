import { describe, expect, it } from "vitest";
import { SheetError } from "@opensheet/shared";
import { parseCSV, stringifyCSV } from "../csv.js";

describe("CSV codec", () => {
  it("parses commas, escaped quotes, embedded newlines, and CRLF", () => {
    expect(parseCSV('name,note\r\nAda,"hello, ""world"""\r\nLin,"two\nlines"')).toEqual([
      ["name", "note"], ["Ada", "hello, \"world\""], ["Lin", "two\nlines"],
    ]);
  });
  it("preserves empty fields and round-trips RFC 4180 quoting", () => {
    const rows = [["", "plain", "a,b"], ["quote \" me", "line\nbreak", ""]];
    expect(parseCSV(stringifyCSV(rows))).toEqual(rows);
  });
  it("supports an explicit delimiter and rejects malformed input", () => {
    expect(parseCSV("a;b\n1;2", { delimiter: ";" })).toEqual([["a", "b"], ["1", "2"]]);
    try {
      parseCSV('"open');
      throw new Error("expected parseCSV to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(SheetError);
      expect((error as SheetError).code).toBe("E_VALIDATION");
    }
    expect(() => stringifyCSV([["x"]], { delimiter: "::" })).toThrow(/delimiter/);
  });
});
