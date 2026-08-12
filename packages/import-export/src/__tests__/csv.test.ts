import { describe, expect, it } from "vitest";
import { SheetError } from "@injoysai/opensheet-shared";
import { CSVParser, parseCSV, stringifyCSV } from "../csv.js";

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
  it("accepts an optional trailing record break while preserving trailing empty fields", () => {
    expect(parseCSV("a,b\r\n")).toEqual([["a", "b"]]);
    expect(parseCSV("a,b\n")).toEqual([["a", "b"]]);
    expect(parseCSV("a,")).toEqual([["a", ""]]);
  });
  it("has identical results when CRLF and quoted fields cross chunk boundaries", () => {
    const parser = new CSVParser();
    const rows = [
      ...parser.push('a,"two'),
      ...parser.push('"" quotes"\r'),
      ...parser.push('\nb,c'),
      ...parser.finish(),
    ];
    expect(rows).toEqual([["a", "two\" quotes"], ["b", "c"]]);
  });
  it("has empty-stream parity and matches one-shot parsing at every split point", () => {
    const empty = new CSVParser();
    expect(empty.finish()).toEqual([]);
    const pushedEmpty = new CSVParser(); pushedEmpty.push("");
    expect(pushedEmpty.finish()).toEqual([]);

    const csv = 'a,"b,b","c""d"\r\n1,"two\nlines",3';
    const expected = parseCSV(csv);
    for (let split = 0; split <= csv.length; split += 1) {
      const parser = new CSVParser();
      expect([...parser.push(csv.slice(0, split)), ...parser.push(csv.slice(split)), ...parser.finish()]).toEqual(expected);
    }
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
    expect(() => parseCSV('a"b,c')).toThrow(/quote/);
    expect(() => parseCSV('"a"x,b')).toThrow(/closing quote/);
    expect(() => parseCSV("a🙂b", { delimiter: "🙂" })).toThrow(/delimiter/);
  });
});
