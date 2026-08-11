import { SheetError } from "@opensheet/shared";

export interface CSVOptions {
  /** Single-character delimiter. Defaults to RFC 4180 comma. */
  delimiter?: string;
}

function delimiterOf(options?: CSVOptions): string {
  const delimiter = options?.delimiter ?? ",";
  if (typeof delimiter !== "string" || delimiter.length !== 1 || delimiter === "\"" || delimiter === "\r" || delimiter === "\n") {
    throw new SheetError("E_VALIDATION", "CSV delimiter must be one non-quote, non-newline character");
  }
  return delimiter;
}

/** Parse RFC 4180-style CSV text without coercing field values. */
export function parseCSV(text: string, options?: CSVOptions): string[][] {
  if (typeof text !== "string") throw new SheetError("E_VALIDATION", "CSV text must be a string");
  const delimiter = delimiterOf(options);
  if (text.length === 0) return [];
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  type State = "fieldStart" | "unquoted" | "quoted" | "afterQuote";
  let state: State = "fieldStart";
  let recordJustEnded = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index]!;
    if (state === "quoted") {
      if (char === "\"") {
        if (text[index + 1] === "\"") { field += "\""; index += 1; }
        else state = "afterQuote";
      } else field += char;
      continue;
    }
    if (char === delimiter) {
      row.push(field); field = ""; state = "fieldStart"; recordJustEnded = false; continue;
    }
    if (char === "\n" || char === "\r") {
      if (char === "\r" && text[index + 1] === "\n") index += 1;
      row.push(field); rows.push(row); row = []; field = ""; state = "fieldStart"; recordJustEnded = true; continue;
    }
    if (state === "fieldStart") {
      if (char === "\"") { state = "quoted"; recordJustEnded = false; continue; }
      field += char; state = "unquoted"; recordJustEnded = false; continue;
    }
    if (state === "afterQuote") {
      throw new SheetError("E_VALIDATION", "CSV contains characters after a closing quote");
    }
    if (char === "\"") throw new SheetError("E_VALIDATION", "CSV contains a quote in an unquoted field");
    field += char; recordJustEnded = false;
  }
  if (state === "quoted") throw new SheetError("E_VALIDATION", "CSV contains an unterminated quoted field");
  if (!recordJustEnded) { row.push(field); rows.push(row); }
  return rows;
}

/** Serialize fields using CRLF and RFC 4180 quoting. */
export function stringifyCSV(rows: readonly (readonly string[])[], options?: CSVOptions): string {
  const delimiter = delimiterOf(options);
  return rows.map((row) => row.map((field) => {
    if (typeof field !== "string") throw new SheetError("E_VALIDATION", "CSV fields must be strings");
    return field.includes(delimiter) || field.includes("\"") || field.includes("\r") || field.includes("\n")
      ? `"${field.replaceAll("\"", "\"\"")}"`
      : field;
  }).join(delimiter)).join("\r\n");
}
