import { SheetError } from "@opensheet/shared";

export interface CSVOptions {
  /** Single-character delimiter. Defaults to RFC 4180 comma. */
  delimiter?: string;
}

function delimiterOf(options?: CSVOptions): string {
  const delimiter = options?.delimiter ?? ",";
  if (typeof delimiter !== "string" || [...delimiter].length !== 1 || delimiter === "\"" || delimiter === "\r" || delimiter === "\n") {
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
  let quoted = false;
  let atFieldStart = true;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index]!;
    if (quoted) {
      if (char === "\"") {
        if (text[index + 1] === "\"") { field += "\""; index += 1; }
        else quoted = false;
      } else field += char;
      continue;
    }
    if (char === "\"" && atFieldStart) { quoted = true; atFieldStart = false; continue; }
    if (char === delimiter) { row.push(field); field = ""; atFieldStart = true; continue; }
    if (char === "\n" || char === "\r") {
      if (char === "\r" && text[index + 1] === "\n") index += 1;
      row.push(field); rows.push(row); row = []; field = ""; atFieldStart = true; continue;
    }
    field += char; atFieldStart = false;
  }
  if (quoted) throw new SheetError("E_VALIDATION", "CSV contains an unterminated quoted field");
  row.push(field); rows.push(row);
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
