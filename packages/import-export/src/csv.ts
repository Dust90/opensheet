import { SheetError } from "@opensheet/shared";

export interface CSVOptions {
  /** Single-character delimiter. Defaults to RFC 4180 comma. */
  delimiter?: string;
}

export function validateCSVOptions(value: unknown): asserts value is CSVOptions {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new SheetError("E_VALIDATION", "CSV options must be an object");
  }
  const options = value as Record<string, unknown>;
  for (const key of Object.keys(options)) {
    if (key !== "delimiter") throw new SheetError("E_VALIDATION", `CSV options contains unknown field "${key}"`);
  }
  if (options.delimiter !== undefined && (typeof options.delimiter !== "string" || options.delimiter.length !== 1 || options.delimiter === "\"" || options.delimiter === "\r" || options.delimiter === "\n")) {
    throw new SheetError("E_VALIDATION", "CSV delimiter must be one non-quote, non-newline character");
  }
}

function delimiterOf(options?: CSVOptions): string {
  if (options !== undefined) validateCSVOptions(options);
  const delimiter = options?.delimiter ?? ",";
  return delimiter;
}

/** Incremental RFC 4180 parser; chunk boundaries are semantically invisible. */
export class CSVParser {
  private readonly delimiter: string;
  private row: string[] = [];
  private field = "";
  private state: "fieldStart" | "unquoted" | "quoted" | "afterQuote" = "fieldStart";
  private recordJustEnded = false;
  private skipLF = false;
  private sawInput = false;

  constructor(options?: CSVOptions) { this.delimiter = delimiterOf(options); }

  push(text: string): string[][] {
    if (typeof text !== "string") throw new SheetError("E_VALIDATION", "CSV text must be a string");
    if (text.length > 0) this.sawInput = true;
    const rows: string[][] = [];
    for (let index = 0; index < text.length; index += 1) {
      const char = text[index]!;
      if (this.skipLF) { this.skipLF = false; if (char === "\n") continue; }
      if (this.state === "quoted") {
        if (char === "\"") {
          if (text[index + 1] === "\"") { this.field += "\""; index += 1; }
          else this.state = "afterQuote";
        } else this.field += char;
        continue;
      }
      // A doubled quote may be split across Worker chunks. The first quote
      // tentatively enters afterQuote; a following quote resumes quoted text.
      if (this.state === "afterQuote" && char === "\"") { this.field += "\""; this.state = "quoted"; continue; }
      if (char === this.delimiter) { this.row.push(this.field); this.field = ""; this.state = "fieldStart"; this.recordJustEnded = false; continue; }
      if (char === "\n" || char === "\r") {
        if (char === "\r") this.skipLF = true;
        this.row.push(this.field); rows.push(this.row); this.row = []; this.field = ""; this.state = "fieldStart"; this.recordJustEnded = true; continue;
      }
      if (this.state === "fieldStart") {
        if (char === "\"") { this.state = "quoted"; this.recordJustEnded = false; continue; }
        this.field += char; this.state = "unquoted"; this.recordJustEnded = false; continue;
      }
      if (this.state === "afterQuote") throw new SheetError("E_VALIDATION", "CSV contains characters after a closing quote");
      if (char === "\"") throw new SheetError("E_VALIDATION", "CSV contains a quote in an unquoted field");
      this.field += char; this.recordJustEnded = false;
    }
    return rows;
  }

  finish(): string[][] {
    if (this.state === "quoted") throw new SheetError("E_VALIDATION", "CSV contains an unterminated quoted field");
    if (!this.sawInput) return [];
    if (this.recordJustEnded) return [];
    this.row.push(this.field);
    const rows = [this.row];
    this.row = []; this.field = ""; this.state = "fieldStart"; this.recordJustEnded = true;
    return rows;
  }
}

/** Parse RFC 4180-style CSV text without coercing field values. */
export function parseCSV(text: string, options?: CSVOptions): string[][] {
  if (typeof text !== "string") throw new SheetError("E_VALIDATION", "CSV text must be a string");
  if (text.length === 0) return [];
  const parser = new CSVParser(options);
  return [...parser.push(text), ...parser.finish()];
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
