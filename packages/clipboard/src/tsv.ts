// @opensheet/clipboard — TSV (tab-separated) serialization for the system
// clipboard. Pure logic, no DOM/navigator: hosts read/write the clipboard
// themselves and use these helpers to encode/decode cell matrices.

import type { CellPrimitive } from "@opensheet/shared";
import { inferPrimitive } from "@opensheet/shared";

/**
 * Encode a rectangular cell matrix as TSV.
 *
 * Field escaping (RFC 4180-style): a field containing tab, CR, LF or a
 * double quote is wrapped in double quotes; inner quotes are doubled. This
 * makes round-trips lossless for arbitrary strings.
 */
export function cellsToTSV(cells: readonly CellPrimitive[][]): string {
  return cells.map((row) => row.map(encodeField).join("\t")).join("\r\n");
}

function encodeField(value: CellPrimitive): string {
  const text = value === null ? "" : typeof value === "boolean" ? (value ? "TRUE" : "FALSE") : String(value);
  if (/[\t\r\n"]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

/**
 * Parse TSV text into a rectangular primitive matrix. Empty trailing rows
 * are dropped; a lone empty line yields an empty matrix. Values are type
 * inferred (numbers/booleans) via shared.inferPrimitive.
 */
export function parseTSV(text: string): CellPrimitive[][] {
  const rows: CellPrimitive[][] = [];
  // Split on \r\n or \n; keep embedded \n inside quoted fields intact.
  for (const line of splitLines(text)) {
    if (line.length === 0) continue; // trailing blank line
    rows.push(parseLine(line));
  }
  return rows;
}

function splitLines(text: string): string[] {
  const lines: string[] = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (inQuotes) {
      current += ch;
      if (ch === '"') {
        if (text[i + 1] === '"') {
          current += '"'; // escaped quote
          i++;
        } else {
          inQuotes = false;
        }
      }
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      current += ch;
      continue;
    }
    if (ch === "\n") {
      lines.push(current);
      current = "";
      continue;
    }
    if (ch === "\r") {
      // \r\n or standalone \r — treat as line break
      if (text[i + 1] === "\n") i++;
      lines.push(current);
      current = "";
      continue;
    }
    current += ch;
  }
  lines.push(current);
  return lines;
}

function parseLine(line: string): CellPrimitive[] {
  const fields: CellPrimitive[] = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]!;
    if (inQuotes) {
      current += ch;
      if (ch === '"') {
        if (line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      }
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      current += ch;
      continue;
    }
    if (ch === "\t") {
      fields.push(inferPrimitive(unquote(current)));
      current = "";
      continue;
    }
    current += ch;
  }
  fields.push(inferPrimitive(unquote(current)));
  return fields;
}

function unquote(field: string): string {
  if (field.length >= 2 && field.startsWith('"') && field.endsWith('"')) {
    return field.slice(1, -1).replace(/""/g, '"');
  }
  return field;
}
