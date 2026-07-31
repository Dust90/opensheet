// Formula tokenizer (M3.1). Pure: string → tokens, no I/O.
//
// Grammar surface (Excel-like):
//   number  12 | 1.5 | 1e3 | .5
//   string  "text" | 'text'  (double quotes preferred; both accepted)
//   bool    TRUE | FALSE
//   null    NULL
//   cell    A1 | $A$1 | A$1 | $A1
//   range   A1:B2  (lexed as CELL COLON CELL; parser combines)
//   op      + - * / ^ % & = <> < > <= >= ( ) , :

import { SheetError } from "@opensheet/shared";

export type TokenType =
  | "number"
  | "string"
  | "ident" // function names and unknown identifiers
  | "cell"
  | "bool"
  | "null"
  | "op"
  | "lparen"
  | "rparen"
  | "comma"
  | "colon"
  | "eof";

export interface Token {
  readonly type: TokenType;
  readonly value: string;
  readonly pos: number;
}

const TWO_CHAR_OPS = new Set(["<=", ">=", "<>"]);
const ONE_CHAR_OPS = new Set(["+", "-", "*", "/", "^", "%", "&", "=", "<", ">", "(", ")", ",", ":"]);

const CELL_RE = /^\$?[A-Za-z]{1,3}\$?\d+$/;

function isCellToken(text: string): boolean {
  if (!CELL_RE.test(text)) return false;
  // Column letters must be a valid A1 column (A..XFD).
  const letters = /^[A-Za-z]+/.exec(text.replace(/^\$/, ""))![0];
  let col = 0;
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return col >= 1 && col <= 16_384;
}

/** Tokenize a formula BODY (the text after the leading "="). */
export function tokenize(input: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  while (i < input.length) {
    const ch = input[i]!;
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      i++;
      continue;
    }
    const pos = i;
    // Number
    if (/[0-9]/.test(ch) || (ch === "." && /[0-9]/.test(input[i + 1] ?? ""))) {
      let j = i;
      while (j < input.length && /[0-9.]/.test(input[j]!)) j++;
      if ((input[j] ?? "") === "e" || (input[j] ?? "") === "E") {
        let k = j + 1;
        if ((input[k] ?? "") === "+" || (input[k] ?? "") === "-") k++;
        if (/[0-9]/.test(input[k] ?? "")) {
          while (k < input.length && /[0-9]/.test(input[k]!)) k++;
          j = k;
        }
      }
      const raw = input.slice(i, j);
      if (Number.isNaN(Number(raw))) {
        throw new SheetError("E_FORMULA_SYNTAX", `Invalid number "${raw}" at position ${pos}`);
      }
      tokens.push({ type: "number", value: raw, pos });
      i = j;
      continue;
    }
    // String literal
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      let out = "";
      while (j < input.length) {
        if (input[j] === ch) {
          if (input[j + 1] === ch) {
            out += ch; // escaped quote ("" or '')
            j += 2;
            continue;
          }
          tokens.push({ type: "string", value: out, pos });
          i = j + 1;
          break;
        }
        out += input[j]!;
        j++;
      }
      if (j >= input.length) {
        throw new SheetError("E_FORMULA_SYNTAX", `Unterminated string at position ${pos}`);
      }
      continue;
    }
    // Identifier: letters/underscore start (function names, TRUE/FALSE/NULL, cells)
    if (/[A-Za-z_$]/.test(ch)) {
      let j = i;
      while (j < input.length && /[A-Za-z0-9_.$]/.test(input[j]!)) j++;
      const raw = input.slice(i, j);
      const upper = raw.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        tokens.push({ type: "bool", value: upper, pos });
      } else if (upper === "NULL") {
        tokens.push({ type: "null", value: "NULL", pos });
      } else if (isCellToken(raw)) {
        tokens.push({ type: "cell", value: raw.toUpperCase(), pos });
      } else {
        tokens.push({ type: "ident", value: raw, pos });
      }
      i = j;
      continue;
    }
    // Operators
    const two = input.slice(i, i + 2);
    if (TWO_CHAR_OPS.has(two)) {
      tokens.push({ type: "op", value: two, pos });
      i += 2;
      continue;
    }
    if (ONE_CHAR_OPS.has(ch)) {
      const map: Record<string, TokenType> = {
        "(": "lparen",
        ")": "rparen",
        ",": "comma",
        ":": "colon",
      };
      tokens.push({ type: map[ch] ?? "op", value: ch, pos });
      i++;
      continue;
    }
    throw new SheetError("E_FORMULA_SYNTAX", `Unexpected character "${ch}" at position ${pos}`);
  }
  tokens.push({ type: "eof", value: "", pos: input.length });
  return tokens;
}
