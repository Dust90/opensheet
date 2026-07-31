// Recursive-descent parser (M3.2): tokens → AST + dependencies.
//
// Precedence (low→high), Excel-like:
//   comparison  = <> < > <= >=
//   concat      &
//   add-sub     + -
//   mul-div     * /
//   exponent    ^   (right-assoc, binds tighter than unary minus in Excel:
//                    -2^2 = -(2^2) = -4)
//   unary       - + %
//   postfix     %   (left-to-right with unary; we fold % into unary)
//   primary     number | string | bool | null | cell | range | call | (expr)

import { SheetError } from "@opensheet/shared";
import { colToName, parseAddress } from "@opensheet/shared";
import type { CellRef, Expr, FormulaParseResult } from "./ast.js";
import { collectDependencies } from "./ast.js";
import { tokenize, type Token } from "./lexer.js";

const COMPARISON = new Set(["=", "<>", "<", ">", "<=", ">="]);
const CONCAT = new Set(["&"]);
const ADDITIVE = new Set(["+", "-"]);
const MULTIPLICATIVE = new Set(["*", "/"]);
const ERROR_LITERALS: Record<string, Expr> = {
  "#REF!": { kind: "error", error: { type: "#REF!" } },
  "#VALUE!": { kind: "error", error: { type: "#VALUE!" } },
  "#DIV/0!": { kind: "error", error: { type: "#DIV/0!" } },
  "#NAME?": { kind: "error", error: { type: "#NAME?" } },
  "#N/A": { kind: "error", error: { type: "#N/A" } },
  "#CYCLE!": { kind: "error", error: { type: "#CYCLE!" } },
  "#NUM!": { kind: "error", error: { type: "#NUM!" } },
};

class Parser {
  private readonly tokens: Token[];
  private index = 0;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  private peek(): Token {
    return this.tokens[this.index]!;
  }

  private next(): Token {
    const token = this.tokens[this.index]!;
    this.index++;
    return token;
  }

  private expect(type: Token["type"], what: string): Token {
    const token = this.peek();
    if (token.type !== type) {
      throw new SheetError("E_FORMULA_SYNTAX", `Expected ${what} at position ${token.pos}, got "${token.value}"`);
    }
    return this.next();
  }

  parse(): Expr {
    if (this.peek().type === "eof") {
      throw new SheetError("E_FORMULA_SYNTAX", "Formula body is empty");
    }
    const expr = this.parseComparison();
    if (this.peek().type !== "eof") {
      throw new SheetError(
        "E_FORMULA_SYNTAX",
        `Unexpected token "${this.peek().value}" at position ${this.peek().pos}`,
      );
    }
    return expr;
  }

  private parseComparison(): Expr {
    let left = this.parseConcat();
    while (COMPARISON.has(this.peek().value)) {
      const op = this.next().value;
      const right = this.parseConcat();
      left = { kind: "binary", op, left, right };
    }
    return left;
  }

  private parseConcat(): Expr {
    let left = this.parseAdditive();
    while (CONCAT.has(this.peek().value)) {
      const op = this.next().value;
      const right = this.parseAdditive();
      left = { kind: "binary", op, left, right };
    }
    return left;
  }

  private parseAdditive(): Expr {
    let left = this.parseMultiplicative();
    while (ADDITIVE.has(this.peek().value)) {
      const op = this.next().value;
      const right = this.parseMultiplicative();
      left = { kind: "binary", op, left, right };
    }
    return left;
  }

  private parseMultiplicative(): Expr {
    let left = this.parseUnary();
    while (MULTIPLICATIVE.has(this.peek().value)) {
      const op = this.next().value;
      const right = this.parseUnary();
      left = { kind: "binary", op, left, right };
    }
    return left;
  }

  private parseUnary(): Expr {
    const token = this.peek();
    if (token.type === "op" && (token.value === "-" || token.value === "+")) {
      this.next();
      const operand = this.parseUnary();
      return { kind: "unary", op: token.value as "-" | "+", operand };
    }
    const expr = this.parseExponent();
    // Postfix percent: A1% == A1/100
    if (this.peek().type === "op" && this.peek().value === "%") {
      this.next();
      return { kind: "unary", op: "%", operand: expr };
    }
    return expr;
  }

  /** Right-associative power; -2^2 parses as -(2^2). The exponent side
   *  accepts a unary prefix so 2^-2 and 2^-n are legal. */
  private parseExponent(): Expr {
    const left = this.parsePrimary();
    if (this.peek().type === "op" && this.peek().value === "^") {
      this.next();
      const right = this.parseUnary();
      return { kind: "binary", op: "^", left, right };
    }
    return left;
  }

  private parsePrimary(): Expr {
    const token = this.peek();
    switch (token.type) {
      case "number": {
        this.next();
        return { kind: "number", value: Number(token.value) };
      }
      case "string": {
        this.next();
        return { kind: "string", value: token.value };
      }
      case "bool": {
        this.next();
        return { kind: "bool", value: token.value === "TRUE" };
      }
      case "null": {
        this.next();
        return { kind: "null" };
      }
      case "cell": {
        this.next();
        const start = parseCellRef(token.value);
        if (this.peek().type === "colon") {
          this.next();
          const endToken = this.expect("cell", "cell reference after ':'");
          const end = parseCellRef(endToken.value);
          return { kind: "range", start, end };
        }
        return { kind: "cell", ref: start };
      }
      case "ident": {
        const name = token.value;
        if (ERROR_LITERALS[name.toUpperCase()] !== undefined) {
          this.next();
          return ERROR_LITERALS[name.toUpperCase()]!;
        }
        this.next();
        this.expect("lparen", "'(' after function name");
        const args: Expr[] = [];
        if (this.peek().type !== "rparen") {
          args.push(this.parseComparison());
          while (this.peek().type === "comma") {
            this.next();
            args.push(this.parseComparison());
          }
        }
        this.expect("rparen", "')'");
        return { kind: "function", name: name.toUpperCase(), args };
      }
      case "error": {
        this.next();
        return ERROR_LITERALS[token.value]!;
      }
      case "lparen": {
        this.next();
        const inner = this.parseComparison();
        this.expect("rparen", "')'");
        return inner;
      }
      default:
        throw new SheetError(
          "E_FORMULA_SYNTAX",
          `Unexpected token "${token.value}" at position ${token.pos}`,
        );
    }
  }
}

/** Parse "A1", "$A$1", "A$1" → 0-based CellRef. */
export function parseCellRef(text: string): CellRef {
  const m = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(text);
  if (m === null) {
    throw new SheetError("E_FORMULA_SYNTAX", `Invalid cell reference "${text}"`);
  }
  const [, colAbs, letters, rowAbs, rowText] = m;
  let col = 0;
  for (const ch of letters!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  const row = Number(rowText) - 1;
  // Validate against the shared address helpers (max bounds + shape).
  parseAddress(`${colToName(col - 1)}${row + 1}`);
  return {
    row,
    col: col - 1,
    rowAbs: rowAbs === "$",
    colAbs: colAbs === "$",
  };
}

/** Parse a full formula string including the leading "=". */
export function parseFormula(input: string): FormulaParseResult {
  const text = input.trim();
  if (!text.startsWith("=")) {
    throw new SheetError("E_FORMULA_SYNTAX", "Formula must start with '='");
  }
  const body = text.slice(1).trim();
  const parser = new Parser(tokenize(body));
  const ast = parser.parse();
  return { ast, dependencies: collectDependencies(ast) };
}
