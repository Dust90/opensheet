// @opensheet/formula-engine — M3 formula kernel (no runtime wiring yet).

export { tokenize } from "./lexer.js";
export type { Token, TokenType } from "./lexer.js";
export { parseFormula, parseCellRef } from "./parser.js";
export type { CellRangeRef, CellRef, Expr, FormulaDependencies, FormulaParseResult } from "./ast.js";
export { collectDependencies, iterateRange, rangeBounds, walkExpr } from "./ast.js";
export { evaluateExpr } from "./evaluate.js";
export type { CellRangeValue, EvaluationBudget, FormulaArgument, FormulaContext } from "./evaluate.js";
export { makeBudget } from "./evaluate.js";
export { finiteNumber } from "./numeric.js";
export { FunctionRegistry, createDefaultFunctions, iterateValues } from "./functions.js";
export type { FunctionImpl } from "./functions.js";
export { DependencyGraph } from "./dependency.js";
export { exprToString, rewriteFormulaReferences, rewrittenDependencies, translateFormulaReferences } from "./rewrite.js";
export type { FormulaReferenceBounds, StructureChange } from "./rewrite.js";
