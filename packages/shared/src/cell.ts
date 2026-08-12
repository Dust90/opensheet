// @injoysai/opensheet-shared — cell data contracts

export type CellPrimitive = string | number | boolean | null;

export type CellErrorType =
  | "#REF!"
  | "#VALUE!"
  | "#DIV/0!"
  | "#NAME?"
  | "#N/A"
  | "#CYCLE!"
  | "#NUM!";
export const CELL_ERROR_TYPES: readonly CellErrorType[] = [
  "#REF!",
  "#VALUE!",
  "#DIV/0!",
  "#NAME?",
  "#N/A",
  "#CYCLE!",
  "#NUM!",
];

export interface CellError {
  type: CellErrorType;
  message?: string;
}

export type CellValue = CellPrimitive | CellError;

export function isCellError(value: CellValue): value is CellError {
  return typeof value === "object" && value !== null && "type" in value;
}

export interface CellData {
  /** Last computed value (formula cache) or the literal value. Renderers read this only. */
  value: CellValue;
  /** Formula source including the leading "=", e.g. "=SUM(A1:A10)". */
  formula?: string;
  /** Reference into the workbook style table. */
  styleId?: string;
  numberFormat?: string;
}

/** Presentation attributes. Stored deduplicated in the workbook style table. */
export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontFamily?: string;
  textColor?: string;
  backgroundColor?: string;
  horizontalAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  border?: {
    top?: BorderEdge;
    right?: BorderEdge;
    bottom?: BorderEdge;
    left?: BorderEdge;
  };
}

export interface BorderEdge {
  style: "thin" | "medium" | "thick" | "dashed";
  color?: string;
}
