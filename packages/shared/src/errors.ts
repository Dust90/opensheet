// @opensheet/shared — error contracts

export type SheetErrorCode =
  | "E_INVALID_ADDRESS"
  | "E_INVALID_RANGE"
  | "E_SHEET_NOT_FOUND"
  | "E_WORKBOOK_NOT_FOUND"
  | "E_VALIDATION"
  | "E_UNKNOWN_COMMAND"
  | "E_OP_FAILED"
  | "E_FORMULA_SYNTAX"
  | "E_NOT_IMPLEMENTED";

/** Base error for all OpenSheet failures. Carries a stable machine-readable code. */
export class SheetError extends Error {
  readonly code: SheetErrorCode;

  constructor(code: SheetErrorCode, message: string) {
    super(message);
    this.name = "SheetError";
    this.code = code;
  }
}

export function isSheetError(error: unknown): error is SheetError {
  return error instanceof SheetError;
}
