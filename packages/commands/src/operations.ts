// applyOperations wire format (public SDK boundary).

import type { CellPrimitive } from "@opensheet/shared";

/** Operation set. */
export type SheetOperation =
  | { type: "cell.set"; range: string; value: CellPrimitive }
  | { type: "cell.clear"; range: string }
  | { type: "range.write"; range: string; values: CellPrimitive[][] }
  | { type: "formula.set"; range: string; formula: string }
  | { type: "sheet.freeze"; frozenRows: number; frozenColumns: number }
  | { type: "range.style"; range: string; style: Partial<import("@opensheet/shared").CellStyle> }
  | { type: "row.insert"; at: number; count?: number }
  | { type: "row.delete"; at: number; count?: number }
  | { type: "column.insert"; at: number; count?: number }
  | { type: "column.delete"; at: number; count?: number };

export interface ApplyOperationsRequest {
  workbookId: string;
  sheetId: string;
  operations: SheetOperation[];
  atomic?: boolean;
}

export interface ApplyOperationsResult {
  operationId: string;
  status: "completed";
  affectedCells: number;
  warnings: string[];
}

export class ApplyOperationsError extends Error {
  readonly operationId: string;
  readonly status = "failed" as const;
  readonly failedOperationIndex: number;
  readonly errorCode: string;

  constructor(init: {
    operationId: string;
    failedOperationIndex: number;
    errorCode: string;
    message: string;
  }) {
    super(init.message);
    this.name = "ApplyOperationsError";
    this.operationId = init.operationId;
    this.failedOperationIndex = init.failedOperationIndex;
    this.errorCode = init.errorCode;
  }
}
