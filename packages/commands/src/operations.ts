// applyOperations wire format (public SDK boundary).

import type { CellPrimitive } from "@opensheet/shared";

/** M0 operation set; extended in later milestones (formula.set, row/col ops, ...). */
export type SheetOperation =
  | { type: "cell.set"; range: string; value: CellPrimitive }
  | { type: "cell.clear"; range: string }
  | { type: "range.write"; range: string; values: CellPrimitive[][] };

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
