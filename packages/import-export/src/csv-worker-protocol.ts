import { SheetError, type SheetErrorCode } from "@injoysai/opensheet-shared";
import { validateCSVOptions, type CSVOptions } from "./csv.js";

export type CSVWorkerRequest =
  | { type: "start"; taskId: string; options?: CSVOptions }
  | { type: "chunk"; taskId: string; text: string }
  | { type: "finish"; taskId: string }
  | { type: "cancel"; taskId: string };

export type CSVWorkerResponse =
  | { type: "rows"; taskId: string; rows: string[][] }
  | { type: "done"; taskId: string; rowCount: number }
  | { type: "error"; taskId: string; code: SheetErrorCode; message: string };

function requireTaskId(value: unknown): asserts value is string {
  if (typeof value !== "string" || value.length === 0) throw new SheetError("E_VALIDATION", "CSV worker taskId must be a non-empty string");
}

/** Validate untyped postMessage input before it reaches the Worker state machine. */
export function validateCSVWorkerRequest(value: unknown): asserts value is CSVWorkerRequest {
  if (typeof value !== "object" || value === null || Array.isArray(value)) throw new SheetError("E_VALIDATION", "CSV worker request must be an object");
  const request = value as Record<string, unknown>;
  requireTaskId(request.taskId);
  if (request.type === "start") {
    if (request.options !== undefined) validateCSVOptions(request.options);
    return;
  }
  if (request.type === "chunk") {
    if (typeof request.text !== "string") throw new SheetError("E_VALIDATION", "CSV worker chunk text must be a string");
    return;
  }
  if (request.type === "finish" || request.type === "cancel") return;
  throw new SheetError("E_VALIDATION", "CSV worker request has an unknown type");
}
