// @opensheet/runtime — composition root and public SDK entry point.

export { createOpenSheet } from "./create-opensheet.js";
export type { OpenSheetOptions } from "./create-opensheet.js";
export { createPersistence, validateSnapshot } from "./persistence.js";
export type { Persistence, PersistenceOptions, StorageLike } from "./persistence.js";
export type {
  ImportCSVResult,
  OpenSheetAPI,
  SheetInfo,
  WorkbookInfo,
} from "./api.js";
export type { ApplyOperationsRequest, ApplyOperationsResult } from "./api.js";
// Re-export contracts hosts commonly need.
export type {
  CellAddress,
  CellData,
  CellPrimitive,
  CellValue,
  ChangeEvent,
  WorkbookSnapshot,
  WorksheetSnapshot,
} from "@opensheet/shared";
export { SheetError } from "@opensheet/shared";
