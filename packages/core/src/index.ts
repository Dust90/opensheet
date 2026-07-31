// @opensheet/core — workbook data model. Depends only on @opensheet/shared.

export { Worksheet } from "./worksheet.js";
export type { WorksheetInit } from "./worksheet.js";
export { Workbook } from "./workbook.js";
export type { WorkbookInit } from "./workbook.js";
export { StyleTable } from "./styles.js";
export {
  toWorkbookSnapshot,
  toWorksheetSnapshot,
  workbookFromSnapshot,
  worksheetFromSnapshot,
} from "./snapshot.js";
export type { LoadOptions } from "./snapshot.js";

export { StringKeyCellStore, stringKeyCellStoreFactory } from "./cell-store/string-key-store.js";
export {
  NumberKeyCellStore,
  numberKeyCellStoreFactory,
  KEY_STRIDE,
} from "./cell-store/number-key-store.js";
export { ChunkedCellStore, chunkedCellStoreFactory } from "./cell-store/chunked-store.js";
