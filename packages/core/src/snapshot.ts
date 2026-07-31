// Workbook <-> WorkbookSnapshot conversion (version 1).

import type {
  CellStoreFactory,
  WorkbookSnapshot,
  WorksheetSnapshot,
} from "@opensheet/shared";
import { SheetError, WORKBOOK_SNAPSHOT_VERSION } from "@opensheet/shared";
import { StyleTable } from "./styles.js";
import { Workbook } from "./workbook.js";
import { Worksheet } from "./worksheet.js";

export function toWorksheetSnapshot(sheet: Worksheet): WorksheetSnapshot {
  const cells: Record<string, import("@opensheet/shared").CellData> = {};
  for (const [row, col, data] of sheet.cellEntries()) {
    cells[`${row}:${col}`] = { ...data };
  }
  return {
    id: sheet.id,
    name: sheet.name,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    cells,
    rowHeights: Object.fromEntries([...sheet.rowHeights].map(([k, v]) => [String(k), v])),
    columnWidths: Object.fromEntries([...sheet.columnWidths].map(([k, v]) => [String(k), v])),
    frozenRows: sheet.frozenRows,
    frozenColumns: sheet.frozenColumns,
  };
}

export function toWorkbookSnapshot(workbook: Workbook): WorkbookSnapshot {
  return {
    id: workbook.id,
    name: workbook.name,
    activeSheetId: workbook.activeSheetId,
    sheets: workbook.listSheets().map(toWorksheetSnapshot),
    styles: workbook.styles.toJSON(),
    version: WORKBOOK_SNAPSHOT_VERSION,
  };
}

export interface LoadOptions {
  storeFactory?: CellStoreFactory;
}

export function worksheetFromSnapshot(
  snapshot: WorksheetSnapshot,
  options?: LoadOptions,
): Worksheet {
  const sheet = new Worksheet({
    id: snapshot.id,
    name: snapshot.name,
    rowCount: snapshot.rowCount,
    columnCount: snapshot.columnCount,
    ...(options?.storeFactory !== undefined ? { storeFactory: options.storeFactory } : {}),
  });
  sheet.frozenRows = snapshot.frozenRows;
  sheet.frozenColumns = snapshot.frozenColumns;
  for (const [key, data] of Object.entries(snapshot.cells)) {
    const sep = key.indexOf(":");
    sheet.setCell(Number(key.slice(0, sep)), Number(key.slice(sep + 1)), { ...data });
  }
  for (const [index, height] of Object.entries(snapshot.rowHeights)) {
    sheet.rowHeights.set(Number(index), height);
  }
  for (const [index, width] of Object.entries(snapshot.columnWidths)) {
    sheet.columnWidths.set(Number(index), width);
  }
  return sheet;
}

export function workbookFromSnapshot(snapshot: WorkbookSnapshot, options?: LoadOptions): Workbook {
  if (snapshot.version !== WORKBOOK_SNAPSHOT_VERSION) {
    throw new SheetError(
      "E_VALIDATION",
      `Unsupported snapshot version ${snapshot.version} (expected ${WORKBOOK_SNAPSHOT_VERSION})`,
    );
  }
  const workbook = new Workbook({ id: snapshot.id, name: snapshot.name });
  workbook.version = snapshot.version;
  workbook.styles.replaceWith(StyleTable.fromJSON(snapshot.styles));
  for (const sheetSnapshot of snapshot.sheets) {
    workbook.addSheet(worksheetFromSnapshot(sheetSnapshot, options));
  }
  if (snapshot.sheets.length > 0) {
    workbook.setActiveSheet(snapshot.activeSheetId);
  }
  return workbook;
}
