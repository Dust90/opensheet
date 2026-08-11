// Workbook <-> WorkbookSnapshot conversion. New writes are V2; V1 loads are
// explicitly migrated at this single boundary before any Worksheet is built.

import type {
  CellStoreFactory,
  WorkbookSnapshot,
  WorkbookSnapshotV1,
  WorkbookSnapshotV2,
  SupportedWorkbookSnapshot,
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
    filter: sheet.filter === null ? null : {
      range: { ...sheet.filter.range },
      hasHeader: sheet.filter.hasHeader,
      conditions: sheet.filter.conditions.map((condition) => ({ ...condition })),
    },
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
  // Keep Snapshot load on the same validation + clone boundary as commands.
  sheet.setFilter(snapshot.filter);
  return sheet;
}

/** Pure strict migration: valid V1 has no ambiguous optional filter field. */
export function migrateV1ToV2(snapshot: WorkbookSnapshotV1): WorkbookSnapshotV2 {
  return {
    id: snapshot.id,
    name: snapshot.name,
    activeSheetId: snapshot.activeSheetId,
    styles: { ...snapshot.styles },
    version: WORKBOOK_SNAPSHOT_VERSION,
    sheets: snapshot.sheets.map((sheet) => ({
      ...sheet,
      cells: Object.fromEntries(Object.entries(sheet.cells).map(([key, data]) => [key, { ...data }])),
      rowHeights: { ...sheet.rowHeights },
      columnWidths: { ...sheet.columnWidths },
      filter: null,
    })),
  };
}

export function workbookFromSnapshot(snapshot: SupportedWorkbookSnapshot, options?: LoadOptions): Workbook {
  const v2 = snapshot.version === 1 ? migrateV1ToV2(snapshot) : snapshot;
  if (v2.version !== WORKBOOK_SNAPSHOT_VERSION) {
    throw new SheetError(
      "E_VALIDATION",
      `Unsupported snapshot version ${v2.version} (expected ${WORKBOOK_SNAPSHOT_VERSION})`,
    );
  }
  const workbook = new Workbook({ id: v2.id, name: v2.name });
  workbook.version = v2.version;
  workbook.styles.replaceWith(StyleTable.fromJSON(v2.styles));
  for (const sheetSnapshot of v2.sheets) {
    workbook.addSheet(worksheetFromSnapshot(sheetSnapshot, options));
  }
  if (v2.sheets.length > 0) {
    workbook.setActiveSheet(v2.activeSheetId);
  }
  return workbook;
}
