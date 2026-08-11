import { expect, test } from "@playwright/test";
import { writeFileSync } from "node:fs";

const ROWS = 100_000;
const COLUMNS = 20;

test("CSV perf: browser Worker imports and exports a dense 100k × 20 file", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("sheet-grid").waitFor();

  const result = await page.evaluate(
    async ({ rows, columns }) => {
      const w = window as unknown as {
        __api: {
          importCSV(options: { file: Blob }): Promise<{
            sheetId: string;
            rowCount: number;
            columnCount: number;
          }>;
          exportCSV(options: { sheetId: string }): Promise<Blob>;
          getWorksheetView(sheetId: string): {
            getCell(row: number, col: number): { value: unknown } | undefined;
          };
        };
      };
      const csv = Array.from({ length: rows }, (_, row) =>
        Array.from(
          { length: columns },
          (_, col) => `r${row + 1}c${col + 1}`,
        ).join(","),
      ).join("\r\n");
      const importStart = performance.now();
      const imported = await w.__api.importCSV({
        file: new File([csv], "dense-100k.csv", { type: "text/csv" }),
      });
      const importMs = performance.now() - importStart;
      const sheet = w.__api.getWorksheetView(imported.sheetId);
      const exportStart = performance.now();
      const exported = await w.__api.exportCSV({ sheetId: imported.sheetId });
      const exportMs = performance.now() - exportStart;
      const text = await exported.text();
      return {
        importMs,
        exportMs,
        rowCount: imported.rowCount,
        columnCount: imported.columnCount,
        first: sheet.getCell(0, 0)?.value,
        last: sheet.getCell(rows - 1, columns - 1)?.value,
        csvBytes: exported.size,
        exportStartsWith: text.startsWith("r1c1,r1c2"),
        exportEndsWith: text.endsWith(`r${rows}c${columns}`),
      };
    },
    { rows: ROWS, columns: COLUMNS },
  );

  expect(result).toMatchObject({
    rowCount: ROWS,
    columnCount: COLUMNS,
    first: "r1c1",
    last: `r${ROWS}c${COLUMNS}`,
    exportStartsWith: true,
    exportEndsWith: true,
  });
  expect(result.csvBytes).toBeGreaterThan(0);

  console.log(JSON.stringify(result));
  if (process.env.BENCH_OUT !== undefined) {
    writeFileSync(process.env.BENCH_OUT, JSON.stringify(result, null, 2));
  }
  expect(result.importMs).toBeLessThanOrEqual(5_000);
  expect(result.exportMs).toBeLessThanOrEqual(2_500);
});
