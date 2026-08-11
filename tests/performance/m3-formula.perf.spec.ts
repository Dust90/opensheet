import { expect, test } from "@playwright/test";
import { writeFileSync } from "node:fs";

const ROWS = 100_000;
const VALUE_COLUMNS = 9;
const CHUNK_ROWS = 5_000;

test("formula perf: lazy SUM over 900k cells stays inside the default budget", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("sheet-grid").waitFor();

  const result = await page.evaluate(
    async ({ rows, valueColumns, chunkRows }) => {
      const w = window as unknown as {
        __api: {
          createSheet(options: {
            name: string;
            rows: number;
            columns: number;
          }): { id: string };
          applyOperations(request: object): Promise<unknown>;
          getWorksheetView(sheetId: string): {
            getCell(row: number, col: number): { value: unknown } | undefined;
          };
        };
        __workbookId: string;
      };
      const sheet = w.__api.createSheet({
        name: "Formula-900k",
        rows,
        columns: valueColumns + 1,
      });
      for (let start = 0; start < rows; start += chunkRows) {
        const height = Math.min(chunkRows, rows - start);
        const values = Array.from({ length: height }, () =>
          Array.from({ length: valueColumns }, () => 1),
        );
        await w.__api.applyOperations({
          workbookId: w.__workbookId,
          sheetId: sheet.id,
          atomic: true,
          operations: [
            {
              type: "range.write",
              range: `A${start + 1}:I${start + height}`,
              values,
            },
          ],
        });
      }
      const started = performance.now();
      await w.__api.applyOperations({
        workbookId: w.__workbookId,
        sheetId: sheet.id,
        atomic: true,
        operations: [
          { type: "formula.set", range: "J1", formula: "=SUM(A1:I100000)" },
        ],
      });
      return {
        formulaMs: performance.now() - started,
        value: w.__api.getWorksheetView(sheet.id).getCell(0, valueColumns)
          ?.value,
      };
    },
    { rows: ROWS, valueColumns: VALUE_COLUMNS, chunkRows: CHUNK_ROWS },
  );

  expect(result.value).toBe(ROWS * VALUE_COLUMNS);
  console.log(JSON.stringify(result));
  if (process.env.BENCH_OUT !== undefined) {
    writeFileSync(process.env.BENCH_OUT, JSON.stringify(result, null, 2));
  }
  expect(result.formulaMs).toBeLessThanOrEqual(1_000);
});
