import { useEffect, useMemo, useRef, useState } from "react";
import { formatAddress } from "@opensheet/shared";
import { SheetGrid } from "@opensheet/renderer-canvas";
import { createOpenSheet, type WorkbookInfo } from "@opensheet/runtime";

/**
 * M1 demo: dual-canvas grid wired to the runtime.
 * - SheetGrid consumes ONLY api.getWorksheetView() (readonly boundary)
 * - All mutations go through applyOperations (Command Bus transactions)
 * - Change events drive dirty-region repaints automatically
 */
function App() {
  const api = useMemo(() => createOpenSheet(), []);
  const [workbook] = useState<WorkbookInfo>(() => api.createWorkbook({ name: "M1 Grid Demo" }));
  const containerRef = useRef<HTMLDivElement>(null);
  const [activeCell, setActiveCell] = useState("A1");
  const [frozen, setFrozen] = useState(false);
  const [loadedRows, setLoadedRows] = useState(0);
  const [status, setStatus] = useState("");
  const [sheetId, setSheetId] = useState(workbook.activeSheetId);

  useEffect(() => {
    const container = containerRef.current;
    if (container === null) return;
    const grid = new SheetGrid({
      container,
      worksheet: api.getWorksheetView(sheetId),
      onChange: (listener) => api.onChange(listener),
      resolveStyle: (styleId) => api.resolveStyle(styleId),
      onSelectionChange: ({ activeRow, activeCol }) =>
        setActiveCell(formatAddress({ row: activeRow, col: activeCol })),
      onFrame: (stats) => {
        (window as unknown as { __frameStats?: { paintMs: number; full: boolean; paintedCells: number }[] }).__frameStats ??= [];
        (window as unknown as { __frameStats: { paintMs: number; full: boolean; paintedCells: number }[] }).__frameStats.push(stats);
      },
    });
    return () => grid.destroy();
  }, [api, sheetId]);

  const reportError = (error: unknown) => {
    setStatus(error instanceof Error ? `Error: ${error.message}` : String(error));
  };

  const writeSample = async () => {
    try {
      await api.applyOperations({
        workbookId: workbook.id,
        sheetId,
        atomic: true,
        operations: [
          { type: "range.write", range: "A1:C1", values: [["Item", "Qty", "Price"]] },
          {
            type: "range.write",
            range: "A2:C4",
            values: [
              ["Apples", 12, 3.5],
              ["Pears", 7, 4.2],
              ["Plums", 25, 2.8],
            ],
          },
        ],
      });
      setStatus("Sample written");
    } catch (error) {
      reportError(error);
    }
  };

  const loadLargeDataset = async () => {
    // The default sheet is 1000×26 — create a purpose-sized sheet first.
    const ROWS = 100_000;
    const COLS = 20;
    const CHUNK = 5_000;
    try {
      setStatus("Creating 100k-row sheet…");
      const big = api.createSheet({ name: "Data-100k", rows: ROWS, columns: COLS });
      setSheetId(big.id);
      const t0 = performance.now();
      for (let start = 0; start < ROWS; start += CHUNK) {
        const height = Math.min(CHUNK, ROWS - start);
        const values: (string | number)[][] = Array.from({ length: height }, (_, r) =>
          Array.from({ length: COLS }, (_, c) =>
            c === 0 ? `row-${start + r + 1}` : ((start + r) * 31 + c * 17) % 10_000,
          ),
        );
        const endRow = start + height;
        await api.applyOperations({
          workbookId: workbook.id,
          sheetId: big.id,
          atomic: true,
          operations: [
            {
              type: "range.write",
              range: `A${start + 1}:${String.fromCharCode(64 + COLS)}${endRow}`,
              values,
            },
          ],
        });
        setStatus(`Loading… ${endRow.toLocaleString()} / ${ROWS.toLocaleString()} rows`);
      }
      const elapsed = Math.round(performance.now() - t0);
      setLoadedRows(ROWS);
      setStatus(`Loaded ${ROWS.toLocaleString()}×${COLS} in ${elapsed}ms`);
    } catch (error) {
      reportError(error);
    }
  };

  const toggleFreeze = async () => {
    try {
      await api.applyOperations({
        workbookId: workbook.id,
        sheetId,
        atomic: true,
        operations: [
          frozen
            ? { type: "sheet.freeze", frozenRows: 0, frozenColumns: 0 }
            : { type: "sheet.freeze", frozenRows: 1, frozenColumns: 1 },
        ],
      });
      setFrozen(!frozen);
    } catch (error) {
      reportError(error);
    }
  };

  return (
    <main style={{ fontFamily: "system-ui, sans-serif", padding: 16, height: "100vh", boxSizing: "border-box", display: "flex", flexDirection: "column" }}>
      <header style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8, flexWrap: "wrap" }}>
        <strong>OpenSheet M1</strong>
        <button type="button" onClick={writeSample}>Write sample</button>
        <button type="button" onClick={loadLargeDataset} disabled={loadedRows > 0}>
          {loadedRows > 0 ? `Loaded ${loadedRows.toLocaleString()} rows` : "Load 100k×20"}
        </button>
        <button type="button" onClick={toggleFreeze}>
          {frozen ? "Unfreeze" : "Freeze row 1 + col A"}
        </button>
        <button type="button" onClick={() => api.undo()}>Undo</button>
        <button type="button" onClick={() => api.redo()}>Redo</button>
        <span style={{ marginLeft: "auto", color: "#5f6368", fontSize: 13 }}>
          Active: {activeCell} · {status}
        </span>
      </header>
      <div
        ref={containerRef}
        data-testid="sheet-grid"
        style={{ flex: 1, minHeight: 0, border: "1px solid #d0d3d9", borderRadius: 4 }}
      />
    </main>
  );
}

export default App;
