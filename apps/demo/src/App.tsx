import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { formatAddress, inferPrimitive, type CellPrimitive } from "@opensheet/shared";
import { SheetGrid, FilteredRowProjection } from "@opensheet/renderer-canvas";
import { createOpenSheet, createPersistence, type WorkbookInfo } from "@opensheet/runtime";
import { cellsToTSV, parseTSV } from "@opensheet/clipboard";

/**
 * M2 demo: DOM editor + clipboard + structure/style commands + persistence.
 * - SheetGrid consumes ONLY api.getWorksheetView() (readonly boundary)
 * - Editor commits / pastes / toolbar ops all go through applyOperations
 *   (Command Bus transactions) — one paste = ONE atomic history entry
 * - Snapshot auto-saves (debounced) from committed events only
 */

function App() {
  const api = useMemo(() => createOpenSheet(), []);
  const [workbook, setWorkbook] = useState<WorkbookInfo | null>(null);
  const [bootStatus, setBootStatus] = useState("");
  const containerRef = useRef<HTMLDivElement>(null);
  const gridRef = useRef<SheetGrid | null>(null);
  const [activeCell, setActiveCell] = useState("A1");
  const [frozen, setFrozen] = useState(false);
  const [loadedRows, setLoadedRows] = useState(0);
  const [status, setStatus] = useState("");
  const [sheetId, setSheetId] = useState("");

  // Boot: restore persisted snapshot or create a fresh workbook.
  useEffect(() => {
    const persistence = createPersistence(api, { debounceMs: 400 });
    const restored = persistence.restore();
    if (restored !== null) {
      setWorkbook(restored);
      setSheetId(restored.activeSheetId);
      setBootStatus("Restored from storage");
    } else {
      const wb = api.createWorkbook({ name: "M2 Demo" });
      setWorkbook(wb);
      setSheetId(wb.activeSheetId);
      setBootStatus("New workbook");
    }
    const stop = persistence.autoSave();
    // Flush the pending debounced save before the page is torn down.
    const flush = () => persistence.flush();
    window.addEventListener("pagehide", flush);
    const onVisibility = () => {
      if (document.visibilityState === "hidden") flush();
    };
    document.addEventListener("visibilitychange", onVisibility);
    return () => {
      window.removeEventListener("pagehide", flush);
      document.removeEventListener("visibilitychange", onVisibility);
      stop();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api]);

  const reportError = useCallback((error: unknown) => {
    setStatus(error instanceof Error ? `Error: ${error.message}` : String(error));
  }, []);

  const apply = useCallback(
    async (operations: Parameters<typeof api.applyOperations>[0]["operations"]) => {
      if (workbook === null) return;
      try {
        await api.applyOperations({
          workbookId: workbook.id,
          sheetId,
          atomic: true,
          operations,
        });
      } catch (error) {
        reportError(error);
      }
    },
    [api, workbook, sheetId, reportError],
  );

  const handleCommitCell = useCallback(
    (init: { row: number; col: number; text: string }) => {
      const address = formatAddress({ row: init.row, col: init.col });
      // M3: input starting with "=" becomes a formula (validated by the
      // command — syntax errors reject the commit); anything else is a
      // literal value (Excel-like inference, any old formula replaced).
      if (init.text.trim().startsWith("=")) {
        void apply([{ type: "formula.set", range: address, formula: init.text.trim() }]);
        return;
      }
      void apply([{ type: "cell.set", range: address, value: inferPrimitive(init.text) }]);
    },
    [apply],
  );

  const handleCopyCells = useCallback(
    (cells: CellPrimitive[][]) => {
      const tsv = cellsToTSV(cells);
      void navigator.clipboard.writeText(tsv).catch(() => undefined);
    },
    [],
  );

  const handlePasteRequest = useCallback(
    (active: { row: number; col: number }) => {
      void (async () => {
        let text: string;
        try {
          text = await navigator.clipboard.readText();
        } catch {
          setStatus("Clipboard unavailable");
          return;
        }
        const cells = parseTSV(text);
        if (cells.length === 0) return;
        const endRow = active.row + cells.length - 1;
        const endCol = active.col + cells[0]!.length - 1;
        const range = `${formatAddress({ row: active.row, col: active.col })}:${formatAddress({ row: endRow, col: endCol })}`;
        await apply([{ type: "range.write", range, values: cells }]);
        setStatus(`Pasted ${cells.length}×${cells[0]!.length}`);
      })();
    },
    [apply],
  );

  useEffect(() => {
    const container = containerRef.current;
    if (container === null || workbook === null) return;
    const grid = new SheetGrid({
      container,
      worksheet: api.getWorksheetView(sheetId),
      onChange: (listener) => api.onChange(listener),
      resolveStyle: (styleId) => api.resolveStyle(styleId),
      onSelectionChange: ({ activeRow, activeCol }) =>
        setActiveCell(formatAddress({ row: activeRow, col: activeCol })),
      onCommitCell: handleCommitCell,
      onCopyCells: handleCopyCells,
      onPasteRequest: handlePasteRequest,
      onFrame: (stats) => {
        (window as unknown as { __frameStats?: { paintMs: number; full: boolean; paintedCells: number }[] }).__frameStats ??= [];
        (window as unknown as { __frameStats: { paintMs: number; full: boolean; paintedCells: number }[] }).__frameStats.push(stats);
      },
    });
    gridRef.current = grid;
    // M4.2-F: Runtime exposes filter data only; the host owns the renderer
    // projection. Coalesce user + derived events in one microtask refresh.
    let refreshQueued = false;
    const refreshProjection = () => {
      refreshQueued = false;
      const state = api.getFilterProjectionState(sheetId);
      if (state.filter === null) {
        grid.setRowProjection(null);
      } else {
        grid.setRowProjection(new FilteredRowProjection(
          api.getWorksheetView(sheetId).rowCount,
          { startRow: state.filter.range.startRow, endRow: state.filter.range.endRow },
          state.visibleRows!,
        ));
      }
    };
    const queueProjectionRefresh = () => {
      if (refreshQueued) return;
      refreshQueued = true;
      queueMicrotask(refreshProjection);
    };
    const projectionUnsubscribe = api.onChange((event) => {
      if (event.sheetId !== sheetId) return;
      if (event.changes.some((change) => change.kind === "filter")) {
        queueProjectionRefresh();
        return;
      }
      const filter = api.getFilterProjectionState(sheetId).filter;
      if (filter !== null && event.changes.some((change) =>
        change.kind === "cells" &&
        change.range.startRow <= filter.range.endRow && change.range.endRow >= filter.range.startRow &&
        change.range.startCol <= filter.range.endCol && change.range.endCol >= filter.range.startCol,
      )) queueProjectionRefresh();
    });
    refreshProjection();
    // Test probes (E2E): expose grid/api for driving combined frames, copying,
    // and snapshot assertions.
    (window as unknown as { __grid?: SheetGrid }).__grid = grid;
    (window as unknown as { __api?: typeof api }).__api = api;
    (window as unknown as { __workbookId?: string }).__workbookId = workbook.id;
    (window as unknown as { __sheetId?: string }).__sheetId = sheetId;
    // M4.1 E2E probe: let tests install projections without a filter UI.
    (window as unknown as { __FilteredRowProjection?: typeof FilteredRowProjection }).__FilteredRowProjection =
      FilteredRowProjection;
    return () => {
      projectionUnsubscribe();
      (window as unknown as { __grid?: SheetGrid }).__grid = undefined;
      (window as unknown as { __api?: typeof api }).__api = undefined;
      (window as unknown as { __workbookId?: string }).__workbookId = undefined;
      (window as unknown as { __sheetId?: string }).__sheetId = undefined;
      (window as unknown as { __FilteredRowProjection?: typeof FilteredRowProjection }).__FilteredRowProjection =
        undefined;
      gridRef.current = null;
      grid.destroy();
    };
  }, [api, workbook, sheetId, handleCommitCell, handleCopyCells, handlePasteRequest]);

  const writeSample = async () => {
    await apply([
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
    ]);
    setStatus("Sample written");
  };

  const loadLargeDataset = async () => {
    if (workbook === null) return;
    const ROWS = 100_000;
    const COLS = 20;
    const CHUNK = 5_000;
    try {
      setStatus("Creating 100k-row sheet…");
      const big = api.createSheet({ name: "Data-100k", rows: ROWS, columns: COLS });
      setSheetId(big.id);
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
      setLoadedRows(ROWS);
      setStatus(`Loaded ${ROWS.toLocaleString()}×${COLS}`);
    } catch (error) {
      reportError(error);
    }
  };

  const toggleFreeze = async () => {
    await apply([frozen ? { type: "sheet.freeze", frozenRows: 0, frozenColumns: 0 } : { type: "sheet.freeze", frozenRows: 1, frozenColumns: 1 }]);
    setFrozen(!frozen);
  };

  const structureOp = (op: "row.insert" | "row.delete" | "column.insert" | "column.delete") => {
    const active = gridRef.current?.getSelectionState().active ?? { row: 0, col: 0 };
    void apply([
      op === "row.insert" ? { type: "row.insert", at: active.row } :
      op === "row.delete" ? { type: "row.delete", at: active.row } :
      op === "column.insert" ? { type: "column.insert", at: active.col } :
      { type: "column.delete", at: active.col },
    ]);
  };

  const styleOp = (style: Record<string, unknown>) => {
    const active = gridRef.current?.getSelectionState().active ?? { row: 0, col: 0 };
    const address = formatAddress({ row: active.row, col: active.col });
    void apply([{ type: "range.style", range: address, style }]);
  };

  const undo = () => api.undo();
  const redo = () => api.redo();

  return (
    <main style={{ fontFamily: "system-ui, sans-serif", padding: 16, height: "100vh", boxSizing: "border-box", display: "flex", flexDirection: "column" }}>
      <header style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8, flexWrap: "wrap" }}>
        <strong>OpenSheet M2</strong>
        <button type="button" onClick={writeSample}>Write sample</button>
        <button type="button" onClick={loadLargeDataset} disabled={loadedRows > 0}>
          {loadedRows > 0 ? `Loaded ${loadedRows.toLocaleString()} rows` : "Load 100k×20"}
        </button>
        <button type="button" onClick={toggleFreeze}>
          {frozen ? "Unfreeze" : "Freeze row 1 + col A"}
        </button>
        <span style={{ display: "inline-flex", gap: 4 }}>
          <button type="button" onClick={() => structureOp("row.insert")} title="Insert row above">⤵ Row+</button>
          <button type="button" onClick={() => structureOp("row.delete")} title="Delete active row">Row−</button>
          <button type="button" onClick={() => structureOp("column.insert")} title="Insert column left">⤴ Col+</button>
          <button type="button" onClick={() => structureOp("column.delete")} title="Delete active column">Col−</button>
        </span>
        <span style={{ display: "inline-flex", gap: 4 }}>
          <button type="button" onClick={() => styleOp({ bold: true })}>B</button>
          <button type="button" onClick={() => styleOp({ italic: true })}><i>I</i></button>
          <button type="button" onClick={() => styleOp({ backgroundColor: "#fce8b2" })} style={{ background: "#fce8b2" }}>Bg</button>
          <button type="button" onClick={() => styleOp({ horizontalAlign: "center" })}>≡</button>
        </span>
        <button type="button" onClick={undo}>Undo</button>
        <button type="button" onClick={redo}>Redo</button>
        <span style={{ marginLeft: "auto", color: "#5f6368", fontSize: 13 }}>
          Active: {activeCell} · {status || bootStatus}
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
