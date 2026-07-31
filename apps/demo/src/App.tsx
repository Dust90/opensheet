import { useMemo, useState } from "react";
import { createOpenSheet, type WorkbookInfo } from "@opensheet/runtime";

/**
 * M0 demo: proves the runtime pipeline end-to-end in a browser —
 * createWorkbook → applyOperations (transaction) → readRange → undo/redo.
 * The real Canvas grid lands in M1; this is a minimal DOM table.
 */
function App() {
  const api = useMemo(() => createOpenSheet(), []);
  const [workbook] = useState<WorkbookInfo>(() =>
    api.createWorkbook({ name: "M0 Smoke Book" }),
  );
  const [version, setVersion] = useState(0);

  const sheetId = workbook.activeSheetId;
  const values = api.readRange({ sheetId, range: "A1:E8" });

  const writeSample = async () => {
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId,
      atomic: true,
      operations: [
        { type: "range.write", range: "A1:C1", values: [["Item", "Qty", "Price"]] },
        { type: "range.write", range: "A2:C3", values: [["Apples", 12, 3.5], ["Pears", 7, 4.2]] },
      ],
    });
    setVersion((v) => v + 1);
  };

  const undo = () => {
    api.undo();
    setVersion((v) => v + 1);
  };

  const redo = () => {
    api.redo();
    setVersion((v) => v + 1);
  };

  return (
    <main style={{ fontFamily: "system-ui, sans-serif", padding: 24 }}>
      <h1>OpenSheet — M0 kernel smoke test</h1>
      <p>
        Workbook: <strong>{workbook.name}</strong> ({workbook.id.slice(0, 8)}…) · v{version}
      </p>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button type="button" onClick={writeSample}>
          Write sample (atomic)
        </button>
        <button type="button" onClick={undo}>
          Undo
        </button>
        <button type="button" onClick={redo}>
          Redo
        </button>
      </div>
      <table style={{ borderCollapse: "collapse" }}>
        <tbody>
          {values.map((row, r) => (
            <tr key={r}>
              {row.map((cell, c) => (
                <td
                  key={c}
                  style={{
                    border: "1px solid #d0d0d0",
                    minWidth: 90,
                    height: 26,
                    padding: "0 8px",
                    fontSize: 13,
                  }}
                >
                  {cell === null ? "" : String(cell)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </main>
  );
}

export default App;
