import { describe, expect, it } from "vitest";
import { CSVWorkerTaskHandler } from "../csv-worker-tasks.js";
import type { CSVWorkerResponse } from "../csv-worker-protocol.js";

function run(requests: Parameters<CSVWorkerTaskHandler["handle"]>[0][]): CSVWorkerResponse[] {
  const messages: CSVWorkerResponse[] = [];
  const handler = new CSVWorkerTaskHandler({ batchRows: 2 });
  for (const request of requests) handler.handle(request, (message) => messages.push(message));
  return messages;
}

describe("CSVWorkerTaskHandler", () => {
  it("streams bounded row batches then emits one final done", () => {
    expect(run([{ type: "start", taskId: "t" }, { type: "chunk", taskId: "t", text: "a\nb\nc" }, { type: "finish", taskId: "t" }])).toEqual([
      { type: "rows", taskId: "t", rows: [["a"], ["b"]] },
      { type: "rows", taskId: "t", rows: [["c"]] },
      { type: "done", taskId: "t", rowCount: 3 },
    ]);
  });
  it("never exceeds the configured batch size for a large single chunk", () => {
    const messages = run([{ type: "start", taskId: "t" }, { type: "chunk", taskId: "t", text: "a\nb\nc\nd\ne" }, { type: "finish", taskId: "t" }]);
    expect(messages.filter((message) => message.type === "rows").map((message) => (message as Extract<CSVWorkerResponse, { type: "rows" }>).rows.length)).toEqual([2, 2, 1]);
  });
  it("rejects duplicate and unknown task operations, and cancel emits no done", () => {
    expect(run([{ type: "start", taskId: "t" }, { type: "start", taskId: "t" }, { type: "cancel", taskId: "t" }, { type: "chunk", taskId: "t", text: "x" }])).toEqual([
      expect.objectContaining({ type: "error", taskId: "t", code: "E_VALIDATION" }),
      expect.objectContaining({ type: "error", taskId: "t", code: "E_VALIDATION" }),
    ]);
  });
  it("maps parser errors and releases the failed task", () => {
    expect(run([{ type: "start", taskId: "t" }, { type: "chunk", taskId: "t", text: '"unterminated' }, { type: "finish", taskId: "t" }, { type: "finish", taskId: "t" }])).toEqual([
      expect.objectContaining({ type: "error", taskId: "t", code: "E_VALIDATION" }),
      expect.objectContaining({ type: "error", taskId: "t", code: "E_VALIDATION" }),
    ]);
  });
});
