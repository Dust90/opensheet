import { describe, expect, it } from "vitest";
import { installCSVWorker, type CSVWorkerPort } from "../csv-worker-entry.js";
import type { CSVWorkerResponse } from "../csv-worker-protocol.js";

class FakePort implements CSVWorkerPort {
  readonly sent: CSVWorkerResponse[] = [];
  private listener: ((event: MessageEvent<unknown>) => void) | undefined;
  started = false;
  postMessage(message: CSVWorkerResponse): void { this.sent.push(message); }
  addEventListener(_type: "message", listener: (event: MessageEvent<unknown>) => void): void { this.listener = listener; }
  start(): void { this.started = true; }
  send(data: unknown): void { this.listener?.({ data } as MessageEvent<unknown>); }
}

describe("CSV Worker entry", () => {
  it("validates incoming messages and streams task responses through the port", () => {
    const port = new FakePort(); installCSVWorker(port, { batchRows: 2 });
    port.send({ type: "start", taskId: "t" }); port.send({ type: "chunk", taskId: "t", text: "a\nb\nc" }); port.send({ type: "finish", taskId: "t" });
    expect(port.started).toBe(true);
    expect(port.sent).toEqual([
      { type: "rows", taskId: "t", rows: [["a"], ["b"]] },
      { type: "rows", taskId: "t", rows: [["c"]] },
      { type: "done", taskId: "t", rowCount: 3 },
    ]);
  });
  it("maps malformed postMessage input to a protocol error", () => {
    const port = new FakePort(); installCSVWorker(port); port.send({ type: "chunk", taskId: "", text: 1 });
    expect(port.sent).toEqual([expect.objectContaining({ type: "error", taskId: "__invalid__", code: "E_VALIDATION" })]);
  });
  it("cleans up a valid task ID when protocol validation fails", () => {
    const port = new FakePort(); installCSVWorker(port);
    port.send({ type: "start", taskId: "t" });
    port.send({ type: "chunk", taskId: "t", text: 1 });
    port.send({ type: "start", taskId: "t" });
    port.send({ type: "finish", taskId: "t" });
    expect(port.sent).toEqual([
      expect.objectContaining({ type: "error", taskId: "t", code: "E_VALIDATION" }),
      { type: "done", taskId: "t", rowCount: 0 },
    ]);
  });
});
