import { describe, expect, it } from "vitest";
import { SheetError } from "@opensheet/shared";
import { validateCSVWorkerRequest } from "../csv-worker-protocol.js";

describe("CSV worker protocol", () => {
  it("accepts the four task-scoped request messages", () => {
    expect(() => validateCSVWorkerRequest({ type: "start", taskId: "t", options: { delimiter: ";" } })).not.toThrow();
    expect(() => validateCSVWorkerRequest({ type: "chunk", taskId: "t", text: "a,b" })).not.toThrow();
    expect(() => validateCSVWorkerRequest({ type: "finish", taskId: "t" })).not.toThrow();
    expect(() => validateCSVWorkerRequest({ type: "cancel", taskId: "t" })).not.toThrow();
  });
  it("rejects malformed task IDs, chunks, options, and message types", () => {
    for (const value of [null, { type: "chunk", taskId: "", text: "x" }, { type: "chunk", taskId: "t", text: 1 }, { type: "start", taskId: "t", options: [] }, { type: "wat", taskId: "t" }]) {
      try { validateCSVWorkerRequest(value); throw new Error("expected validation failure"); }
      catch (error) { expect(error).toBeInstanceOf(SheetError); expect((error as SheetError).code).toBe("E_VALIDATION"); }
    }
  });
});
