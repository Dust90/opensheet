import { isSheetError, SheetError } from "@opensheet/shared";
import { validateCSVWorkerRequest, type CSVWorkerResponse } from "./csv-worker-protocol.js";
import { CSVWorkerTaskHandler } from "./csv-worker-tasks.js";

export interface CSVWorkerPort {
  postMessage(message: CSVWorkerResponse): void;
  addEventListener(type: "message", listener: (event: MessageEvent<unknown>) => void): void;
  start?(): void;
}

/** Install the thin browser Worker adapter around the pure task state machine. */
export function installCSVWorker(port: CSVWorkerPort, options?: { batchRows?: number }): void {
  const tasks = new CSVWorkerTaskHandler(options);
  port.addEventListener("message", (event) => {
    try {
      validateCSVWorkerRequest(event.data);
      tasks.handle(event.data, (response) => port.postMessage(response));
    } catch (error) {
      const source = typeof event.data === "object" && event.data !== null ? event.data as { taskId?: unknown } : undefined;
      const taskId = typeof source?.taskId === "string" && source.taskId.length > 0 ? source.taskId : "__invalid__";
      if (taskId !== "__invalid__") tasks.abort(taskId);
      const mapped = isSheetError(error) ? error : new SheetError("E_OP_FAILED", error instanceof Error ? error.message : String(error));
      port.postMessage({ type: "error", taskId, code: mapped.code, message: mapped.message });
    }
  });
  port.start?.();
}
