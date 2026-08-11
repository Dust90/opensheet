import type { CSVWorkerRequest, CSVWorkerResponse } from "./csv-worker-protocol.js";

/** Minimal transport surface shared by browser Worker and Runtime fallback. */
export interface CSVWorkerTransport {
  postMessage(message: CSVWorkerRequest): void;
  addEventListener(type: "message", listener: (event: MessageEvent<CSVWorkerResponse>) => void): void;
  addEventListener(type: "error" | "messageerror", listener: (event: ErrorEvent) => void): void;
  removeEventListener?(type: "message", listener: (event: MessageEvent<CSVWorkerResponse>) => void): void;
  removeEventListener?(type: "error" | "messageerror", listener: (event: ErrorEvent) => void): void;
  terminate?(): void;
}

/** Create the real browser transport. Callers provide a non-browser fallback. */
export function createBrowserCSVWorker(): CSVWorkerTransport | null {
  if (typeof Worker === "undefined") return null;
  return new Worker(new URL("./csv.worker.js", import.meta.url), { type: "module" }) as unknown as CSVWorkerTransport;
}
