import { isSheetError, SheetError } from "@opensheet/shared";
import { CSVParser } from "./csv.js";
import type { CSVWorkerRequest, CSVWorkerResponse } from "./csv-worker-protocol.js";

export type CSVWorkerEmit = (response: CSVWorkerResponse) => void;

interface Task {
  parser: CSVParser;
  pendingRows: string[][];
  rowCount: number;
}

/** Pure task state machine used by the browser Worker entry and unit tests. */
export class CSVWorkerTaskHandler {
  private readonly tasks = new Map<string, Task>();
  private readonly batchRows: number;

  constructor(options?: { batchRows?: number }) {
    const batchRows = options?.batchRows ?? 512;
    if (!Number.isSafeInteger(batchRows) || batchRows < 1) throw new SheetError("E_VALIDATION", "CSV worker batchRows must be a positive safe integer");
    this.batchRows = batchRows;
  }

  handle(request: CSVWorkerRequest, emit: CSVWorkerEmit): void {
    switch (request.type) {
      case "start": {
        if (this.tasks.has(request.taskId)) return this.error(request.taskId, new SheetError("E_VALIDATION", "CSV worker task already exists"), emit);
        this.tasks.set(request.taskId, { parser: new CSVParser(request.options), pendingRows: [], rowCount: 0 });
        return;
      }
      case "cancel": {
        if (!this.tasks.delete(request.taskId)) this.error(request.taskId, new SheetError("E_VALIDATION", "CSV worker task does not exist"), emit);
        return;
      }
      case "chunk": return this.withTask(request.taskId, emit, (task) => this.append(request.taskId, task, task.parser.push(request.text), emit, false));
      case "finish": return this.withTask(request.taskId, emit, (task) => {
        this.append(request.taskId, task, task.parser.finish(), emit, true);
        emit({ type: "done", taskId: request.taskId, rowCount: task.rowCount });
        this.tasks.delete(request.taskId);
      });
    }
  }

  private withTask(taskId: string, emit: CSVWorkerEmit, action: (task: Task) => void): void {
    const task = this.tasks.get(taskId);
    if (task === undefined) { this.error(taskId, new SheetError("E_VALIDATION", "CSV worker task does not exist"), emit); return; }
    try { action(task); }
    catch (error) { this.tasks.delete(taskId); this.error(taskId, error, emit); }
  }

  private append(taskId: string, task: Task, rows: string[][], emit: CSVWorkerEmit, force: boolean): void {
    task.pendingRows.push(...rows); task.rowCount += rows.length;
    while (task.pendingRows.length >= this.batchRows || (force && task.pendingRows.length > 0)) {
      const count = force ? task.pendingRows.length : this.batchRows;
      emit({ type: "rows", taskId, rows: task.pendingRows.splice(0, count) });
    }
  }

  private error(taskId: string, error: unknown, emit: CSVWorkerEmit): void {
    const mapped = isSheetError(error) ? error : new SheetError("E_OP_FAILED", error instanceof Error ? error.message : String(error));
    emit({ type: "error", taskId, code: mapped.code, message: mapped.message });
  }
}
