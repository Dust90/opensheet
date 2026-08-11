import { installCSVWorker, type CSVWorkerPort } from "./csv-worker-entry.js";

// This module is a dedicated Worker entry point.  It intentionally contains
// no parsing/business state: csv-worker-entry owns the protocol boundary and
// CSVWorkerTaskHandler owns the task lifecycle.
installCSVWorker(self as unknown as CSVWorkerPort);
