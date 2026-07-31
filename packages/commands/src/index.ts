// @opensheet/commands — command bus, registry, transactions, inverse journal.

export { CommandBus } from "./bus.js";
export type { CommandBusOptions, OperationEnvelope } from "./bus.js";
export { CommandRegistry } from "./registry.js";
export type {
  BeforeCommitHook,
  CommandContext,
  CommandOutcome,
  DerivedWriter,
  HistorySink,
  JournalBatch,
  JournalEntry,
  JournalReplayContext,
  PendingChange,
  SheetCommand,
} from "./types.js";
export {
  ApplyOperationsError,
  type ApplyOperationsRequest,
  type ApplyOperationsResult,
  type SheetOperation,
} from "./operations.js";
export { cellClearCommand, cellSetCommand, rangeWriteCommand } from "./commands/cells.js";
export { sheetCreateCommand } from "./commands/sheets.js";
export type { SheetCreatePayload, SheetCreateResult } from "./commands/sheets.js";
export { sheetFreezeCommand } from "./commands/freeze.js";
export type { SheetFreezePayload } from "./commands/freeze.js";
export { rangeStyleCommand } from "./commands/style.js";
export type { RangeStylePayload } from "./commands/style.js";
export {
  columnDeleteCommand,
  columnInsertCommand,
  rowDeleteCommand,
  rowInsertCommand,
} from "./commands/structure.js";
export type { StructurePayload } from "./commands/structure.js";

import { CommandRegistry } from "./registry.js";
import { cellClearCommand, cellSetCommand, rangeWriteCommand } from "./commands/cells.js";
import { sheetCreateCommand } from "./commands/sheets.js";
import { sheetFreezeCommand } from "./commands/freeze.js";
import { rangeStyleCommand } from "./commands/style.js";
import {
  columnDeleteCommand,
  columnInsertCommand,
  rowDeleteCommand,
  rowInsertCommand,
} from "./commands/structure.js";

/** Registry pre-populated with all built-in commands. */
export function createDefaultRegistry(): CommandRegistry {
  const registry = new CommandRegistry();
  registry.register(cellSetCommand);
  registry.register(cellClearCommand);
  registry.register(rangeWriteCommand);
  registry.register(sheetCreateCommand);
  registry.register(sheetFreezeCommand);
  registry.register(rangeStyleCommand);
  registry.register(rowInsertCommand);
  registry.register(rowDeleteCommand);
  registry.register(columnInsertCommand);
  registry.register(columnDeleteCommand);
  return registry;
}
