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

import { CommandRegistry } from "./registry.js";
import { cellClearCommand, cellSetCommand, rangeWriteCommand } from "./commands/cells.js";
import { sheetCreateCommand } from "./commands/sheets.js";
import { sheetFreezeCommand } from "./commands/freeze.js";

/** Registry pre-populated with all built-in (M0) commands. */
export function createDefaultRegistry(): CommandRegistry {
  const registry = new CommandRegistry();
  registry.register(cellSetCommand);
  registry.register(cellClearCommand);
  registry.register(rangeWriteCommand);
  registry.register(sheetCreateCommand);
  registry.register(sheetFreezeCommand);
  return registry;
}
