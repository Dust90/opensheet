// Plugin contracts. plugin-api depends ONLY on @opensheet/shared:
// it never sees runtime, React, or the full OpenSheetAPI.

import type {
  CellPrimitive,
  CellStyle,
  CellValue,
  DedupeSpec,
  FilterSpec,
  SortSpec,
  Unsubscribe,
} from "@opensheet/shared";

/** Built-in operation vocabulary a plugin command may compose atomically. */
export type PluginOperation =
  | { type: "cell.set"; range: string; value: CellPrimitive }
  | { type: "cell.clear"; range: string }
  | { type: "range.write"; range: string; values: CellPrimitive[][] }
  | { type: "formula.set"; range: string; formula: string }
  | { type: "sheet.freeze"; frozenRows: number; frozenColumns: number }
  | { type: "range.style"; range: string; style: Partial<CellStyle> }
  | { type: "row.insert"; at: number; count?: number }
  | { type: "row.delete"; at: number; count?: number }
  | { type: "column.insert"; at: number; count?: number }
  | { type: "column.delete"; at: number; count?: number }
  | { type: "filter.apply"; spec: FilterSpec }
  | { type: "range.sort"; spec: SortSpec }
  | { type: "range.dedupe"; spec: DedupeSpec }
  | { type: "filter.clear" };

/** Read-only physical worksheet view supplied to executable plugin commands. */
export interface PluginCommandContext {
  readonly workbookId: string;
  readonly sheetId: string;
  readonly rowCount: number;
  readonly columnCount: number;
  getCell(row: number, col: number): CellValue;
}

/** Registration metadata for a command contributed by a plugin. */
export interface CommandContribution {
  id: string;
  description?: string;
  /** Optional runtime validator. Throw SheetError(E_VALIDATION) on rejection. */
  validate?(payload: unknown): void;
  /**
   * Return existing operations to execute as one atomic CommandBus batch.
   * Plugins cannot mutate Workbook directly or manufacture Journal entries.
   */
  execute?(context: PluginCommandContext, payload: unknown): readonly PluginOperation[];
}

export interface CommandRegistry {
  registerCommand(contribution: CommandContribution): void;
}

export interface FormulaFunctionContribution {
  name: string;
  minArgs: number;
  maxArgs: number;
  description?: string;
}

export interface FunctionRegistry {
  registerFunction(contribution: FormulaFunctionContribution): void;
}

export type MenuLocation = "menu" | "toolbar" | "context";

export interface MenuItemContribution {
  id: string;
  label: string;
  location: MenuLocation;
  /** Command invoked when the item is activated. */
  commandId?: string;
}

export interface MenuRegistry {
  registerMenuItem(contribution: MenuItemContribution): void;
}

export interface CommandHookPayload {
  commandId: string;
  source: "user" | "api" | "undo" | "redo";
}

export interface PluginHooks {
  onBeforeCommand(callback: (payload: CommandHookPayload) => void): Unsubscribe;
  onAfterCommand(callback: (payload: CommandHookPayload) => void): Unsubscribe;
  onWorkbookLoaded(callback: (workbookId: string) => void): Unsubscribe;
}

export interface OpenSheetPluginContext {
  readonly commands: CommandRegistry;
  readonly functions: FunctionRegistry;
  readonly menus: MenuRegistry;
  readonly hooks: PluginHooks;
}

export interface OpenSheetPlugin {
  readonly id: string;
  setup(context: OpenSheetPluginContext): void | Promise<void>;
  dispose?(): void | Promise<void>;
}
