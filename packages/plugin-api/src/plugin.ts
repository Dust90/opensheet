// Plugin contracts. plugin-api depends ONLY on @opensheet/shared:
// it never sees runtime, React, or the full OpenSheetAPI.

import type { Unsubscribe } from "@opensheet/shared";

/** Registration metadata for a command contributed by a plugin. */
export interface CommandContribution {
  id: string;
  description?: string;
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
