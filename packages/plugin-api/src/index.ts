// @opensheet/plugin-api — plugin contracts and in-memory host.
// Depends only on @opensheet/shared. No runtime, no React, no OpenSheetAPI.

export type {
  CommandContribution,
  PluginCommandContext,
  PluginOperation,
  CommandHookPayload,
  CommandRegistry,
  FormulaFunctionContribution,
  PluginFormulaArgument,
  PluginFormulaFunction,
  PluginRangeArgument,
  FunctionRegistry,
  MenuItemContribution,
  MenuLocation,
  MenuRegistry,
  OpenSheetPlugin,
  OpenSheetPluginContext,
  PluginHooks,
} from "./plugin.js";
export { createPluginHost, PluginHost } from "./host.js";
export type { PluginHostOptions } from "./host.js";
