// @opensheet/plugin-api — plugin contracts and in-memory host.
// Depends only on @opensheet/shared. No runtime, no React, no OpenSheetAPI.

export type {
  CommandContribution,
  CommandHookPayload,
  CommandRegistry,
  FormulaFunctionContribution,
  FunctionRegistry,
  MenuItemContribution,
  MenuLocation,
  MenuRegistry,
  OpenSheetPlugin,
  OpenSheetPluginContext,
  PluginHooks,
} from "./plugin.js";
export { createPluginHost, PluginHost } from "./host.js";
