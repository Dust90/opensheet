// In-memory plugin host: registries + hook dispatch + plugin lifecycle.

import { SheetError, type Unsubscribe } from "@opensheet/shared";
import type {
  CommandContribution,
  CommandHookPayload,
  FormulaFunctionContribution,
  MenuItemContribution,
  OpenSheetPlugin,
  OpenSheetPluginContext,
} from "./plugin.js";

/**
 * The host stores plugin contributions and lets the runtime drain them into
 * the real systems (command bus, formula engine, UI). Plugins never touch
 * those systems directly.
 */
export class PluginHost implements OpenSheetPluginContext {
  private readonly commandList: CommandContribution[] = [];
  private readonly functionList: FormulaFunctionContribution[] = [];
  private readonly menuList: MenuItemContribution[] = [];
  private readonly beforeCommand = new Set<(p: CommandHookPayload) => void>();
  private readonly afterCommand = new Set<(p: CommandHookPayload) => void>();
  private readonly workbookLoaded = new Set<(workbookId: string) => void>();
  private readonly plugins = new Map<string, OpenSheetPlugin>();

  readonly commands = {
    registerCommand: (contribution: CommandContribution): void => {
      if (this.commandList.some((c) => c.id === contribution.id)) {
        throw new SheetError("E_VALIDATION", `Duplicate command contribution: ${contribution.id}`);
      }
      this.commandList.push(contribution);
    },
  };

  readonly functions = {
    registerFunction: (contribution: FormulaFunctionContribution): void => {
      if (this.functionList.some((f) => f.name === contribution.name)) {
        throw new SheetError("E_VALIDATION", `Duplicate function contribution: ${contribution.name}`);
      }
      this.functionList.push(contribution);
    },
  };

  readonly menus = {
    registerMenuItem: (contribution: MenuItemContribution): void => {
      this.menuList.push(contribution);
    },
  };

  readonly hooks = {
    onBeforeCommand: (cb: (p: CommandHookPayload) => void): Unsubscribe => {
      this.beforeCommand.add(cb);
      return () => this.beforeCommand.delete(cb);
    },
    onAfterCommand: (cb: (p: CommandHookPayload) => void): Unsubscribe => {
      this.afterCommand.add(cb);
      return () => this.afterCommand.delete(cb);
    },
    onWorkbookLoaded: (cb: (workbookId: string) => void): Unsubscribe => {
      this.workbookLoaded.add(cb);
      return () => this.workbookLoaded.delete(cb);
    },
  };

  async use(plugin: OpenSheetPlugin): Promise<void> {
    if (this.plugins.has(plugin.id)) {
      throw new SheetError("E_VALIDATION", `Plugin already installed: ${plugin.id}`);
    }
    await plugin.setup(this);
    this.plugins.set(plugin.id, plugin);
  }

  async dispose(pluginId: string): Promise<void> {
    const plugin = this.plugins.get(pluginId);
    if (plugin === undefined) return;
    await plugin.dispose?.();
    this.plugins.delete(pluginId);
  }

  // --- runtime-facing drains ------------------------------------------------

  listCommandContributions(): readonly CommandContribution[] {
    return this.commandList;
  }

  listFunctionContributions(): readonly FormulaFunctionContribution[] {
    return this.functionList;
  }

  listMenuContributions(): readonly MenuItemContribution[] {
    return this.menuList;
  }

  emitBeforeCommand(payload: CommandHookPayload): void {
    for (const cb of this.beforeCommand) cb(payload);
  }

  emitAfterCommand(payload: CommandHookPayload): void {
    for (const cb of this.afterCommand) cb(payload);
  }

  emitWorkbookLoaded(workbookId: string): void {
    for (const cb of this.workbookLoaded) cb(workbookId);
  }
}

export function createPluginHost(): PluginHost {
  return new PluginHost();
}
