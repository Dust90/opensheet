// In-memory plugin host: registries + hook dispatch + plugin lifecycle.

import { SheetError, type Unsubscribe } from "@opensheet/shared";
import type {
  CommandContribution,
  CommandHookPayload,
  CommandRegistry,
  FormulaFunctionContribution,
  FunctionRegistry,
  MenuItemContribution,
  MenuRegistry,
  OpenSheetPlugin,
  OpenSheetPluginContext,
  PluginHooks,
} from "./plugin.js";

interface InstalledPlugin {
  plugin: OpenSheetPlugin;
  cleanups: Set<() => void>;
}

export interface PluginHostOptions {
  /** Runtime-reserved ids (for example built-in CommandBus command ids). */
  reservedCommandIds?: readonly string[];
  /** Runtime-reserved formula names (for example built-in functions). */
  reservedFunctionNames?: readonly string[];
}

/**
 * The host stores plugin contributions and lets Runtime expose/drain them.
 * Each installed plugin gets a scoped registration context, so dispose (and a
 * failed async setup) removes its menus, metadata and hooks atomically.
 */
export class PluginHost implements OpenSheetPluginContext {
  private readonly commandList: CommandContribution[] = [];
  private readonly functionList: FormulaFunctionContribution[] = [];
  private readonly menuList: MenuItemContribution[] = [];
  private readonly beforeCommand = new Set<(p: CommandHookPayload) => void>();
  private readonly afterCommand = new Set<(p: CommandHookPayload) => void>();
  private readonly workbookLoaded = new Set<(workbookId: string) => void>();
  private readonly plugins = new Map<string, InstalledPlugin>();
  private readonly installing = new Set<string>();
  private readonly reservedCommandIds: ReadonlySet<string>;
  private readonly reservedFunctionNames: ReadonlySet<string>;

  readonly commands: CommandRegistry;
  readonly functions: FunctionRegistry;
  readonly menus: MenuRegistry;
  readonly hooks: PluginHooks;

  constructor(options?: PluginHostOptions) {
    this.reservedCommandIds = new Set(options?.reservedCommandIds ?? []);
    this.reservedFunctionNames = new Set((options?.reservedFunctionNames ?? []).map((name) => name.toUpperCase()));
    const root = this.contextFor(undefined);
    this.commands = root.commands;
    this.functions = root.functions;
    this.menus = root.menus;
    this.hooks = root.hooks;
  }

  async use(plugin: OpenSheetPlugin): Promise<void> {
    if (this.plugins.has(plugin.id) || this.installing.has(plugin.id)) {
      throw new SheetError("E_VALIDATION", `Plugin already installed: ${plugin.id}`);
    }
    this.installing.add(plugin.id);
    const cleanups = new Set<() => void>();
    try {
      await plugin.setup(this.contextFor(cleanups));
      this.plugins.set(plugin.id, { plugin, cleanups });
    } catch (error) {
      this.runCleanups(cleanups);
      throw error;
    } finally {
      this.installing.delete(plugin.id);
    }
  }

  async dispose(pluginId: string): Promise<void> {
    const installed = this.plugins.get(pluginId);
    if (installed === undefined) return;
    // Preserve the plugin and its contributions if plugin.dispose itself
    // fails; callers may retry instead of seeing a half-disposed plugin.
    await installed.plugin.dispose?.();
    this.runCleanups(installed.cleanups);
    this.plugins.delete(pluginId);
  }

  // --- runtime-facing snapshots -------------------------------------------

  listCommandContributions(): readonly CommandContribution[] {
    return this.commandList.map((contribution) => ({ ...contribution }));
  }

  listFunctionContributions(): readonly FormulaFunctionContribution[] {
    return this.functionList.map((contribution) => ({ ...contribution }));
  }

  listMenuContributions(): readonly MenuItemContribution[] {
    return this.menuList.map((contribution) => ({ ...contribution }));
  }

  emitBeforeCommand(payload: CommandHookPayload): void {
    this.emitSafely(this.beforeCommand, payload);
  }

  emitAfterCommand(payload: CommandHookPayload): void {
    this.emitSafely(this.afterCommand, payload);
  }

  emitWorkbookLoaded(workbookId: string): void {
    this.emitSafely(this.workbookLoaded, workbookId);
  }

  private contextFor(cleanups: Set<() => void> | undefined): OpenSheetPluginContext {
    const own = (cleanup: () => void) => cleanups?.add(cleanup);
    const commands: CommandRegistry = {
      registerCommand: (contribution) => {
        if (typeof contribution.id !== "string" || contribution.id.length === 0) {
          throw new SheetError("E_VALIDATION", "Plugin command contribution requires a non-empty id");
        }
        if (this.reservedCommandIds.has(contribution.id)) {
          throw new SheetError("E_VALIDATION", `Plugin command id is reserved by Runtime: ${contribution.id}`);
        }
        if (contribution.description !== undefined && typeof contribution.description !== "string") {
          throw new SheetError("E_VALIDATION", "Plugin command contribution description must be a string");
        }
        if (contribution.validate !== undefined && typeof contribution.validate !== "function") {
          throw new SheetError("E_VALIDATION", "Plugin command contribution validate must be a function");
        }
        if (contribution.execute !== undefined && typeof contribution.execute !== "function") {
          throw new SheetError("E_VALIDATION", "Plugin command contribution execute must be a function");
        }
        if (this.commandList.some((candidate) => candidate.id === contribution.id)) {
          throw new SheetError("E_VALIDATION", `Duplicate command contribution: ${contribution.id}`);
        }
        const stored = { ...contribution };
        this.commandList.push(stored);
        own(() => this.remove(this.commandList, stored));
      },
    };
    const functions: FunctionRegistry = {
      registerFunction: (contribution) => {
        if (typeof contribution.name !== "string" || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(contribution.name)) {
          throw new SheetError("E_VALIDATION", "Plugin formula function name must be an identifier");
        }
        const name = contribution.name.toUpperCase();
        if (this.reservedFunctionNames.has(name)) {
          throw new SheetError("E_VALIDATION", `Plugin formula function name is reserved by Runtime: ${name}`);
        }
        if (!Number.isSafeInteger(contribution.minArgs) || !Number.isSafeInteger(contribution.maxArgs) || contribution.minArgs < 0 || contribution.maxArgs < contribution.minArgs) {
          throw new SheetError("E_VALIDATION", "Plugin formula function requires normalized non-negative integer argument bounds");
        }
        if (contribution.description !== undefined && typeof contribution.description !== "string") {
          throw new SheetError("E_VALIDATION", "Plugin formula function description must be a string");
        }
        if (contribution.execute !== undefined && typeof contribution.execute !== "function") {
          throw new SheetError("E_VALIDATION", "Plugin formula function execute must be a function");
        }
        if (this.functionList.some((candidate) => candidate.name === name)) {
          throw new SheetError("E_VALIDATION", `Duplicate function contribution: ${name}`);
        }
        const stored = { ...contribution, name };
        this.functionList.push(stored);
        own(() => this.remove(this.functionList, stored));
      },
    };
    const menus: MenuRegistry = {
      registerMenuItem: (contribution) => {
        const stored = { ...contribution };
        this.menuList.push(stored);
        own(() => this.remove(this.menuList, stored));
      },
    };
    const hooks: PluginHooks = {
      onBeforeCommand: (callback) => this.registerHook(this.beforeCommand, callback, own),
      onAfterCommand: (callback) => this.registerHook(this.afterCommand, callback, own),
      onWorkbookLoaded: (callback) => this.registerHook(this.workbookLoaded, callback, own),
    };
    return { commands, functions, menus, hooks };
  }

  private registerHook<T>(set: Set<T>, callback: T, own: (cleanup: () => void) => void): Unsubscribe {
    set.add(callback);
    const unsubscribe = () => set.delete(callback);
    own(unsubscribe);
    return unsubscribe;
  }

  private remove<T>(list: T[], value: T): void {
    const index = list.indexOf(value);
    if (index >= 0) list.splice(index, 1);
  }

  private runCleanups(cleanups: Set<() => void>): void {
    for (const cleanup of [...cleanups].reverse()) cleanup();
    cleanups.clear();
  }

  /** Plugin hooks are observational: an extension must never alter core work. */
  private emitSafely<T>(callbacks: ReadonlySet<(payload: T) => void>, payload: T): void {
    for (const callback of callbacks) {
      try {
        callback(payload);
      } catch {
        // A future host-level reporter may observe this, but hook failures
        // must not make a committed command look like it failed.
      }
    }
  }
}

export function createPluginHost(options?: PluginHostOptions): PluginHost {
  return new PluginHost(options);
}
