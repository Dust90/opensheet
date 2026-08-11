import { describe, expect, it } from "vitest";
import { createOpenSheet } from "../create-opensheet.js";

describe("Runtime plugin assembly", () => {
  it("exposes plugin metadata and hooks, then removes them on disposal", async () => {
    const api = createOpenSheet();
    const calls: string[] = [];
    const loaded: string[] = [];
    await api.usePlugin({
      id: "sample",
      setup(context) {
        context.commands.registerCommand({ id: "sample.command" });
        context.functions.registerFunction({ name: "SAMPLE", minArgs: 0, maxArgs: 0 });
        context.menus.registerMenuItem({ id: "sample.menu", label: "Sample", location: "toolbar" });
        context.hooks.onBeforeCommand(({ commandId }) => calls.push(`before:${commandId}`));
        context.hooks.onAfterCommand(({ commandId }) => calls.push(`after:${commandId}`));
        context.hooks.onWorkbookLoaded((workbookId) => loaded.push(workbookId));
      },
    });
    const workbook = api.createWorkbook({ name: "Book" });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      atomic: true,
      operations: [{ type: "cell.set", range: "A1", value: 1 }],
    });
    expect(loaded).toEqual([workbook.id]);
    expect(calls).toEqual(["before:applyOperations", "after:applyOperations"]);
    expect(api.getPluginContributions()).toEqual({
      commands: [{ id: "sample.command" }],
      functions: [{ name: "SAMPLE", minArgs: 0, maxArgs: 0 }],
      menus: [{ id: "sample.menu", label: "Sample", location: "toolbar" }],
    });

    api.createSheet({ name: "Second" });
    api.undo();
    expect(calls).toEqual([
      "before:applyOperations", "after:applyOperations",
      "before:sheet.create", "after:sheet.create",
      "before:history.undo", "after:history.undo",
    ]);

    await api.disposePlugin("sample");
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      operations: [{ type: "cell.set", range: "A2", value: 2 }],
    });
    expect(calls).toEqual([
      "before:applyOperations", "after:applyOperations",
      "before:sheet.create", "after:sheet.create",
      "before:history.undo", "after:history.undo",
    ]);
    expect(api.getPluginContributions()).toEqual({ commands: [], functions: [], menus: [] });
  });

  it("does not let observer failures change Runtime command success", async () => {
    const api = createOpenSheet();
    await api.usePlugin({
      id: "broken-observer",
      setup(context) {
        context.hooks.onBeforeCommand(() => { throw new Error("before failed"); });
        context.hooks.onAfterCommand(() => { throw new Error("after failed"); });
        context.hooks.onWorkbookLoaded(() => { throw new Error("load failed"); });
      },
    });
    const workbook = api.createWorkbook({ name: "Book" });
    await expect(api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      operations: [{ type: "cell.set", range: "A1", value: "once" }],
    })).resolves.toMatchObject({ status: "completed" });
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "A1" })).toEqual([["once"]]);
  });

  it("executes plugin operations atomically through one History batch", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    const hooks: string[] = [];
    await api.usePlugin({
      id: "increment",
      setup(context) {
        context.commands.registerCommand({
          id: "increment.cells",
          validate(payload) {
            if (typeof payload !== "object" || payload === null || !Number.isFinite((payload as { by?: unknown }).by)) {
              throw new Error("invalid increment payload");
            }
          },
          execute(command, payload) {
            const by = (payload as { by: number }).by;
            const current = command.getCell(0, 0);
            return [
              { type: "cell.set", range: "A1", value: typeof current === "number" ? current + by : by },
              { type: "cell.set", range: "B1", value: "plugin" },
            ];
          },
        });
        context.hooks.onBeforeCommand(({ commandId }) => hooks.push(`before:${commandId}`));
        context.hooks.onAfterCommand(({ commandId }) => hooks.push(`after:${commandId}`));
      },
    });

    await api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      operations: [{ type: "cell.set", range: "A1", value: 2 }],
    });
    await expect(api.executePluginCommand({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      commandId: "increment.cells",
      payload: { by: 3 },
    })).resolves.toMatchObject({ status: "completed" });
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "A1:B1" })).toEqual([[5, "plugin"]]);
    api.undo();
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "A1:B1" })).toEqual([[2, null]]);
    expect(hooks).toContain("before:increment.cells");
    expect(hooks).toContain("after:increment.cells");
  });

  it("rolls back every plugin operation when a later generated operation is invalid", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    await api.usePlugin({
      id: "broken-command",
      setup(context) {
        context.commands.registerCommand({
          id: "broken.command",
          execute() {
            return [
              { type: "cell.set", range: "A1", value: "transient" },
              { type: "cell.set", range: "ZZ100000", value: "invalid" },
            ];
          },
        });
      },
    });

    await expect(api.executePluginCommand({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      commandId: "broken.command",
      payload: null,
    })).rejects.toMatchObject({ errorCode: "E_INVALID_RANGE" });
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "A1" })).toEqual([[null]]);
  });

  it("maps unexpected plugin handler errors to SheetError without running operations", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    await api.usePlugin({
      id: "throwing-command",
      setup(context) {
        context.commands.registerCommand({
          id: "throwing.command",
          execute() { throw new Error("plugin failed"); },
        });
      },
    });
    await expect(api.executePluginCommand({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      commandId: "throwing.command",
      payload: null,
    })).rejects.toMatchObject({ code: "E_OP_FAILED", message: "plugin failed" });
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "A1" })).toEqual([[null]]);
  });

  it("rejects plugins attempting to claim a built-in command id", async () => {
    const api = createOpenSheet();
    await expect(api.usePlugin({
      id: "collision",
      setup(context) { context.commands.registerCommand({ id: "cell.set" }); },
    })).rejects.toMatchObject({ code: "E_VALIDATION" });
    expect(api.getPluginContributions().commands).toEqual([]);
  });

  it("installs and removes pure plugin formula functions across existing workbooks", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      atomic: true,
      operations: [
        { type: "cell.set", range: "A1", value: 2 },
        { type: "cell.set", range: "A2", value: 3 },
        { type: "formula.set", range: "B1", formula: "=PLUGIN_SUM(A1:A2)" },
      ],
    });
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "B1" })[0]![0]).toMatchObject({ type: "#NAME?" });

    await api.usePlugin({
      id: "formula-plugin",
      setup(context) {
        context.functions.registerFunction({
          name: "plugin_sum",
          minArgs: 1,
          maxArgs: 1,
          execute(args) {
            const argument = args[0]!;
            if (
              typeof argument !== "object" ||
              argument === null ||
              !("kind" in argument) ||
              argument.kind !== "range"
            ) {
              return { type: "#VALUE!", message: "range required" };
            }
            let total = 0;
            for (const value of argument.values()) {
              if (typeof value === "number") total += value;
            }
            return total;
          },
        });
      },
    });
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "B1" })).toEqual([[5]]);

    // Formula registries are per-workbook. A workbook created after plugin
    // installation must receive the same executable contribution.
    const laterWorkbook = api.createWorkbook({ name: "Later" });
    await api.applyOperations({
      workbookId: laterWorkbook.id,
      sheetId: laterWorkbook.activeSheetId,
      atomic: true,
      operations: [
        { type: "cell.set", range: "A1", value: 4 },
        { type: "cell.set", range: "A2", value: 5 },
        { type: "formula.set", range: "B1", formula: "=PLUGIN_SUM(A1:A2)" },
      ],
    });
    expect(api.readRange({ sheetId: laterWorkbook.activeSheetId, range: "B1" })).toEqual([[9]]);

    await api.disposePlugin("formula-plugin");
    expect(api.readRange({ sheetId: laterWorkbook.activeSheetId, range: "B1" })[0]![0]).toMatchObject({ type: "#NAME?" });
  });

  it("maps unexpected formula function failures and non-finite returns to stable cell errors", async () => {
    const api = createOpenSheet();
    const workbook = api.createWorkbook({ name: "Book" });
    await api.usePlugin({
      id: "error-functions",
      setup(context) {
        context.functions.registerFunction({
          name: "THROWING",
          minArgs: 0,
          maxArgs: 0,
          execute() { throw new Error("boom"); },
        });
        context.functions.registerFunction({
          name: "INFINITE",
          minArgs: 0,
          maxArgs: 0,
          execute() { return Infinity; },
        });
        context.functions.registerFunction({
          name: "REFERR",
          minArgs: 0,
          maxArgs: 0,
          execute() { return { type: "#REF!" }; },
        });
      },
    });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId: workbook.activeSheetId,
      atomic: true,
      operations: [
        { type: "formula.set", range: "A1", formula: "=THROWING()" },
        { type: "formula.set", range: "A2", formula: "=INFINITE()" },
        { type: "formula.set", range: "A3", formula: "=REFERR()" },
      ],
    });
    expect(api.readRange({ sheetId: workbook.activeSheetId, range: "A1:A3" })).toEqual([
      [{ type: "#VALUE!", message: "Plugin function THROWING failed" }],
      [{ type: "#NUM!", message: "Plugin function returned a non-finite number" }],
      [{ type: "#REF!" }],
    ]);
  });

  it("rejects plugin functions that collide with built-in names", async () => {
    const api = createOpenSheet();
    await expect(api.usePlugin({
      id: "function-collision",
      setup(context) {
        context.functions.registerFunction({ name: "sum", minArgs: 0, maxArgs: 0, execute: () => 0 });
      },
    })).rejects.toMatchObject({ code: "E_VALIDATION" });
    expect(api.getPluginContributions().functions).toEqual([]);
  });

  it("rejects plugins attempting to claim lazy special-form names", async () => {
    const api = createOpenSheet();
    for (const name of ["if", "AND", "or"]) {
      await expect(api.usePlugin({
        id: `special-${name}`,
        setup(context) {
          context.functions.registerFunction({ name, minArgs: 0, maxArgs: 0, execute: () => 123 });
        },
      })).rejects.toMatchObject({ code: "E_VALIDATION" });
    }
    expect(api.getPluginContributions().functions).toEqual([]);
  });
});
