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
});
