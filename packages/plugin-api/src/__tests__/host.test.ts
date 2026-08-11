import { describe, expect, it } from "vitest";
import { createPluginHost } from "../host.js";

describe("PluginHost lifecycle", () => {
  it("removes all scoped contributions and hooks when a plugin is disposed", async () => {
    const host = createPluginHost();
    const calls: string[] = [];
    await host.use({
      id: "example",
      setup(context) {
        context.commands.registerCommand({ id: "example.command", description: "Example" });
        context.functions.registerFunction({ name: "EXAMPLE", minArgs: 1, maxArgs: 1 });
        context.menus.registerMenuItem({ id: "example.menu", label: "Example", location: "toolbar" });
        context.hooks.onBeforeCommand(({ commandId }) => calls.push(commandId));
      },
    });
    host.emitBeforeCommand({ commandId: "cell.set", source: "api" });
    expect(calls).toEqual(["cell.set"]);
    expect(host.listCommandContributions()).toEqual([{ id: "example.command", description: "Example" }]);

    await host.dispose("example");
    host.emitBeforeCommand({ commandId: "cell.clear", source: "api" });
    expect(calls).toEqual(["cell.set"]);
    expect(host.listCommandContributions()).toEqual([]);
    expect(host.listFunctionContributions()).toEqual([]);
    expect(host.listMenuContributions()).toEqual([]);
  });

  it("rolls back scoped setup registrations when plugin setup rejects", async () => {
    const host = createPluginHost();
    await expect(host.use({
      id: "broken",
      setup(context) {
        context.menus.registerMenuItem({ id: "broken.menu", label: "Broken", location: "menu" });
        throw new Error("setup failed");
      },
    })).rejects.toThrow("setup failed");
    expect(host.listMenuContributions()).toEqual([]);

    await expect(host.use({ id: "broken", setup() {} })).resolves.toBeUndefined();
  });

  it("isolates observational hook failures from other observers and the host", () => {
    const host = createPluginHost();
    const seen: string[] = [];
    host.hooks.onBeforeCommand(() => { throw new Error("observer failed"); });
    host.hooks.onBeforeCommand(({ commandId }) => seen.push(commandId));
    host.hooks.onWorkbookLoaded(() => { throw new Error("observer failed"); });
    host.hooks.onWorkbookLoaded((workbookId) => seen.push(workbookId));

    expect(() => host.emitBeforeCommand({ commandId: "cell.set", source: "api" })).not.toThrow();
    expect(() => host.emitWorkbookLoaded("workbook")).not.toThrow();
    expect(seen).toEqual(["cell.set", "workbook"]);
  });

  it("rejects command ids reserved by the embedding runtime", async () => {
    const host = createPluginHost({ reservedCommandIds: ["cell.set"] });
    await expect(host.use({
      id: "collision",
      setup(context) { context.commands.registerCommand({ id: "cell.set" }); },
    })).rejects.toMatchObject({ code: "E_VALIDATION" });
    expect(host.listCommandContributions()).toEqual([]);
  });
});
