// Command registry.

import { SheetError } from "@injoysai/opensheet-shared";
import type { SheetCommand } from "./types.js";

/* eslint-disable @typescript-eslint/no-explicit-any */
export class CommandRegistry {
  private readonly commands = new Map<string, SheetCommand<any, any>>();

  register(command: SheetCommand<any, any>): void {
    if (this.commands.has(command.id)) {
      throw new SheetError("E_VALIDATION", `Command already registered: "${command.id}"`);
    }
    this.commands.set(command.id, command);
  }

  get(id: string): SheetCommand<any, any> {
    const command = this.commands.get(id);
    if (command === undefined) {
      throw new SheetError("E_UNKNOWN_COMMAND", `Unknown command: "${id}"`);
    }
    return command;
  }

  has(id: string): boolean {
    return this.commands.has(id);
  }

  list(): readonly string[] {
    return [...this.commands.keys()];
  }
}
