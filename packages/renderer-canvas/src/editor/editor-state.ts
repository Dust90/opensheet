// Editor keystroke state machine — PURE logic, no DOM, unit-testable in node.
//
// The editor has two orthogonal concerns that must NOT be coupled:
//   1. When to start editing (dblclick / F2 / printable char).
//   2. What a key means WHILE editing (commit / cancel / move / nothing).
//
// IME contract (M2 semantics): during an active composition (phase
// "composing") Enter must NOT commit — the composition's final keydown fires
// with key "Process" or similar and the real Enter arrives AFTER
// compositionend, when phase is "editing" again.
//
// Text→primitive inference and display-text formatting live in the shared package.
// shared (infer.ts) so the editor and the clipboard paste path share one
// canonical implementation.

export { cellDisplayText, inferPrimitive } from "@injoysai/opensheet-shared";

export type EditorPhase = "idle" | "editing" | "composing";

export interface EditorKeyInfo {
  key: string;
  shift?: boolean;
  ctrl?: boolean;
  meta?: boolean;
  alt?: boolean;
}

export type EditorAction =
  | { kind: "none" }
  | { kind: "start-editing" }
  | { kind: "commit"; moveRow: number; moveCol: number } // commit + move selection
  | { kind: "commit-stay" }
  | { kind: "cancel" };

/** Printable character that should start inline editing (excludes modifiers). */
export function isPrintableKey(key: string, mods: Pick<EditorKeyInfo, "ctrl" | "meta" | "alt">): boolean {
  if (key.length !== 1 || mods.ctrl || mods.meta || mods.alt) return false;
  // Space is printable but often used for navigation in grids; treat it as
  // printable so typing starts with a space — Excel does the same.
  return true;
}

/** What the grid's keydown handler should do in the CURRENT phase. */
export function decideKeyInPhase(phase: EditorPhase, info: EditorKeyInfo): EditorAction {
  if (phase === "composing") {
    // All keys belong to the composition; the editor DOM receives them
    // natively. Nothing to decide until compositionend.
    return { kind: "none" };
  }
  if (phase === "editing") {
    if (info.ctrl || info.meta || info.alt) return { kind: "none" };
    switch (info.key) {
      case "Enter":
        return { kind: "commit", moveRow: info.shift ? -1 : 1, moveCol: 0 };
      case "Tab":
        return { kind: "commit", moveRow: 0, moveCol: info.shift ? -1 : 1 };
      case "Escape":
        return { kind: "cancel" };
      default:
        return { kind: "none" }; // text input handled natively by the textarea
    }
  }
  // idle
  if (info.ctrl || info.meta || info.alt) return { kind: "none" };
  switch (info.key) {
    case "F2":
      return { kind: "start-editing" };
    case "Enter":
    case "Tab":
    case "Escape":
    case "ArrowUp":
    case "ArrowDown":
    case "ArrowLeft":
    case "ArrowRight":
    case "Home":
    case "End":
    case "PageUp":
    case "PageDown":
    case "Backspace":
    case "Delete":
      return { kind: "none" }; // grid navigation handles these
    default:
      return isPrintableKey(info.key, info) ? { kind: "start-editing" } : { kind: "none" };
  }
}
