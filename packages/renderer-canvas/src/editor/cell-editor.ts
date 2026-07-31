// CellEditor: DOM textarea overlay for inline cell editing.
//
// Responsibilities (M2 semantics):
//   - Position itself over the edited cell (absolute CSS coords provided by
//     the grid via coordinate-mapper layout).
//   - Own the IME lifecycle: while composing, Enter never commits.
//   - Emit only *requests* (onCommit / onCancel / onMove) — it never touches
//     worksheet data; the host routes commits through runtime.applyOperations.

import { decideKeyInPhase, type EditorKeyInfo } from "./editor-state.js";

export interface CellEditorCallbacks {
  /** Commit the current text. Host writes it via the Command Bus.
   *  `move` (from Enter/Tab) is non-null when the selection should advance. */
  onCommit: (text: string, move: { row: number; col: number } | null) => void;
  /** Abandon the edit (no history). */
  onCancel: () => void;
}

export interface CellEditorRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

const EDITOR_PADDING = 4;

export class CellEditor {
  private readonly host: HTMLElement;
  private readonly callbacks: CellEditorCallbacks;
  private readonly textarea: HTMLTextAreaElement;
  private composing = false;
  private active = false;

  constructor(host: HTMLElement, callbacks: CellEditorCallbacks) {
    this.host = host;
    this.callbacks = callbacks;
    this.textarea = document.createElement("textarea");
    this.textarea.setAttribute("data-testid", "cell-editor");
    this.textarea.style.position = "absolute";
    this.textarea.style.boxSizing = "border-box";
    this.textarea.style.padding = `${EDITOR_PADDING}px 6px`;
    this.textarea.style.margin = "0";
    this.textarea.style.border = "2px solid #0b57d0";
    this.textarea.style.borderRadius = "2px";
    this.textarea.style.outline = "none";
    this.textarea.style.font = '13px -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif';
    this.textarea.style.color = "#1f2329";
    this.textarea.style.background = "#ffffff";
    this.textarea.style.resize = "none";
    this.textarea.style.overflow = "hidden";
    this.textarea.style.whiteSpace = "pre";
    this.textarea.style.display = "none";
    this.textarea.spellcheck = false;
    this.host.append(this.textarea);

    this.textarea.addEventListener("compositionstart", this.handleCompositionStartBound);
    this.textarea.addEventListener("compositionend", this.handleCompositionEndBound);
    this.textarea.addEventListener("keydown", this.handleKeyDownBound);
    this.textarea.addEventListener("blur", this.handleBlurBound);
    // Composition events also fire with native keydown ("Process") — the
    // keydown handler must not act during composition, which decideKeyInPhase
    // enforces via the phase.
  }

  get isActive(): boolean {
    return this.active;
  }

  get isComposing(): boolean {
    return this.composing;
  }

  /** Show the editor over `rect` (canvas CSS px) pre-filled with `text`. */
  open(rect: CellEditorRect, text: string): void {
    this.textarea.style.left = `${rect.x}px`;
    this.textarea.style.top = `${rect.y}px`;
    this.textarea.style.width = `${Math.max(24, rect.width)}px`;
    this.textarea.style.height = `${Math.max(20, rect.height)}px`;
    this.textarea.value = text;
    this.textarea.style.display = "block";
    this.active = true;
    this.composing = false;
    this.textarea.focus();
    this.textarea.setSelectionRange(text.length, text.length);
  }

  commit(move: { row: number; col: number } | null = null): void {
    if (!this.active) return;
    const text = this.textarea.value;
    this.hide();
    this.callbacks.onCommit(text, move);
  }

  cancel(): void {
    if (!this.active) return;
    this.hide();
    this.callbacks.onCancel();
  }

  destroy(): void {
    this.textarea.removeEventListener("compositionstart", this.handleCompositionStartBound);
    this.textarea.removeEventListener("compositionend", this.handleCompositionEndBound);
    this.textarea.removeEventListener("keydown", this.handleKeyDownBound);
    this.textarea.removeEventListener("blur", this.handleBlurBound);
    this.textarea.remove();
  }

  private hide(): void {
    this.textarea.style.display = "none";
    this.active = false;
    this.composing = false;
  }

  // --- input handlers (bound refs so destroy() can remove them) --------------

  private handleCompositionStartBound = (): void => {
    this.composing = true;
  };

  private handleCompositionEndBound = (): void => {
    this.composing = false;
  };

  private handleKeyDownBound = (e: KeyboardEvent): void => {
    const info: EditorKeyInfo = {
      key: e.key,
      shift: e.shiftKey,
      ctrl: e.ctrlKey,
      meta: e.metaKey,
      alt: e.altKey,
    };
    const phase = this.composing ? "composing" : "editing";
    const action = decideKeyInPhase(phase, info);
    switch (action.kind) {
      case "commit": {
        e.preventDefault();
        e.stopPropagation();
        this.commit({ row: action.moveRow, col: action.moveCol });
        break;
      }
      case "commit-stay":
        e.preventDefault();
        this.commit();
        break;
      case "cancel":
        e.preventDefault();
        e.stopPropagation();
        this.cancel();
        break;
      case "start-editing":
      case "none":
        break;
    }
  };

  private handleBlurBound = (): void => {
    if (!this.active) return;
    // Clicking elsewhere in the grid dismisses and commits (Excel-like).
    // The container mousedown (which selected the new cell) already ran; the
    // commit targets the OLD editing cell because commit() reads `this.active`
    // and the grid snapshots the edited cell before blur handlers fire.
    this.commit();
  };
}
