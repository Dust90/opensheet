import { describe, expect, it } from "vitest";
import { decideKeyInPhase, isPrintableKey, type EditorKeyInfo } from "../editor/editor-state.js";

const key = (partial: Partial<EditorKeyInfo>): EditorKeyInfo => ({
  key: "",
  shift: false,
  ctrl: false,
  meta: false,
  alt: false,
  ...partial,
});

describe("decideKeyInPhase", () => {
  it("idle: F2 starts editing", () => {
    expect(decideKeyInPhase("idle", key({ key: "F2" }))).toEqual({ kind: "start-editing" });
  });

  it("idle: printable char starts editing (replace-mode initial text)", () => {
    expect(decideKeyInPhase("idle", key({ key: "a" }))).toEqual({ kind: "start-editing" });
    expect(decideKeyInPhase("idle", key({ key: "1" }))).toEqual({ kind: "start-editing" });
    expect(decideKeyInPhase("idle", key({ key: " " }))).toEqual({ kind: "start-editing" });
  });

  it("idle: navigation keys stay with the grid", () => {
    for (const k of ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Enter", "Tab", "Home", "End", "PageUp", "PageDown", "Backspace", "Delete"]) {
      expect(decideKeyInPhase("idle", key({ key: k }))).toEqual({ kind: "none" });
    }
  });

  it("idle: modifier chords never start editing", () => {
    expect(decideKeyInPhase("idle", key({ key: "c", meta: true }))).toEqual({ kind: "none" });
    expect(decideKeyInPhase("idle", key({ key: "v", ctrl: true }))).toEqual({ kind: "none" });
    expect(decideKeyInPhase("idle", key({ key: "a", alt: true }))).toEqual({ kind: "none" });
  });

  it("editing: Enter commits and moves down (shift → up)", () => {
    expect(decideKeyInPhase("editing", key({ key: "Enter" }))).toEqual({ kind: "commit", moveRow: 1, moveCol: 0 });
    expect(decideKeyInPhase("editing", key({ key: "Enter", shift: true }))).toEqual({ kind: "commit", moveRow: -1, moveCol: 0 });
  });

  it("editing: Tab commits and moves right (shift → left)", () => {
    expect(decideKeyInPhase("editing", key({ key: "Tab" }))).toEqual({ kind: "commit", moveRow: 0, moveCol: 1 });
    expect(decideKeyInPhase("editing", key({ key: "Tab", shift: true }))).toEqual({ kind: "commit", moveRow: 0, moveCol: -1 });
  });

  it("editing: Escape cancels", () => {
    expect(decideKeyInPhase("editing", key({ key: "Escape" }))).toEqual({ kind: "cancel" });
  });

  it("editing: plain text keys are handled natively by the textarea", () => {
    expect(decideKeyInPhase("editing", key({ key: "x" }))).toEqual({ kind: "none" });
    expect(decideKeyInPhase("editing", key({ key: "ArrowLeft" }))).toEqual({ kind: "none" });
  });

  it("composing: EVERY key is inert — Enter must not commit mid-composition", () => {
    expect(decideKeyInPhase("composing", key({ key: "Enter" }))).toEqual({ kind: "none" });
    expect(decideKeyInPhase("composing", key({ key: "Process" }))).toEqual({ kind: "none" });
    expect(decideKeyInPhase("composing", key({ key: "Escape" }))).toEqual({ kind: "none" });
    expect(decideKeyInPhase("composing", key({ key: "Tab" }))).toEqual({ kind: "none" });
  });
});

describe("isPrintableKey", () => {
  it("single chars without modifiers are printable", () => {
    expect(isPrintableKey("a", { ctrl: false, meta: false, alt: false })).toBe(true);
    expect(isPrintableKey(" ", { ctrl: false, meta: false, alt: false })).toBe(true);
    expect(isPrintableKey("!", { ctrl: false, meta: false, alt: false })).toBe(true);
  });

  it("non-printable keys and modified chords are excluded", () => {
    expect(isPrintableKey("F2", { ctrl: false, meta: false, alt: false })).toBe(false);
    expect(isPrintableKey("ArrowDown", { ctrl: false, meta: false, alt: false })).toBe(false);
    expect(isPrintableKey("c", { ctrl: true, meta: false, alt: false })).toBe(false);
    expect(isPrintableKey("c", { ctrl: false, meta: true, alt: false })).toBe(false);
  });
});
