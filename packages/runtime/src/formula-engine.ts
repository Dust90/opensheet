// M3 runtime integration: per-workbook formula engine state.
//
// Owns: per-sheet DependencyGraph + parsedFormula cache, and per-call
// evaluation budget. The beforeCommit hook (registered by createOpenSheet)
// reconciles the graph for the transaction's pending changes and recomputes
// ONLY the affected formulas via DerivedWriter.
//
// Recalc order per transaction (M3.5 / M3.6 guardrails):
//   1. directly changed formula cells are re-registered + evaluated first
//      (their value comes from THEIR OWN formula, even though they are also
//      "roots" of downstream cells);
//   2. cycle members are written #CYCLE! BEFORE the downstream pass, so the
//      downstream formulas naturally propagate the cycle error;
//   3. downstream formulas run in topological (dependency-first) order,
//      excluding cells already evaluated in step 1 (no double-evaluation).
//
// Fix 1 (M3.6): each Sheet has its own DependencyGraph and parsedCache.
//   Cross-sheet refs are not supported in M3 but the state is correctly
//   isolated so that Sheet2 recalc never touches Sheet1's graph.

import type { WorkbookView } from "@opensheet/core";
import type { PendingChange } from "@opensheet/commands";
import {
  createDefaultFunctions,
  DependencyGraph,
  evaluateExpr,
  makeBudget,
  parseFormula,
  type FormulaContext,
  type FormulaDependencies,
  type FunctionImpl,
} from "@opensheet/formula-engine";
import type { CellAddress, CellValue } from "@opensheet/shared";

export interface FormulaEngineOptions {
  /** Max cell reads per single formula evaluation (M3.5 guardrail). */
  maxCellReadsPerFormula?: number;
  /**
   * Max cell reads shared across ALL formulas in one transaction (M3.6
   * guardrail). Prevents 1 000 dirty formulas × 1 M reads from blocking the
   * main thread. Default: 5 000 000.
   */
  maxCellReadsPerTransaction?: number;
}

const DEFAULT_MAX_READS_PER_FORMULA = 1_000_000;
const DEFAULT_MAX_READS_PER_TX = 5_000_000;

/**
 * Per-sheet formula state (Fix 1 / M3.6): each Sheet owns its own graph and
 * cache so that setting a formula on Sheet2 never alters Sheet1's graph.
 */
interface SheetFormulaState {
  graph: DependencyGraph;
  parsedCache: Map<string, FormulaDependencies>;
}

export class FormulaEngine {
  private readonly registry = createDefaultFunctions();
  private readonly pluginFunctionNames = new Set<string>();
  private readonly maxReadsPerFormula: number;
  private readonly maxReadsPerTx: number;

  /** Fix 1: per-sheet state. Lazily created on first access. */
  private readonly states = new Map<string, SheetFormulaState>();

  constructor(options?: FormulaEngineOptions) {
    this.maxReadsPerFormula = options?.maxCellReadsPerFormula ?? DEFAULT_MAX_READS_PER_FORMULA;
    this.maxReadsPerTx = options?.maxCellReadsPerTransaction ?? DEFAULT_MAX_READS_PER_TX;
  }

  listFunctionNames(): readonly string[] {
    return this.registry.list();
  }

  registerPluginFunction(name: string, implementation: FunctionImpl): void {
    const key = name.toUpperCase();
    if (this.registry.has(key)) {
      throw new Error(`Formula function already registered: ${key}`);
    }
    this.registry.register(key, implementation);
    this.pluginFunctionNames.add(key);
  }

  unregisterPluginFunction(name: string): boolean {
    const key = name.toUpperCase();
    if (!this.pluginFunctionNames.delete(key)) return false;
    return this.registry.unregister(key);
  }

  // ── Per-sheet state accessors ───────────────────────────────────────────

  private getState(sheetId: string): SheetFormulaState {
    let state = this.states.get(sheetId);
    if (state === undefined) {
      state = { graph: new DependencyGraph(), parsedCache: new Map() };
      this.states.set(sheetId, state);
    }
    return state;
  }

  /** Register (or update) the graph node for a formula cell. */
  setFormula(sheetId: string, row: number, col: number, formula: string): void {
    const state = this.getState(sheetId);
    const { dependencies } = parseFormula(formula);
    state.parsedCache.set(`${row}:${col}`, dependencies);
    state.graph.setFormula(row, col, dependencies, formula);
  }

  /** Remove the graph node (formula overwritten by a literal / cleared). */
  removeFormula(sheetId: string, row: number, col: number): void {
    const state = this.getState(sheetId);
    state.parsedCache.delete(`${row}:${col}`);
    state.graph.removeFormula(row, col);
  }

  hasFormula(sheetId: string, row: number, col: number): boolean {
    return this.states.get(sheetId)?.graph.hasFormula(row, col) ?? false;
  }

  formulaSourceOf(sheetId: string, row: number, col: number): string | undefined {
    return this.states.get(sheetId)?.graph.formulaSourceOf(row, col);
  }

  /** Return formula cells in the given range from the sheet's Graph (Fix 2). */
  graphFormulaCellsInRange(sheetId: string, range: { startRow: number; startCol: number; endRow: number; endCol: number }): import("@opensheet/shared").CellAddress[] {
    return this.states.get(sheetId)?.graph.formulaCellsInRange(range) ?? [];
  }

  // ── Incremental recalculation ──────────────────────────────────────────

  /**
   * Reconcile + recalculate after the transaction's commands have executed,
   * writing results through the derived writer (journaled for rollback).
   *
   * `changedFormulas` carries (sheetId, row, col) triples whose formula
   * SOURCE changed (formula.set, literal overwrite, undo/redo, structure
   * rewrite).
   */
  recalculate(
    workbook: WorkbookView,
    changes: readonly PendingChange[],
    changedFormulas: readonly (CellAddress & { sheetId: string })[],
    derived: { setComputedValue(sheetId: string, row: number, col: number, value: CellValue): void },
  ): void {
    // Fix 7: shared transaction budget across all formula evaluations.
    const txBudget = makeBudget(this.maxReadsPerTx);

    // 0. Reconcile graph nodes for formulas whose source changed, grouped by sheet.
    this.reconcileFormulas(workbook, changedFormulas);

    // Group changes by sheetId so each sheet's graph is processed independently.
    const changesBySheet = new Map<string, readonly PendingChange[]>();
    for (const change of changes) {
      const existing = changesBySheet.get(change.sheetId) ?? [];
      changesBySheet.set(change.sheetId, [...existing, change]);
    }

    for (const [sheetId, sheetChanges] of changesBySheet) {
      const sheetView = workbook.getSheetView(sheetId);
      const state = this.getState(sheetId);

      const ctx: FormulaContext = {
        getCellValue: (ref) => sheetView.getCell(ref.row, ref.col)?.value ?? null,
      };

      // 1. Classify affected cells via range-level queries (no expansion).
      const { directlyChangedFormulas, order, cyclic } = state.graph.topoOrderForChanges(sheetChanges);

      // 1b. Topologically order directly changed formulas among themselves.
      const directTopo = state.graph.topoOrderForCells(directlyChangedFormulas);

      // 2. Write #CYCLE! FIRST (Fix 3 already correct in incremental path).
      for (const cell of [...cyclic, ...directTopo.cyclic]) {
        derived.setComputedValue(sheetId, cell.row, cell.col, {
          type: "#CYCLE!",
          message: "Circular reference",
        });
      }
      const cycleKeys = new Set([
        ...cyclic.map((c) => `${c.row}:${c.col}`),
        ...directTopo.cyclic.map((c) => `${c.row}:${c.col}`),
      ]);

      // 3. Evaluate directly changed formulas in dependency order.
      const directKeys = new Set<string>();
      for (const cell of directTopo.order) {
        if (cycleKeys.has(`${cell.row}:${cell.col}`)) continue;
        directKeys.add(`${cell.row}:${cell.col}`);
        const value = this.evaluateFormulaAt(workbook, sheetId, cell.row, cell.col, ctx, txBudget);
        derived.setComputedValue(sheetId, cell.row, cell.col, value);
      }

      // 4. Downstream formulas in dependency order — Fix 6: skip already-evaluated cells.
      for (const cell of order) {
        const key = `${cell.row}:${cell.col}`;
        if (cycleKeys.has(key)) continue; // already written as #CYCLE!
        if (directKeys.has(key)) continue; // Fix 6: already evaluated in step 3
        const value = this.evaluateFormulaAt(workbook, sheetId, cell.row, cell.col, ctx, txBudget);
        derived.setComputedValue(sheetId, cell.row, cell.col, value);
      }
    }
  }

  /** Rebuild the graph from scratch (Snapshot load) and recalc every formula. */
  rebuildAndRecalculateAll(
    workbook: WorkbookView,
    derived: { setComputedValue(sheetId: string, row: number, col: number, value: CellValue): void },
  ): void {
    // Clear all per-sheet state.
    this.states.clear();

    // Fix 1: iterate ALL sheets, not just the first.
    for (const sheet of workbook.listSheetViews()) {
      const sheetId = sheet.id;
      const state = this.getState(sheetId);
      const ctx: FormulaContext = {
        getCellValue: (ref) => sheet.getCell(ref.row, ref.col)?.value ?? null,
      };

      for (const [row, col, data] of sheet.cellEntries()) {
        if (data.formula !== undefined) {
          const { dependencies } = parseFormula(data.formula);
          state.parsedCache.set(`${row}:${col}`, dependencies);
          state.graph.setFormula(row, col, dependencies, data.formula);
        }
      }

      const { order, cyclic } = state.graph.topoOrderAll();

      // Fix 3: write #CYCLE! BEFORE evaluating downstream formulas, so
      // downstream reads see the cycle error rather than stale cached values.
      for (const cell of cyclic) {
        derived.setComputedValue(sheetId, cell.row, cell.col, { type: "#CYCLE!", message: "Circular reference" });
      }
      const cycleKeys = new Set(cyclic.map((c) => `${c.row}:${c.col}`));

      const txBudget = makeBudget(this.maxReadsPerTx);
      for (const cell of order) {
        if (cycleKeys.has(`${cell.row}:${cell.col}`)) continue;
        const value = this.evaluateFormulaAt(workbook, sheetId, cell.row, cell.col, ctx, txBudget);
        derived.setComputedValue(sheetId, cell.row, cell.col, value);
      }
    }
  }

  /**
   * Rebuild the dependency graph for a single sheet from scratch. Called by
   * the beforeCommit hook after structural commands (insertRows / deleteRows /
   * insertColumns / deleteColumns) so that formula coordinates are correct
   * after the store has already been mutated.
   */
  rebuildSheetGraph(workbook: WorkbookView, sheetId: string): void {
    const state = this.getState(sheetId);
    state.graph.removeAll();
    state.parsedCache.clear();
    const sheet = workbook.getSheetView(sheetId);
    for (const [row, col, data] of sheet.cellEntries()) {
      if (data.formula !== undefined) {
        const { dependencies } = parseFormula(data.formula);
        state.parsedCache.set(`${row}:${col}`, dependencies);
        state.graph.setFormula(row, col, dependencies, data.formula);
      }
    }
  }

  private reconcileFormulas(
    workbook: WorkbookView,
    changed: readonly (CellAddress & { sheetId: string })[],
  ): void {
    for (const cell of changed) {
      const sheet = workbook.getSheetView(cell.sheetId);
      const data = sheet.getCell(cell.row, cell.col);
      if (data?.formula !== undefined) {
        this.setFormula(cell.sheetId, cell.row, cell.col, data.formula);
      } else {
        this.removeFormula(cell.sheetId, cell.row, cell.col);
      }
    }
  }

  private evaluateFormulaAt(
    workbook: WorkbookView,
    sheetId: string,
    row: number,
    col: number,
    ctx: FormulaContext,
    txBudget?: ReturnType<typeof makeBudget>,
  ): CellValue {
    if (txBudget !== undefined && txBudget.remaining <= 0) {
      return { type: "#VALUE!", message: "Transaction cell-read limit exceeded" };
    }
    const cell = workbook.getSheetView(sheetId).getCell(row, col);
    const formula = cell?.formula;
    if (formula === undefined) return null;
    const { ast } = parseFormula(formula);
    // Per-formula budget: limit one formula's reads independently.
    const formulaBudget = makeBudget(this.maxReadsPerFormula);
    // Wrap txBudget so both budgets tick down from a single consume call.
    const budget: ReturnType<typeof makeBudget> = {
      maxCellReads: formulaBudget.maxCellReads,
      get remaining() { return Math.min(formulaBudget.remaining, txBudget?.remaining ?? Infinity); },
      consume: (n = 1) => {
        if (txBudget !== undefined && !txBudget.consume(n)) return false;
        return formulaBudget.consume(n);
      },
    };
    return evaluateExpr(ast, ctx, this.registry, budget);
  }
}
