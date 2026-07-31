// M3 runtime integration: per-workbook formula engine state.
//
// Owns: the DependencyGraph, the parsed-formula cache, and the per-call
// evaluation budget. The beforeCommit hook (registered by createOpenSheet)
// reconciles the graph for the transaction's pending changes and recomputes
// ONLY the affected formulas via DerivedWriter.
//
// Recalc order per transaction (M3.5 guardrail):
//   1. directly changed formula cells are re-registered + evaluated first
//      (their value comes from THEIR OWN formula, even though they are also
//      "roots" of downstream cells);
//   2. cycle members are written #CYCLE! BEFORE the downstream pass, so the
//      downstream formulas naturally propagate the cycle error;
//   3. downstream formulas run in topological (dependency-first) order.

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
} from "@opensheet/formula-engine";
import type { CellAddress, CellValue } from "@opensheet/shared";

export interface FormulaEngineOptions {
  /** Max cell reads per single formula evaluation (M3.5 guardrail). */
  maxCellReadsPerFormula?: number;
}

const DEFAULT_MAX_READS = 1_000_000;

export class FormulaEngine {
  readonly graph = new DependencyGraph();
  private readonly registry = createDefaultFunctions();
  private readonly maxReads: number;
  /** cached parsed deps per formula cell (avoid re-parsing every commit). */
  private readonly parsedCache = new Map<string, FormulaDependencies>();

  constructor(options?: FormulaEngineOptions) {
    this.maxReads = options?.maxCellReadsPerFormula ?? DEFAULT_MAX_READS;
  }

  /** Register (or update) the graph node for a formula cell. */
  setFormula(row: number, col: number, formula: string): void {
    const { dependencies } = parseFormula(formula);
    this.parsedCache.set(`${row}:${col}`, dependencies);
    this.graph.setFormula(row, col, dependencies, formula);
  }

  /** Remove the graph node (formula overwritten by a literal / cleared). */
  removeFormula(row: number, col: number): void {
    this.parsedCache.delete(`${row}:${col}`);
    this.graph.removeFormula(row, col);
  }

  /**
   * Reconcile + recalculate after the transaction's commands have executed,
   * writing results through the derived writer (journaled for rollback).
   * `changedFormulas` are cells whose formula SOURCE changed (formula.set,
   * literal overwrite, undo/redo, structure rewrite).
   */
  recalculate(
    workbook: WorkbookView,
    changes: readonly PendingChange[],
    changedFormulas: readonly CellAddress[],
    derived: { setComputedValue(sheetId: string, row: number, col: number, value: CellValue): void },
  ): void {
    // 0. Reconcile graph nodes for formulas whose source changed.
    this.reconcileFormulas(workbook, changedFormulas);

    const sheetId = workbook.listSheetViews()[0]?.id ?? "";
    if (sheetId === "") return;

    // 1. Classify affected cells via range-level queries (no expansion).
    const { directlyChangedFormulas, order, cyclic } = this.graph.topoOrderForChanges(changes);

    // 1b. Directly changed formulas may depend on EACH OTHER (a transaction
    // can formula.set A1 and B1 where B1 reads A1) — topologically order them
    // and detect cycles among them before evaluating.
    const directTopo = this.graph.topoOrderForCells(directlyChangedFormulas);

    const ctx: FormulaContext = {
      getCellValue: (ref) => workbook.getSheetView(sheetId).getCell(ref.row, ref.col)?.value ?? null,
    };

    // 2. Write #CYCLE! for ACTUAL cycle members FIRST (both among directly
    //    changed formulas and in the downstream graph), so every later
    //    evaluation reads the cycle error and propagates it naturally.
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

    // 3. Evaluate directly changed formulas in dependency order (their own
    //    value) — cycle members were already marked, so dependents propagate.
    for (const cell of directTopo.order) {
      if (cycleKeys.has(`${cell.row}:${cell.col}`)) continue;
      const value = this.evaluateFormulaAt(workbook, sheetId, cell.row, cell.col, ctx);
      derived.setComputedValue(sheetId, cell.row, cell.col, value);
    }

    // 4. Downstream formulas in dependency order — propagate the cycle error.
    for (const cell of order) {
      if (cycleKeys.has(`${cell.row}:${cell.col}`)) continue; // already written
      const value = this.evaluateFormulaAt(workbook, sheetId, cell.row, cell.col, ctx);
      derived.setComputedValue(sheetId, cell.row, cell.col, value);
    }
  }

  /** Rebuild the graph from scratch (Snapshot load) and recalc every formula. */
  rebuildAndRecalculateAll(
    workbook: WorkbookView,
    derived: { setComputedValue(sheetId: string, row: number, col: number, value: CellValue): void },
  ): void {
    this.graph.removeAll();
    this.parsedCache.clear();
    const sheet = workbook.listSheetViews()[0];
    if (sheet === undefined) return;
    const sheetId = sheet.id;
    const ctx: FormulaContext = {
      getCellValue: (ref) => sheet.getCell(ref.row, ref.col)?.value ?? null,
    };
    for (const [row, col, data] of sheet.cellEntries()) {
      if (data.formula !== undefined) {
        const { dependencies } = parseFormula(data.formula);
        this.parsedCache.set(`${row}:${col}`, dependencies);
        this.graph.setFormula(row, col, dependencies);
      }
    }
    const { order, cyclic } = this.graph.topoOrderAll();
    for (const cell of order) {
      const value = this.evaluateFormulaAt(workbook, sheetId, cell.row, cell.col, ctx);
      derived.setComputedValue(sheetId, cell.row, cell.col, value);
    }
    for (const cell of cyclic) {
      derived.setComputedValue(sheetId, cell.row, cell.col, { type: "#CYCLE!", message: "Circular reference" });
    }
  }

  private reconcileFormulas(workbook: WorkbookView, changed: readonly CellAddress[]): void {
    const sheet = workbook.listSheetViews()[0];
    if (sheet === undefined) return;
    for (const cell of changed) {
      const data = sheet.getCell(cell.row, cell.col);
      if (data?.formula !== undefined) this.setFormula(cell.row, cell.col, data.formula);
      else this.removeFormula(cell.row, cell.col);
    }
  }

  private evaluateFormulaAt(
    workbook: WorkbookView,
    sheetId: string,
    row: number,
    col: number,
    ctx: FormulaContext,
  ): CellValue {
    const cell = workbook.getSheetView(sheetId).getCell(row, col);
    const formula = cell?.formula;
    if (formula === undefined) return null;
    const { ast } = parseFormula(formula); // cached parse not needed here; parse is cheap
    const budget = makeBudget(this.maxReads);
    return evaluateExpr(ast, ctx, this.registry, budget);
  }
}
