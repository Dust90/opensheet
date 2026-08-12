// Dependency graph (M3.3/M3.4): formula → referenced cells + range deps,
// reverse edges, and a topological recompute order. Same-sheet only in M3
// phase 1. Ranges are stored as intervals (never expanded), with
// containment queries against changed cells.

import type { CellAddress } from "@injoysai/opensheet-shared";
import type { CellRangeRef } from "./ast.js";
import { rangeBounds } from "./ast.js";

export interface FormulaDependencies {
  /** Individual cell refs (deduplicated). */
  cells: CellAddress[];
  /** Range refs as written — stored as intervals, queried by containment. */
  ranges: CellRangeRef[];
}

function keyOf(row: number, col: number): string {
  return `${row}:${col}`;
}

function addrOf(key: string): CellAddress {
  const [r, c] = key.split(":").map(Number);
  return { row: r!, col: c! };
}

/** Does a range (as written) contain the given cell? */
function rangeContains(range: CellRangeRef, row: number, col: number): boolean {
  const { row1, col1, row2, col2 } = rangeBounds(range);
  return row >= row1 && row <= row2 && col >= col1 && col <= col2;
}

/**
 * Directed graph: formula cell → its dependencies. Reverse edges let the
 * engine find everything that must be recomputed when a source cell changes.
 */
export class DependencyGraph {
  private readonly formulas = new Map<string, FormulaDependencies>(); // cell → deps
  private readonly dependents = new Map<string, Set<string>>(); // dep cell → formula cells
  private readonly formulaSources = new Map<string, string>();

  /** Register (or replace) a formula cell's dependencies. */
  setFormula(row: number, col: number, dependencies: FormulaDependencies, source?: string): void {
    const key = keyOf(row, col);
    // Remove stale reverse edges.
    const previous = this.formulas.get(key);
    if (previous !== undefined) {
      for (const dep of previous.cells) {
        this.dependents.get(keyOf(dep.row, dep.col))?.delete(key);
      }
    }
    this.formulas.set(key, {
      cells: dependencies.cells.map((d) => ({ ...d })),
      ranges: dependencies.ranges.map((r) => ({ start: { ...r.start }, end: { ...r.end } })),
    });
    if (source !== undefined) this.formulaSources.set(key, source);
    // Reverse edges only for individual cell deps; range deps are resolved
    // by containment queries (a huge range must not expand into edges).
    for (const dep of dependencies.cells) {
      const set = this.dependents.get(keyOf(dep.row, dep.col)) ?? new Set<string>();
      set.add(key);
      this.dependents.set(keyOf(dep.row, dep.col), set);
    }
  }

  removeFormula(row: number, col: number): void {
    const key = keyOf(row, col);
    const previous = this.formulas.get(key);
    if (previous !== undefined) {
      for (const dep of previous.cells) {
        this.dependents.get(keyOf(dep.row, dep.col))?.delete(key);
      }
    }
    this.formulas.delete(key);
    this.formulaSources.delete(key);
  }

  /** All formula cells that read the given cell (cell deps + range deps). */
  directDependents(row: number, col: number): CellAddress[] {
    const result = new Set<string>();
    // Individual cell dependencies via reverse edges.
    for (const key of this.dependents.get(keyOf(row, col)) ?? []) {
      result.add(key);
    }
    // Range dependencies via containment scan (formulas count is small).
    for (const [key, deps] of this.formulas) {
      if (result.has(key)) continue;
      for (const range of deps.ranges) {
        if (rangeContains(range, row, col)) {
          result.add(key);
          break;
        }
      }
    }
    return [...result].map(addrOf);
  }

  dependenciesOf(row: number, col: number): FormulaDependencies {
    const deps = this.formulas.get(keyOf(row, col));
    return {
      cells: deps?.cells.map((d) => ({ ...d })) ?? [],
      ranges: deps?.ranges.map((r) => ({ start: { ...r.start }, end: { ...r.end } })) ?? [],
    };
  }

  /** The raw formula source cached for a cell (undefined if not a formula). */
  formulaSourceOf(row: number, col: number): string | undefined {
    return this.formulaSources.get(keyOf(row, col));
  }

  get size(): number {
    return this.formulas.size;
  }

  /**
   * Compute the transitive closure of FORMULA CELLS that depend on the given
   * roots, in recompute order (dependencies before dependents). Returns
   * { order, cyclic } — only the ACTUAL cycle members (the suffix of the DFS
   * stack starting at the back-edge target) are reported as cyclic; cells
   * merely feeding into a cycle stay normal.
   */
  topoOrder(roots: readonly CellAddress[]): { order: CellAddress[]; cyclic: CellAddress[] } {
    const visited = new Set<string>();
    const cyclic = new Set<string>();
    const order: string[] = [];
    const inProgress = new Set<string>();
    const stack: string[] = [];

    const visit = (key: string): void => {
      if (visited.has(key)) return;
      if (inProgress.has(key)) {
        // Back edge: only the DFS-stack suffix starting at the target forms
        // the actual cycle.
        const start = stack.indexOf(key);
        for (let i = start; i < stack.length; i++) cyclic.add(stack[i]!);
        return;
      }
      inProgress.add(key);
      stack.push(key);
      // Traverse REVERSE edges: who reads this cell? (cell deps + range deps
      // resolved by containment — a formula over A1:A1000000 still follows).
      const { row, col } = addrOf(key);
      for (const dependent of this.directDependents(row, col)) {
        visit(keyOf(dependent.row, dependent.col));
      }
      stack.pop();
      inProgress.delete(key);
      visited.add(key);
      order.push(key);
    };

    for (const root of roots) visit(keyOf(root.row, root.col));
    // `order` is dependents-first; reverse for dependency-first recompute
    // order, then drop cyclic members and roots.
    const filtered = order
      .filter((key) => !cyclic.has(key) && !roots.some((r) => keyOf(r.row, r.col) === key))
      .reverse();
    return { order: filtered.map(addrOf), cyclic: [...cyclic].map(addrOf) };
  }

  hasFormula(row: number, col: number): boolean {
    return this.formulas.has(keyOf(row, col));
  }

  /** Remove every formula node (Snapshot reload / full rebuild). */
  removeAll(): void {
    this.formulas.clear();
    this.dependents.clear();
    this.formulaSources.clear();
  }

  /**
   * Return the keys of every formula cell whose coordinate falls inside
   * `range` — used by the sparse-scan Hook to find stale Graph nodes that
   * belong to a changed area without iterating every coordinate.
   */
  formulaCellsInRange(range: { startRow: number; startCol: number; endRow: number; endCol: number }): CellAddress[] {
    const result: CellAddress[] = [];
    for (const key of this.formulas.keys()) {
      const { row, col } = addrOf(key);
      if (row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol) {
        result.push({ row, col });
      }
    }
    return result;
  }

  /**
   * Topological order over EVERY formula (Snapshot-load rebuild): includes
   * all non-cyclic formulas (dependencies first); cycle members listed
   * separately. Unlike topoOrder(), nothing is treated as a "root" to drop.
   */
  topoOrderAll(): { order: CellAddress[]; cyclic: CellAddress[] } {
    return this.topoOrderInternal([...this.formulas.keys()], new Set());
  }

  /**
   * M3.5 runtime entry point: given the transaction's pending change RANGES,
   * find every formula that must be recomputed — WITHOUT expanding the
   * changed areas into per-cell roots. Complexity ≈ O(formulas × ranges).
   *
   * Returns:
   *   directlyChangedFormulas — formula cells inside a changed range (their
   *       formula source was rewritten/replaced; they must be re-registered
   *       and evaluated FIRST).
   *   order — downstream formulas (dependencies before dependents), already
   *       excluding direct cells and cyclic members.
   *   cyclic — actual cycle members (write #CYCLE! before downstream pass).
   */
  topoOrderForChanges(changes: readonly { sheetId: string; range: import("@injoysai/opensheet-shared").Range }[]): {
    directlyChangedFormulas: CellAddress[];
    order: CellAddress[];
    cyclic: CellAddress[];
  } {
    const directly = new Set<string>();
    const roots = new Set<string>();
    const rangeChanged = (row: number, col: number): boolean =>
      changes.some((c) => row >= c.range.startRow && row <= c.range.endRow && col >= c.range.startCol && col <= c.range.endCol);
    const rangeIntersects = (r: CellRangeRef): boolean => {
      const { row1, col1, row2, col2 } = rangeBounds(r);
      return changes.some(
        (c) => row1 <= c.range.endRow && row2 >= c.range.startRow && col1 <= c.range.endCol && col2 >= c.range.startCol,
      );
    };

    for (const [key, deps] of this.formulas) {
      const { row, col } = addrOf(key);
      if (rangeChanged(row, col)) {
        directly.add(key); // the formula cell itself was written
        continue;
      }
      // A formula is a root if any cell dep falls inside a changed range...
      const depHits = deps.cells.some((d) => rangeChanged(d.row, d.col));
      // ...or any range dep intersects a changed range.
      const rangeHits = deps.ranges.some(rangeIntersects);
      if (depHits || rangeHits) roots.add(key);
    }

    const { order, cyclic } = this.topoOrderInternal([...roots], directly);
    return {
      directlyChangedFormulas: [...directly].map(addrOf),
      order,
      cyclic,
    };
  }

  /**
   * Topological order over a specific set of formula cells (e.g. cells whose
   * source was just rewritten in ONE transaction — they may depend on each
   * other). Cycle members among them are reported separately.
   */
  topoOrderForCells(cells: readonly CellAddress[]): { order: CellAddress[]; cyclic: CellAddress[] } {
    return this.topoOrderInternal(cells.map((c) => keyOf(c.row, c.col)), new Set());
  }

  private topoOrderInternal(roots: readonly string[], excluded: Set<string>): { order: CellAddress[]; cyclic: CellAddress[] } {
    const visited = new Set<string>();
    const cyclic = new Set<string>();
    const order: string[] = [];
    const inProgress = new Set<string>();
    const stack: string[] = [];

    const visit = (key: string): void => {
      if (visited.has(key)) return;
      if (inProgress.has(key)) {
        const start = stack.indexOf(key);
        for (let i = start; i < stack.length; i++) cyclic.add(stack[i]!);
        return;
      }
      inProgress.add(key);
      stack.push(key);
      const { row, col } = addrOf(key);
      for (const dependent of this.directDependents(row, col)) {
        visit(keyOf(dependent.row, dependent.col));
      }
      stack.pop();
      inProgress.delete(key);
      visited.add(key);
      // Excluded cells (roots / directly changed) still TRAVERSE so their
      // dependents are found, but they are not emitted here.
      if (!excluded.has(key)) order.push(key);
    };

    for (const root of roots) visit(root);
    const filtered = order.filter((key) => !cyclic.has(key)).reverse();
    return { order: filtered.map(addrOf), cyclic: [...cyclic].map(addrOf) };
  }
}
