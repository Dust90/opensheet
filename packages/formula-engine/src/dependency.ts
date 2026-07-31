// Dependency graph (M3.3/M3.4): formula → referenced cells + range deps,
// reverse edges, and a topological recompute order. Same-sheet only in M3
// phase 1. Ranges are stored as intervals (never expanded), with
// containment queries against changed cells.

import type { CellAddress } from "@opensheet/shared";
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

  /** Register (or replace) a formula cell's dependencies. */
  setFormula(row: number, col: number, dependencies: FormulaDependencies): void {
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
}
