// Dependency graph (M3.3): formula → referenced cells, reverse edges, and
// a topological recompute order. Same-sheet only in M3 phase 1.

import type { CellAddress } from "@opensheet/shared";

export interface FormulaCell {
  row: number;
  col: number;
  /** Cell refs this formula reads (expanded ranges). */
  dependencies: CellAddress[];
}

function keyOf(row: number, col: number): string {
  return `${row}:${col}`;
}

/**
 * Directed graph: formula cell → its dependencies. Reverse edges let the
 * engine find everything that must be recomputed when a source cell changes.
 */
export class DependencyGraph {
  private readonly formulas = new Map<string, CellAddress[]>(); // cell → deps
  private readonly dependents = new Map<string, Set<string>>(); // dep → formula cells

  /** Register (or replace) a formula cell's dependencies. */
  setFormula(row: number, col: number, dependencies: readonly CellAddress[]): void {
    const key = keyOf(row, col);
    // Remove stale reverse edges.
    const previous = this.formulas.get(key);
    if (previous !== undefined) {
      for (const dep of previous) {
        this.dependents.get(keyOf(dep.row, dep.col))?.delete(key);
      }
    }
    this.formulas.set(key, dependencies.map((d) => ({ ...d })));
    for (const dep of dependencies) {
      const set = this.dependents.get(keyOf(dep.row, dep.col)) ?? new Set<string>();
      set.add(key);
      this.dependents.set(keyOf(dep.row, dep.col), set);
    }
  }

  removeFormula(row: number, col: number): void {
    const key = keyOf(row, col);
    const previous = this.formulas.get(key);
    if (previous !== undefined) {
      for (const dep of previous) {
        this.dependents.get(keyOf(dep.row, dep.col))?.delete(key);
      }
    }
    this.formulas.delete(key);
  }

  /** All formula cells that directly read the given cell. */
  directDependents(row: number, col: number): CellAddress[] {
    const set = this.dependents.get(keyOf(row, col));
    if (set === undefined) return [];
    return [...set].map((key) => {
      const [r, c] = key.split(":").map(Number);
      return { row: r!, col: c! };
    });
  }

  dependenciesOf(row: number, col: number): readonly CellAddress[] {
    return this.formulas.get(keyOf(row, col)) ?? [];
  }

  get size(): number {
    return this.formulas.size;
  }

  /**
   * Compute the transitive closure of FORMULA CELLS that depend on the given
   * roots, in recompute order (dependencies before dependents), so a single
   * pass recomputes them safely. Returns { order, cyclic } — `cyclic` lists
   * cells involved in a dependency cycle (recompute those as #CYCLE!).
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
        // Cycle detected: mark the whole in-progress chain as cyclic.
        cyclic.add(key);
        for (const s of stack) cyclic.add(s);
        return;
      }
      inProgress.add(key);
      stack.push(key);
      // Traverse REVERSE edges: who reads this cell?
      for (const dependent of this.dependents.get(key) ?? []) {
        visit(dependent);
      }
      stack.pop();
      inProgress.delete(key);
      visited.add(key);
      order.push(key);
    };

    for (const root of roots) visit(keyOf(root.row, root.col));
    // `order` is dependents-first (deepest formula first); reverse for
    // dependency-first recompute order, then drop cyclic members and roots.
    const filtered = order
      .filter((key) => !cyclic.has(key) && !roots.some((r) => keyOf(r.row, r.col) === key))
      .reverse();
    return { order: filtered.map(addrOf), cyclic: [...cyclic].map(addrOf) };
  }

  hasFormula(row: number, col: number): boolean {
    return this.formulas.has(keyOf(row, col));
  }
}

function addrOf(key: string): CellAddress {
  const [r, c] = key.split(":").map(Number);
  return { row: r!, col: c! };
}
