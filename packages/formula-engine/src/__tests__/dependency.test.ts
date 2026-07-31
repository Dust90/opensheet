import { describe, expect, it } from "vitest";
import { DependencyGraph } from "../dependency.js";

const cell = (row: number, col: number) => ({ row, col });

describe("DependencyGraph", () => {
  it("tracks reverse edges and recompute order (dependencies first)", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 1, [cell(0, 0)]); // B1 depends on A1
    g.setFormula(0, 2, [cell(0, 1)]); // C1 depends on B1
    g.setFormula(1, 0, [cell(0, 2)]); // A2 depends on C1

    expect(g.directDependents(0, 0)).toEqual([cell(0, 1)]);
    const { order, cyclic } = g.topoOrder([cell(0, 0)]);
    expect(cyclic).toEqual([]);
    // Chain must be recomputed in dependency order.
    expect(order.map((c) => `${c.row}:${c.col}`)).toEqual(["0:1", "0:2", "1:0"]);
  });

  it("replaces formulas atomically (stale reverse edges removed)", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 1, [cell(0, 0)]);
    g.setFormula(0, 1, [cell(0, 3)]); // B1 now depends on D1
    expect(g.directDependents(0, 0)).toEqual([]);
    expect(g.directDependents(0, 3)).toEqual([cell(0, 1)]);
    expect(g.dependenciesOf(0, 1)).toEqual([cell(0, 3)]);
  });

  it("removes formulas", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 1, [cell(0, 0)]);
    g.removeFormula(0, 1);
    expect(g.size).toBe(0);
    expect(g.directDependents(0, 0)).toEqual([]);
  });

  it("detects cycles and excludes them from recompute order", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 0, [cell(0, 1)]); // A1 ← B1
    g.setFormula(0, 1, [cell(0, 0)]); // B1 ← A1  (cycle)
    g.setFormula(0, 2, [cell(0, 0)]); // C1 depends on A1 (outside the cycle)

    const { order, cyclic } = g.topoOrder([cell(0, 0)]);
    expect(cyclic.map((c) => `${c.row}:${c.col}`).sort()).toEqual(["0:0", "0:1"]);
    // C1 still recomputable; cycle members excluded.
    expect(order.map((c) => `${c.row}:${c.col}`)).toEqual(["0:2"]);
  });

  it("supports self-reference detection (A1 = A1)", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 0, [cell(0, 0)]);
    const { cyclic } = g.topoOrder([cell(0, 0)]);
    expect(cyclic).toHaveLength(1);
    expect(cyclic[0]).toEqual(cell(0, 0));
  });
});
