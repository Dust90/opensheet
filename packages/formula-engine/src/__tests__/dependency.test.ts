import { describe, expect, it } from "vitest";
import { DependencyGraph, type FormulaDependencies } from "../dependency.js";

const cell = (row: number, col: number) => ({ row, col });
const deps = (cells: { row: number; col: number }[], ranges: FormulaDependencies["ranges"] = []): FormulaDependencies => ({
  cells: cells.map((c) => ({ ...c })),
  ranges,
});

describe("DependencyGraph", () => {
  it("tracks reverse edges and recompute order (dependencies first)", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 1, deps([cell(0, 0)])); // B1 depends on A1
    g.setFormula(0, 2, deps([cell(0, 1)])); // C1 depends on B1
    g.setFormula(1, 0, deps([cell(0, 2)])); // A2 depends on C1

    expect(g.directDependents(0, 0)).toEqual([cell(0, 1)]);
    const { order, cyclic } = g.topoOrder([cell(0, 0)]);
    expect(cyclic).toEqual([]);
    expect(order.map((c) => `${c.row}:${c.col}`)).toEqual(["0:1", "0:2", "1:0"]);
  });

  it("replaces formulas atomically (stale reverse edges removed)", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 1, deps([cell(0, 0)]));
    g.setFormula(0, 1, deps([cell(0, 3)])); // B1 now depends on D1
    expect(g.directDependents(0, 0)).toEqual([]);
    expect(g.directDependents(0, 3)).toEqual([cell(0, 1)]);
    expect(g.dependenciesOf(0, 1).cells).toEqual([cell(0, 3)]);
  });

  it("removes formulas", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 1, deps([cell(0, 0)]));
    g.removeFormula(0, 1);
    expect(g.size).toBe(0);
    expect(g.directDependents(0, 0)).toEqual([]);
  });

  it("range dependencies resolve by containment, never expanding", () => {
    const g = new DependencyGraph();
    // B1 = SUM(A1:A1000000) — stored as an interval, NOT a million edges.
    g.setFormula(0, 1, deps([], [
      {
        start: { row: 0, col: 0, rowAbs: false, colAbs: false },
        end: { row: 999_999, col: 0, rowAbs: false, colAbs: false },
      },
    ]));

    expect(g.directDependents(500_000, 0)).toEqual([cell(0, 1)]); // inside the range
    expect(g.directDependents(0, 0)).toEqual([cell(0, 1)]); // top-left corner too
    expect(g.directDependents(0, 2)).toEqual([]); // outside the range (col C)
    expect(g.directDependents(0, 1)).toEqual([]); // B1 itself is not read by the range
    expect(g.size).toBe(1); // only one formula stored, no per-cell entries

    const { order } = g.topoOrder([cell(500_000, 0)]);
    expect(order).toEqual([cell(0, 1)]); // the formula recomputes
  });

  it("detects ONLY the actual cycle members (prefix cells stay normal)", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 1, deps([cell(0, 0)])); // B1 = A1  (feeds the cycle)
    g.setFormula(0, 2, deps([cell(0, 1), cell(0, 3)])); // C1 = B1 + D1
    g.setFormula(0, 3, deps([cell(0, 2)])); // D1 = C1   (C1 ↔ D1 cycle)

    const { order, cyclic } = g.topoOrder([cell(0, 0)]);
    expect(cyclic.map((c) => `${c.row}:${c.col}`).sort()).toEqual(["0:2", "0:3"]);
    // B1 is NOT part of the cycle and must stay in the recompute order.
    expect(order.map((c) => `${c.row}:${c.col}`)).toEqual(["0:1"]);
  });

  it("self-reference (A1 = A1) is detected", () => {
    const g = new DependencyGraph();
    g.setFormula(0, 0, deps([cell(0, 0)]));
    const { cyclic } = g.topoOrder([cell(0, 0)]);
    expect(cyclic).toHaveLength(1);
    expect(cyclic[0]).toEqual(cell(0, 0));
  });
});
