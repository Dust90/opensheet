// Workbook-level style table with hash-based deduplication.

import type { CellStyle } from "@opensheet/shared";

function stableStringify(style: CellStyle): string {
  const keys = Object.keys(style).sort() as (keyof CellStyle)[];
  const parts: string[] = [];
  for (const key of keys) {
    const value = style[key];
    if (value === undefined) continue;
    parts.push(`${key}:${JSON.stringify(value)}`);
  }
  return parts.join("|");
}

/**
 * Deduplicated style registry. Cells reference styles by id, so thousands of
 * identically-styled cells share one CellStyle object.
 */
export class StyleTable {
  private readonly styles = new Map<string, CellStyle>();
  private readonly hashToId = new Map<string, string>();
  private nextId = 1;

  /** Register a style; returns an existing id if an identical style exists. */
  register(style: CellStyle): string {
    const hash = stableStringify(style);
    const existing = this.hashToId.get(hash);
    if (existing !== undefined) return existing;
    const id = `s${this.nextId++}`;
    this.styles.set(id, { ...style });
    this.hashToId.set(hash, id);
    return id;
  }

  get(id: string): Readonly<CellStyle> | undefined {
    return this.styles.get(id);
  }

  get size(): number {
    return this.styles.size;
  }

  toJSON(): Record<string, CellStyle> {
    // Deep copy: callers must never reach internal style objects.
    return Object.fromEntries([...this.styles].map(([id, style]) => [id, structuredClone(style)]));
  }

  /** Replace the entire table content (used by snapshot loading). */
  replaceWith(other: StyleTable): void {
    this.styles.clear();
    this.hashToId.clear();
    for (const [id, style] of other.styles) this.styles.set(id, style);
    for (const [hash, id] of other.hashToId) this.hashToId.set(hash, id);
    this.nextId = other.nextId;
  }

  static fromJSON(record: Record<string, CellStyle>): StyleTable {
    const table = new StyleTable();
    let maxId = 0;
    for (const [id, style] of Object.entries(record)) {
      table.styles.set(id, { ...style });
      table.hashToId.set(stableStringify(style), id);
      const numeric = /^s(\d+)$/.exec(id);
      if (numeric) maxId = Math.max(maxId, Number(numeric[1]));
    }
    table.nextId = maxId + 1;
    return table;
  }
}
