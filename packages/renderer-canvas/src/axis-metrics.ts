// AxisMetrics: prefix sums over variable row heights / column widths.
// Position → index via binary search; index → position via prefix lookup.

export type SizeAccessor = (index: number) => number | undefined;

/**
 * One-dimensional layout for one axis (rows or columns).
 * Rebuilt lazily when counts or overrides change; rebuild is O(count),
 * acceptable for 1M rows (~1e6 additions, single-digit ms).
 */
export class AxisMetrics {
  private prefix: Float64Array = new Float64Array(1);
  private resolvedCount = 0;

  constructor(count: number, defaultSize: number, overrideAt: SizeAccessor) {
    this.rebuild(count, defaultSize, overrideAt);
  }

  rebuild(count: number, _defaultSize: number, overrideAt: SizeAccessor): void {
    const prefix = new Float64Array(count + 1);
    let acc = 0;
    for (let i = 0; i < count; i++) {
      prefix[i] = acc;
      acc += overrideAt(i) ?? _defaultSize;
    }
    prefix[count] = acc;
    this.prefix = prefix;
    this.resolvedCount = count;
  }

  get length(): number {
    return this.resolvedCount;
  }

  get totalSize(): number {
    return this.prefix[this.resolvedCount] ?? 0;
  }

  /** Pixel offset where `index` starts (relative to axis origin). */
  positionOf(index: number): number {
    if (index <= 0) return 0;
    if (index >= this.resolvedCount) return this.totalSize;
    return this.prefix[index] ?? 0;
  }

  /** Pixel size of a single index. */
  sizeOf(index: number): number {
    if (index < 0 || index >= this.resolvedCount) return 0;
    return (this.prefix[index + 1] ?? 0) - (this.prefix[index] ?? 0);
  }

  /**
   * Index containing `position` (pixels from axis origin). Binary search over
   * prefix sums. Clamped to [0, count-1]; positions before 0 → 0, beyond the
   * end → count-1.
   */
  indexAt(position: number): number {
    if (this.resolvedCount === 0) return 0;
    if (position <= 0) return 0;
    if (position >= this.totalSize) return this.resolvedCount - 1;
    let lo = 0;
    let hi = this.resolvedCount; // invariant: prefix[lo] <= position < prefix[hi]
    while (lo + 1 < hi) {
      const mid = (lo + hi) >> 1;
      if ((this.prefix[mid] ?? 0) <= position) lo = mid;
      else hi = mid;
    }
    return lo;
  }
}
