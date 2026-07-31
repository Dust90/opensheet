# 数据模型

## 运行时模型（内存）

```ts
interface CellData {
  value: CellValue;          // 字面值或最近一次公式计算缓存（含 CellError）
  formula?: string;          // 公式源码，含前导 "="
  styleId?: string;          // → Workbook.styles
  numberFormat?: string;
}

class Worksheet {
  id; name; rowCount; columnCount;
  frozenRows; frozenColumns;
  rowHeights: Map<number, number>;
  columnWidths: Map<number, number>;
  // cells: CellStore（稀疏，ADR-0005 冻结为 128×128 分块）
}

class Workbook {
  id; name; version;
  styles: StyleTable;        // 哈希去重（ADR-0004）
  // sheets: Worksheet[]；activeSheetId
  // 事件总线：beginBatch/endBatch 事务缓冲（ADR-0003）
}
```

要点：

- **稀疏**：空单元格无条目。`CellStore` 接口（`get/set/delete/entries/forEachInRange`）隔离实现，可注入替换。
- **value 与 formula 分离**：渲染只读 `value`；公式引擎（M3）独立重算后经 derived 通道写回 value。
- **样式去重**：单元格仅存 `styleId`，样式表按内容哈希复用。
- **core 无 DOM 依赖**：可在 node 环境直接单测。

## 持久化模型（Snapshot v1）

```ts
interface WorkbookSnapshot {
  id; name; activeSheetId;
  sheets: WorksheetSnapshot[];   // cells: Record<"row:col", CellData>
  styles: Record<string, CellStyle>;
  version: 1;
}
```

- 纯 JSON，可 `JSON.stringify` 往返；`version` 用于未来 migration，加载时拒绝未知版本。
- 两套模型分离的原因：运行时结构挂载监听器等非序列化状态；Snapshot 需格式稳定、可跨版本恢复。
- 往返恒等由 fast-check 属性测试保证（`load(save(x)) ≡ x`）。

## 变更事件

```ts
interface ChangeEvent {
  workbookId; sheetId;
  changes: { range; kind }[];          // cells/style/rows/columns/structure/metadata
  source: "user" | "api" | "undo" | "redo" | "derived";
  batch: boolean;
}
```

- 事务内事件缓冲，提交时按 `sheetId + source` 合并（每个 sheet+source 至多一条；普通写入 + derived 重算会产生两条），渲染器据此计算 Dirty Region。
- `derived` 事件只来自公式重算，永不进入 Undo 历史。
