# ADR-0002: 稀疏运行时模型与 JSON Snapshot 分离

## 状态

已接受（M0 实现）

## 背景

10w 行 × 100 列的逻辑工作表中只有部分单元格有值。为空白单元格预建对象会造成数量级的内存浪费。同时持久化格式必须可 JSON 往返、可版本迁移。

## 决策

- 运行时：`Worksheet` 持有 `CellStore`（稀疏，空单元格无条目）+ `Map<number, number>` 行列尺寸。存储实现经 ADR-0005 benchmark 冻结为分块方案。
- 持久化：`WorkbookSnapshot`（`Record`-based，`cells` 键为 `"row:col"`），`version` 字段支持未来 migration，加载时校验版本。
- 运行时/持久化分离的原因：运行时结构挂载监听器等非序列化状态；Snapshot 必须是纯数据且格式稳定，便于跨版本恢复。
- `CellData.value` 与 `CellData.formula` 分离：渲染永远读 value（最近计算缓存），公式引擎独立维护重算，渲染器与求值器解耦。

## 后果

- 任何序列化都经过 `toWorkbookSnapshot` / `workbookFromSnapshot` 单一通道，属性测试保证 `load(save(x)) ≡ x`。
- 行列插入/删除目前为 O(非空单元格) 的 store 重建，已在 `Worksheet.insertRows` 等方法内隔离，后续可优化为分块位移。
