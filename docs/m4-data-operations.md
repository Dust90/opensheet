# M4 数据操作契约（排序 / 筛选 / 查找 / 去重）

> 状态：M4 已完成。本文档保留为排序、筛选、查找与去重的语义基准。
> 类型定义见 `packages/shared/src/data-operations.ts`。

## 范围

本轮 M4 只包含：**排序、筛选、查找、去重**。
旧架构文档中的“追加行”移出本轮，避免同时扩大数据重排、视图投影和 API 范围。

子阶段：

| 子阶段 | 内容 |
|---|---|
| M4.0 ✅ | 数据操作契约（本文档）+ `filter`/`reorder` ChangeKind + 同事务预算测试 |
| M4.1 ✅ | RowProjection（视觉行 ↔ 物理行映射，Renderer 全坐标路径接入） |
| M4.2 ✅ | 筛选（compileFilter、filter.apply/clear、Snapshot V2） |
| M4.3 ✅ | 稳定多键排序（公式随行改写、permutation History） |
| M4.4 ✅ | 查找（findCells/findNext、Ctrl+F 面板） |
| M4.5 ✅ | 去重（稳定保留第一条、尾部清空） |
| M4.6 ✅ | 集成 E2E + 性能门槛（100k 行） |

## 冻结的基础语义

以下语义在所有数据操作中一致，不允许各子阶段各自解释：

### 空白

```text
null 代表真正空白
"" 是普通字符串
null !== ""
```

### 类型

```text
数字 1 !== 字符串 "1"
布尔值与数字不进行去重等价
CellError 按 type + message 比较
```

去重 Key 的规范化形式：

```text
null        → null
数字 1      → number:1
字符串 "1"  → string:1
布尔 true   → boolean:true
CellError   → error:<JSON [type,message]>
```

### 比较与匹配

- 筛选条件使用公式的**计算结果** `CellData.value`，不使用公式源码；
- `isBlank` 只匹配 `null`；
- `contains` 将普通值转成显示文本后匹配；
- 数值比较仅在双方都可以安全转换为有限数时成立；
- 错误值默认不匹配普通比较条件；
- 多条件只支持 AND（MVP）。

## ChangeKind 规则

`packages/shared/src/events.ts` 新增两个 kind，语义已接线进 runtime 的
beforeCommit hook（`packages/runtime/src/create-opensheet.ts`）：

| kind | 触发方 | 公式引擎行为 |
|---|---|---|
| `filter` | filter.apply / filter.clear | **不触发重算**——可见性变化不涉及任何值，hook 直接丢弃该 change |
| `reorder` | range.sort / range.dedupe | 目标 Sheet **稀疏重建 Graph**，随后重算受影响公式（与 rows/columns 同路径） |

- Find 不产生 ChangeEvent，也不进入 History；
- 筛选操作本身进入 History（Undo 恢复旧 FilterSpec，Redo 再应用）；
- Sort / Dedupe 各自一次操作一条 History（permutation + 公式原文，不存两份完整 CellData）。
- History 受 `maxMemoryBytes` 限制；大型操作的 Journal 可能被淘汰，随后 Undo 不再可用。性能基准可显式提高该预算以测量 Undo/Redo 算法本身，普通运行时默认仍为 16 MiB。

## Filter 与数据修改的冲突（MVP）

```text
存在活动筛选时：
range.sort 与 range.dedupe 只要其**行跨度**与筛选范围的行跨度相交 → 拒绝
（原子拒绝：零修改、零历史）。列是否重叠不影响该规则：筛选隐藏的是整行，
因此不允许修改任何隐藏行中的其它列。
```

理由：用户看不到隐藏行时，不允许修改隐藏数据。

行列插入/删除影响活动 Filter 时，MVP **直接清除 Filter**（Undo 时恢复），
不在本阶段实现 FilterSpec 坐标改写。

## 契约校验

`data-operations.ts` 导出 `validateSortSpec` / `validateFilterSpec` /
`validateDedupeSpec` / `validateFindOptions`。M4.0.1 起校验器接受 `unknown`
并以断言签名收窄类型（SDK、插件、Snapshot JSON 不受 TypeScript 保护）：

- 输入必须是普通对象，非法输入抛 `SheetError`，绝不抛原生 `TypeError`；
- Range 必须是归一化（start ≤ end）的非负**安全整数**（`Number.isSafeInteger`）；
- `columnOffset` 是**相对 range.startCol 的偏移**（不是绝对列号），必须落在 range 宽度内；
- Sort 至少一个 key 且不允许重复 `columnOffset`；`direction` 枚举校验；`locale` 存在时必须非空字符串；
- Filter 至少一个 condition；`operator` 枚举校验；`caseSensitive` 必须 boolean；
  除 `isBlank` / `notBlank` 外必须携带 `value`，且 `value` 必须是有限数、字符串、布尔或 `null`；
- Dedupe 的 `keep` 必须严格等于 `"first"`；Key 偏移不允许重复；
- Find 的 `query` 必须非空字符串，三个枚举字段（`searchIn`/`scope`/`direction`）逐一校验。

枚举类型全部由常量元组派生（`SORT_DIRECTIONS`、`FILTER_OPERATORS`、`FIND_*`），
类型与运行时校验不会漂移。

## Snapshot

筛选需要刷新后恢复，M4.2 升级为 **Snapshot V2**（`WorksheetSnapshotV2.filter: FilterSpec | null`），
并提供显式迁移 `V1 → V2`（`filter = null`）。不使用可选字段或 `?? null` 弱化严格校验。
