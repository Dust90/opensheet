# OpenSheet 接入指南

OpenSheet 是一个 TypeScript 电子表格内核。它负责 Workbook、Worksheet、公式、事务、Undo/Redo、筛选、排序、查找、去重、CSV 和插件；宿主应用负责自己的 UI、Canvas/DOM 渲染、路由与存储策略。

本文以公开组合根 `@injoysai/opensheet` 为准。除编写插件类型时需要的 `@injoysai/opensheet-plugin-api` 外，应用代码应只依赖主包，而不直接依赖 `core`、`commands` 或 `formula-engine` 等实现包。

## 1. 安装

首个发布版本使用 `next` tag 时：

```bash
pnpm add @injoysai/opensheet@next
```

发布为稳定版后可省略 `@next`：

```bash
pnpm add @injoysai/opensheet
```

如果要为插件提供明确的 TypeScript 类型，再添加：

```bash
pnpm add @injoysai/opensheet-plugin-api
```

OpenSheet 面向现代 JavaScript 运行时和浏览器；浏览器 CSV 导入会使用 module Worker。请使用能处理 ESM 与 `new URL(..., import.meta.url)` 的现代打包器（Vite、Webpack 5、Rspack 等）。

## 2. 最小可运行示例

创建一个 Runtime 实例，然后创建 Workbook。一个 Runtime 可以管理多个 Workbook；没有显式传 `workbookId` 的少数查询 API 会使用最近访问的 Workbook，因此应用通常应自行保存当前 Workbook 和 Sheet 的 ID。

```ts
import { createOpenSheet } from "@injoysai/opensheet";

const api = createOpenSheet();

const workbook = api.createWorkbook({ name: "预算" });
const workbookId = workbook.id;
const sheetId = workbook.activeSheetId;

await api.applyOperations({
  workbookId,
  sheetId,
  atomic: true,
  operations: [
    { type: "range.write", range: "A1:B3", values: [
      ["项目", "金额"],
      ["服务", 1200],
      ["软件", 800],
    ] },
    { type: "formula.set", range: "B4", formula: "=SUM(B2:B3)" },
  ],
});

console.log(api.readRange({ sheetId, range: "A1:B4" }));
// [["项目", "金额"], ["服务", 1200], ["软件", 800], [null, 2000]]
```

坐标与范围遵循 A1 语法：`A1` 是第 0 行第 0 列，`A1:C10` 是闭区间。对于 Sort、Filter、Dedupe 的 `range`，使用 0-based 数值坐标。

## 3. 写入、读取与事务

所有 Workbook mutation 都通过 `applyOperations()` 完成。`atomic: true` 表示该调用中的操作要么全部成功、要么全部回滚，并作为一条 History batch 撤销；这是批量写入、先改值再排序等业务操作的推荐默认值。

```ts
await api.applyOperations({
  workbookId,
  sheetId,
  atomic: true,
  operations: [
    { type: "cell.set", range: "C1", value: "状态" },
    { type: "cell.set", range: "C2", value: "已确认" },
    { type: "cell.clear", range: "C3" },
    { type: "range.style", range: "A1:C1", style: { bold: true } },
  ],
});

const values = api.readRange({ sheetId, range: "A1:C3" });
```

支持的值为有限 number、string、boolean 和 `null`。`null` 是唯一的真正空白值；`""` 是普通空字符串，二者在筛选、去重和 CSV used range 中不同。

常用 operation：

| 目的 | operation |
| --- | --- |
| 写一个单元格 | `{ type: "cell.set", range: "A1", value }` |
| 清空单元格 | `{ type: "cell.clear", range: "A1" }` |
| 批量写矩形 | `{ type: "range.write", range: "A1:B2", values }` |
| 设置公式 | `{ type: "formula.set", range: "B2", formula: "=A2*2" }` |
| 插入/删除行列 | `row.insert`、`row.delete`、`column.insert`、`column.delete` |
| 冻结区域 | `{ type: "sheet.freeze", frozenRows, frozenColumns }` |
| 样式 | `{ type: "range.style", range, style }` |

操作校验失败时，`applyOperations()` 会抛出 `ApplyOperationsError`，其中包含失败 operation 的索引和错误码。不要在捕获错误后盲目重试一个已返回成功结果的请求。

## 4. 公式与 Undo/Redo

公式使用以 `=` 开头的 A1 引用语法。公式计算值会在同一个事务提交时自动更新；`readRange()` 返回的是当前 computed value，而不是公式源码。

```ts
await api.applyOperations({
  workbookId,
  sheetId,
  atomic: true,
  operations: [
    { type: "cell.set", range: "A1", value: 21 },
    { type: "formula.set", range: "B1", formula: "=A1*2" },
  ],
});

console.log(api.readRange({ sheetId, range: "B1" })); // [[42]]

api.undo(); // 撤销整个上面的 atomic 调用
api.redo(); // 恢复整个调用
```

History 有内存上限（默认 16 MiB）。大型操作在超过 `maxMemoryBytes` 后可能被淘汰，因此“操作成功”不等价于“任意大小操作永远可撤销”。如有明确产品需求，可在创建 Runtime 时设置更高的预算：

```ts
const api = createOpenSheet({
  history: { maxMemoryBytes: 64 * 1024 * 1024 },
});
```

## 5. 监听变更并接入自己的 UI

Runtime 不绑定 React、Vue 或 Canvas。宿主订阅合并后的 `ChangeEvent`，在回调中更新自己的渲染或状态层：

```ts
const unsubscribe = api.onChange((event) => {
  console.log(event.source, event.changes);
  // 根据 event.changes 刷新 UI、同步状态或使缓存失效。
});

// 组件卸载或实例销毁时：
unsubscribe();
```

`event.source` 可能是 `api`、`user`、`undo`、`redo` 或 `derived`。公式重算产生的 computed value 更新使用 `derived` 通道；不要把它当作新的用户编辑再次写回 Runtime，否则会形成循环。

若要构建自己的 Renderer，可通过 `getWorksheetView(sheetId)` 取得 readonly 视图。不要直接修改它；写入仍必须走 `applyOperations()`。

## 6. 筛选、排序、去重与查找

### 筛选

```ts
await api.applyOperations({
  workbookId,
  sheetId,
  atomic: true,
  operations: [{
    type: "filter.apply",
    spec: {
      range: { startRow: 0, startCol: 0, endRow: 100, endCol: 2 },
      hasHeader: true,
      conditions: [{ columnOffset: 1, operator: "greaterThan", value: 1000 }],
    },
  }],
});

const projection = api.getFilterProjectionState(sheetId);
// projection.filter === null 表示没有 Filter。
// projection.visibleRows 是 active Filter 在其 range 内可见的物理行；
// 空 Uint32Array 表示 Filter 存在但没有匹配行。
```

Filter 是视觉投影，不会重排物理数据。清除使用 `{ type: "filter.clear" }`。当 Filter 激活时，只要 Sort/Dedupe 的行跨度与 Filter 行跨度相交，操作会被拒绝，避免在隐藏行上静默修改数据。

### 稳定排序

```ts
await api.applyOperations({
  workbookId,
  sheetId,
  atomic: true,
  operations: [{
    type: "range.sort",
    spec: {
      range: { startRow: 0, startCol: 0, endRow: 50, endCol: 2 },
      hasHeader: true,
      keys: [
        { columnOffset: 1, direction: "desc" },
        { columnOffset: 0, direction: "asc" },
      ],
      locale: "zh-CN",
    },
  }],
});
```

排序稳定，header 不参与排序；单元格的值、公式、样式与 number format 会随内容移动，行高和冻结设置不会移动。范围内公式会按目标行平移相对引用；范围外公式源码不变，但会重新计算。

### 稳定去重

```ts
await api.applyOperations({
  workbookId,
  sheetId,
  atomic: true,
  operations: [{
    type: "range.dedupe",
    spec: {
      range: { startRow: 0, startCol: 0, endRow: 50, endCol: 2 },
      hasHeader: true,
      keyColumnOffsets: [0, 1],
      keep: "first",
    },
  }],
});
```

去重按原始顺序保留第一次出现的行，并将保留行稳定压紧到上方，尾部单元格清空；Worksheet 的 `rowCount` 不变。`keyColumnOffsets: []` 表示比较整个 range 的所有列。公式单元格按当前 computed value 比较，而不是按公式文本比较。

### 查找

```ts
const matches = api.findCells({
  sheetId,
  query: "服务",
  matchCase: false,
  wholeCell: false,
  searchIn: "values", // 或 "formulas"
  scope: "visible",   // 或 "all"
  direction: "forward",
});

const next = api.findNext({
  sheetId,
  from: { row: 0, col: 0 },
  query: "服务",
  matchCase: false,
  wholeCell: false,
  searchIn: "values",
  scope: "visible",
  direction: "forward",
});
```

`findNext()` 使用物理坐标。它严格寻找当前地址之后（或之前）的结果，到边界后会 wrap；只有一个匹配时会返回自身。`scope: "visible"` 会跳过 Filter 隐藏行，但 Filter range 外的行仍属于可见行。

## 7. CSV 导入与导出

浏览器中可将 `<input type="file">` 的 `File` 直接传给 `importCSV()`：

```ts
const input = document.querySelector<HTMLInputElement>("#csv")!;
const file = input.files?.[0];
if (!file) throw new Error("请选择 CSV 文件");

const imported = await api.importCSV({ file });
console.log(imported);
// { sheetId, rowCount, columnCount }
```

导入总是创建一个新 Worksheet，从 A1 写入；不会覆盖已有 Sheet。文件名会成为 Sheet 名，发生重名时自动加后缀。解析与写入采用 staging，两遍解析任一阶段失败都不会留下半成品 Sheet。一次导入对应一条 History batch，Undo 删除导入的 Sheet，Redo 恢复它。

传入分隔符可导入 TSV 或分号 CSV：

```ts
const imported = await api.importCSV({ file, delimiter: ";" });
```

导出使用 Worksheet 的 used range：从 A1 到最后一个 computed value 非 `null` 的单元格。内部空白会保留，末尾全空行/列会裁剪；`""` 仍是有效值并会扩大 used range，style-only 单元格不会。公式导出的是当前 computed value，错误值导出稳定错误文本，例如 `#DIV/0!`。

```ts
const blob = await api.exportCSV({ sheetId });
const url = URL.createObjectURL(blob);
const anchor = document.createElement("a");
anchor.href = url;
anchor.download = "budget.csv";
anchor.click();
URL.revokeObjectURL(url);
```

CSV 导入在浏览器优先使用 Worker；Worker 创建、加载、协议或解析失败会拒绝 Promise，不会静默回退为一个部分导入的 Sheet。

## 8. 保存与恢复

使用 `getWorkbookSnapshot()` 得到 OpenSheet 原生 Snapshot，再在应用自己的存储层保存。恢复时使用 `loadWorkbook()`；公式缓存会重新计算，不能把 Snapshot 内的公式缓存当作可信数据源。

```ts
localStorage.setItem("workbook", JSON.stringify(api.getWorkbookSnapshot()));

const raw = localStorage.getItem("workbook");
if (raw) {
  const restored = api.loadWorkbook(JSON.parse(raw));
  console.log(restored.activeSheetId);
}
```

CSV 只用于值的交换，不保存公式源码、样式、冻结、Filter、History 或插件状态；需要完整恢复 Workbook 时请使用 Snapshot。

## 9. 编写插件

插件安装在单个 Runtime 实例上。插件可以注册菜单元数据、观察 hook、原子命令和纯同步公式函数。安装/卸载会同步现有与后续 Workbook 的公式引擎。

### 原子插件命令

插件命令不能直接拿到 Workbook、CommandBus 或 History。它接收只读物理 Sheet 视图，并返回现有 operation；Runtime 将这些 operation 作为一个 atomic batch 执行。

```ts
import type { OpenSheetPlugin } from "@injoysai/opensheet-plugin-api";
import { SheetError } from "@injoysai/opensheet";

const fillStatus: OpenSheetPlugin = {
  id: "example.fill-status",
  setup({ commands, menus }) {
    commands.registerCommand({
      id: "example.fill-status.apply",
      description: "写入状态列",
      validate(payload) {
        if (typeof payload !== "object" || payload === null) {
          throw new SheetError("E_VALIDATION", "payload must be an object");
        }
      },
      execute(context) {
        return [{
          type: "cell.set",
          range: "C2",
          value: `rows: ${context.rowCount}`,
        }];
      },
    });
    menus.registerMenuItem({
      id: "example.fill-status.menu",
      label: "填充状态",
      location: "toolbar",
      commandId: "example.fill-status.apply",
    });
  },
};

await api.usePlugin(fillStatus);

await api.executePluginCommand({
  workbookId,
  sheetId,
  commandId: "example.fill-status.apply",
  payload: {},
});

api.undo(); // 撤销插件命令产生的整个 operation batch
await api.disposePlugin("example.fill-status");
```

原子性只覆盖 handler 返回的 `PluginOperation[]` 对 Workbook 的变更。插件自行执行的网络请求、localStorage 写入或其他外部副作用无法由 Runtime 回滚。

### 自定义公式函数

公式函数必须纯同步：不能 `await`，不应修改外部状态，也不应依赖未声明的可变全局状态。参数是 `CellValue` 或惰性 range；仅在需要时遍历 `range.values()`，避免把大范围预先物化。

```ts
const doublePlugin: OpenSheetPlugin = {
  id: "example.double",
  setup({ functions }) {
    functions.registerFunction({
      name: "PLUGIN_DOUBLE",
      minArgs: 1,
      maxArgs: 1,
      description: "将数字乘以二",
      execute: ([value]) => {
        if (typeof value !== "number") return { type: "#VALUE!" };
        return value * 2;
      },
    });
  },
};

await api.usePlugin(doublePlugin);

await api.applyOperations({
  workbookId,
  sheetId,
  atomic: true,
  operations: [
    { type: "cell.set", range: "A1", value: 21 },
    { type: "formula.set", range: "B1", formula: "=PLUGIN_DOUBLE(A1)" },
  ],
});

console.log(api.readRange({ sheetId, range: "B1" })); // [[42]]

await api.disposePlugin("example.double");
// 现有 =PLUGIN_DOUBLE(...) 会重新计算为 #NAME?
```

函数名不区分大小写，且不能覆盖内置函数或 special form（包括 `IF`、`AND`、`OR`）。普通异常会转为 `#VALUE!`，`NaN`/`Infinity` 会转为 `#NUM!`；返回合法 `CellError` 时会保留该错误类型。

观察 hook 仅用于日志、指标或 UI 同步。Runtime 会隔离 observer 异常：hook 抛错不会取消命令，也不会把已成功的命令伪装为失败。

## 10. 错误处理与排查

公开 API 会对运行时输入做校验。常见错误：

| 情况 | 处理建议 |
| --- | --- |
| `SheetError(E_VALIDATION)` | 检查 range、Filter/Sort/Dedupe spec、CSV options 或插件 payload 的形状。 |
| `E_INVALID_RANGE` | 检查 0-based 数值 range 是否超出 Sheet 边界；A1 字符串范围是否写对。 |
| `ApplyOperationsError` | 查看 `failedOperationIndex`、`errorCode`，修复该批次中失败的 operation。 |
| CSV 导入 Promise rejected | CSV 可能有未闭合引号、非法 quote，或 Worker/浏览器策略失败；原 Workbook 不会被部分修改。 |
| 公式为 `#NAME?` | 函数拼写不正确，或对应插件已被卸载。 |

对外部 JSON、插件 payload 和用户输入应始终当作 `unknown` 处理；让 Runtime 的 validator 返回结构化错误，而不要绕过公开 API 修改内部对象。

## 11. 发布前接入检查清单

- [ ] 应用只从 `@injoysai/opensheet` 导入 Runtime API。
- [ ] 每个 mutation 都带正确的 `workbookId`、`sheetId`，批量业务操作使用 `atomic: true`。
- [ ] UI 通过 `onChange()` 失效/刷新，不直接修改 readonly worksheet view。
- [ ] CSV 导入在真实浏览器和你的 CSP 配置下验证过 Worker 可加载。
- [ ] Snapshot 与 CSV 的用途已区分：前者恢复 Workbook，后者交换值数据。
- [ ] 插件 command 不依赖不可回滚的外部副作用来保证数据一致性。
- [ ] 大型操作的 History budget、CSV 文件尺寸和前端内存预算已在目标设备上压测。

## 相关文档

- [架构与里程碑](./architecture.md)
- [M4 数据操作契约](./m4-data-operations.md)
- [发布清单](./release.md)
- [性能报告](./performance-report.md)
- [变更记录](../CHANGELOG.md)
