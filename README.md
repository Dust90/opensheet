# OpenSheet

独立的高性能 Web 电子表格内核。基于 React 18、TypeScript 与 Canvas 2D，提供类传统电子表格的核心编辑体验，并以公开 API 供任何 Web / 桌面 / AI 宿主嵌入。

OpenSheet 独立维护、独立构建、独立发布，不依赖任何专有商业电子表格组件。

> **当前状态：M1（Canvas 渲染内核）已完成，进入 M2。** 以下标注 *Planned* 的能力尚未实现，以 `docs/architecture.md` 里程碑计划为准。

## 环境要求

- Node.js >= 20.19
- pnpm 11（推荐 `corepack enable`，根 `package.json` 的 `packageManager` 字段会锁定版本）
- 现代浏览器（Chrome / Edge / Firefox / Safari 最新两个大版本）

## 安装

```bash
corepack enable   # 可选，用于锁定 pnpm 版本
pnpm install
```

## 开发

```bash
pnpm dev            # 启动 apps/demo 开发服务器（默认 http://127.0.0.1:5173）
```

## 测试

```bash
pnpm test              # Vitest 单元测试 + 集成测试
pnpm test:coverage     # 覆盖率报告
pnpm test:e2e          # Playwright E2E（自动拉起 demo dev server）
pnpm bench:cell-store  # CellStore 存储方案 benchmark（手动执行，不属于常规 CI）
```

## 类型检查、边界与构建

```bash
pnpm typecheck         # 全部包 tsc --noEmit
pnpm check:boundaries  # 包间依赖边界检查（core/plugin-api 只依赖 shared 等）
pnpm licenses:check    # 依赖许可证 allowlist 检查
pnpm build             # 全部包 tsup 构建 + demo 生产构建
```

## Planned SDK API

> The following API is planned. `createOpenSheet()` 与 M0 子集（createWorkbook / applyOperations / Snapshot / undo / redo / readRange / searchCells）已在 `@opensheet/runtime` 提供骨架实现；CSV 与完整能力将在 M5 落地。

```ts
import { createOpenSheet } from "@opensheet/runtime";

const sheet = createOpenSheet();

const workbook = sheet.createWorkbook({ name: "Q3 Budget" });

await sheet.applyOperations({
  workbookId: workbook.id,
  sheetId: workbook.activeSheetId,
  atomic: true,
  operations: [
    { type: "cell.set", range: "A1", value: "Item" },
    { type: "cell.set", range: "B1", value: "Amount" },
    { type: "formula.set", range: "B10", formula: "=SUM(B2:B9)" },
  ],
});

const snapshot = sheet.getWorkbookSnapshot();
const csv = await sheet.exportCSV({ sheetId: workbook.activeSheetId });
```

## Target Monorepo Structure

> 目标结构；`apps/playground`、`tests/performance` 中的部分目录与包将随里程碑逐步补齐。

```text
apps/
  demo/            # 独立可运行的表格应用
  playground/      # 性能测试页（10w 行 × 100 列场景）
packages/
  shared/          # 契约：A1 地址、Range、错误、事件、CellStore 接口（零依赖）
  core/            # 数据模型、Snapshot、变更事件（仅依赖 shared）
  commands/        # Command Bus、inverse patch journal、事务
  history/         # Undo/Redo，条数 + 内存双上限
  formula-engine/  # Tokenizer / Parser / AST / 依赖图 / 函数注册
  clipboard/       # TSV 与外部表格粘贴
  import-export/   # CSV 导入导出 + Worker 解析
  renderer-canvas/ # 双 Canvas 渲染、虚拟化、DOM 编辑器
  plugin-api/      # 插件契约与能力注册表（仅依赖 shared）
  react/           # 组件、Hooks、主题
  runtime/         # 组合根：createOpenSheet() / OpenSheetAPI
tests/
  integration/     # Command → Core → Event → Renderer 链路、事务回滚
  e2e/             # Playwright 场景
  performance/     # 性能基准（不进常规 CI）
docs/              # 架构、数据模型、插件系统、性能报告、ADR
```

## 当前功能

- [x] Monorepo 骨架、构建与测试工具链、依赖边界检查、许可证检查
- [x] M0：shared 契约、core 数据模型与 Snapshot、Command Bus + 事务、History、plugin-api、runtime 骨架
- [x] M1：双 Canvas 渲染内核（前缀和二分定位、虚拟化、冻结四分区、选区与键盘导航、滚动条；实测 2M 单元格滚动绘制 p95=7.1ms）
- [ ] M2：单元格编辑、剪贴板、行列操作、样式、Undo/Redo UI
- [ ] M3：公式引擎（28 个基础函数，derived-update 通道）
- [ ] M4：排序、筛选、查找、去重
- [ ] M5：CSV 导入导出、完整公开 SDK API

## 已知限制

第一阶段不实现：XLSX 兼容、多人协作、图表、数据透视表、宏、打印、评论、云端同步、账号系统、AI 集成、跨工作表公式引用。

## 许可证

[MIT](./LICENSE)。第三方依赖许可证清单见 [THIRD_PARTY_LICENSES.md](./THIRD_PARTY_LICENSES.md)，由 `pnpm licenses:check` 在 CI 中持续校验。
