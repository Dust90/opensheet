# OpenSheet

独立的高性能 Web 电子表格内核。基于 React 18、TypeScript 与 Canvas 2D，提供类传统电子表格的核心编辑体验，并以公开 API 供任何 Web / 桌面 / AI 宿主嵌入。

OpenSheet 独立维护、独立构建、独立发布，不依赖任何专有商业电子表格组件。

## 环境要求

- Node.js >= 20.19
- pnpm ^11（`corepack enable` 或 `npm i -g pnpm`）
- 现代浏览器（Chrome / Edge / Firefox / Safari 最新两个大版本）

## 安装

```bash
pnpm install
```

## 开发

```bash
pnpm dev            # 启动 apps/demo 开发服务器（默认 http://127.0.0.1:5173）
```

## 测试

```bash
pnpm test           # Vitest 单元测试 + 集成测试
pnpm test:coverage  # 覆盖率报告
pnpm test:e2e       # Playwright E2E（自动拉起 demo dev server）
```

## 类型检查与构建

```bash
pnpm typecheck      # 全部包 tsc --noEmit
pnpm build          # 全部包 tsup 构建 + demo 生产构建
```

## SDK 使用示例

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

## 项目结构

```text
apps/
  demo/            # 独立可运行的表格应用
  playground/      # 性能测试页（10w 行 × 100 列场景）
packages/
  shared/          # 契约：A1 地址、Range、错误、事件（零依赖）
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
  performance/     # 性能基准
docs/              # 架构、数据模型、插件系统、性能报告、ADR
```

## 当前功能

> 项目处于 M0（架构与骨架）阶段，功能清单随里程碑更新。目标功能见 `docs/architecture.md` 里程碑计划。

- [x] Monorepo 骨架、构建与测试工具链
- [ ] Canvas 渲染内核
- [ ] 单元格编辑、剪贴板、Undo/Redo、Snapshot
- [ ] 公式引擎（28 个基础函数）
- [ ] 排序、筛选、查找、去重
- [ ] CSV 导入导出、公开 SDK API

## 已知限制

第一阶段不实现：XLSX 兼容、多人协作、图表、数据透视表、宏、打印、评论、云端同步、账号系统、AI 集成、跨工作表公式引用。

## 许可证

[ISC](./LICENSE)。第三方依赖许可证见 [THIRD_PARTY_LICENSES.md](./THIRD_PARTY_LICENSES.md)。
