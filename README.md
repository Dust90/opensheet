# OpenSheet

InjoysAI 开发的独立高性能 Web 电子表格内核。基于 React 18、TypeScript 与 Canvas 2D，提供类传统电子表格的核心编辑体验，并以公开 API 供任何 Web / 桌面 / AI 宿主嵌入。

OpenSheet 独立维护、独立构建、独立发布，不依赖任何专有商业电子表格组件。

> **当前状态：M5（CSV Worker、公开 SDK、插件执行）已完成，进入 M6 发布与性能收口。** 里程碑计划见 `docs/architecture.md`。

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
pnpm bench:filter      # M4.2 100k × 20 Filter benchmark（手动执行）
pnpm bench:formula     # M3 默认预算下 900k-cell lazy SUM benchmark（手动执行）
pnpm bench:sort        # M4.3 100k × 20 Sort benchmark（手动执行）
pnpm bench:find        # M4.4 100k × 20 Find benchmark（手动执行）
pnpm bench:dedupe      # M4.5 100k × 20 Dedupe benchmark（手动执行）
pnpm bench:csv         # M5 浏览器 Worker 100k × 20 CSV import/export benchmark（手动执行）
```

## 类型检查、边界与构建

```bash
pnpm typecheck         # 全部包 tsc --noEmit
pnpm check:boundaries  # 包间依赖边界检查（core/plugin-api 只依赖 shared 等）
pnpm check:pack        # 从 tarball 安装的 Vite consumer：Formula / Filter / Sort / Dedupe / Plugin / CSV Worker smoke
pnpm release:dry-run   # 8 个公开包按依赖顺序 npm publish --dry-run（不实际发布）
pnpm release:publish-next -- 0.1.0-next.0  # 显式确认版本后发布 8 个包到 npm tag next
pnpm licenses:check    # 依赖许可证 allowlist 检查
pnpm build             # 全部包 tsup 构建 + demo 生产构建
```

## Public SDK API

`@injoysai/opensheet` 提供完整的公开组合根。CSV 导入始终新建 Worksheet 并以 A1 为起点，失败不会留下半成品 Sheet；CSV 导出使用 computed values 的 used range。插件命令返回内置 operation 组成一个原子 History batch；插件公式函数为纯同步函数并复用公式引擎的 lazy range 与预算。

完整的安装、Runtime、事务、公式、数据操作、CSV、Snapshot 与插件接入说明见 [接入指南](./docs/getting-started.md)。

```ts
import { createOpenSheet } from "@injoysai/opensheet";

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

const imported = await sheet.importCSV({
  file: new File(["Name,Amount\r\nAda,10"], "sales.csv", { type: "text/csv" }),
});

await sheet.usePlugin({
  id: "example",
  setup({ functions }) {
    functions.registerFunction({
      name: "DOUBLE",
      minArgs: 1,
      maxArgs: 1,
      execute: ([value]) =>
        typeof value === "number" ? value * 2 : { type: "#VALUE!" },
    });
  },
});
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
- [x] M2：单元格编辑、剪贴板、行列操作、样式、Undo/Redo UI
- [x] M3：公式引擎（28 个基础函数，derived-update 通道）
- [x] M4：排序、筛选、查找、去重（契约见 `docs/m4-data-operations.md`）
- [x] M5：CSV Worker 导入导出、完整公开 SDK API、插件生命周期与可执行命令/公式函数

## 已知限制

第一阶段不实现：XLSX 兼容、多人协作、图表、数据透视表、宏、打印、评论、云端同步、账号系统、AI 集成、跨工作表公式引用。

## 许可证

[MIT](./LICENSE)。第三方依赖许可证清单见 [THIRD_PARTY_LICENSES.md](./THIRD_PARTY_LICENSES.md)，由 `pnpm licenses:check` 在 CI 中持续校验。

## 发布

M6 的发布前工程检查与性能报告见 [docs/release.md](./docs/release.md) 和 [docs/performance-report.md](./docs/performance-report.md)；待发布的首发版本记录在 [CHANGELOG.md](./CHANGELOG.md)。
