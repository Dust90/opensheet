# Third-Party Licenses

OpenSheet 本体采用 [MIT](./LICENSE)。

本清单记录**直接依赖**的用途与许可证；完整依赖树（含传递依赖）由 CI 中的 `pnpm licenses:check` 按 allowlist 持续校验（允许：MIT / ISC / Apache-2.0 / BSD-2-Clause / BSD-3-Clause / 0BSD / CC0-1.0 / Unlicense / Python-2.0 / BlueOak-1.0.0，以及两个经评审的例外，见下文）。

## 运行时依赖

| 依赖 | 用途 | 许可证 |
|---|---|---|
| react / react-dom 18 | UI 组件层（apps/demo、packages/react） | MIT |
| zustand 5 | React 状态订阅 | MIT |

## 构建与工具链

| 依赖 | 用途 | 许可证 |
|---|---|---|
| typescript ~6.0.3 | 类型系统（全仓库统一版本） | Apache-2.0 |
| vite ^8 | apps/demo 开发与构建 | MIT |
| @vitejs/plugin-react | Vite React 插件 | MIT |
| tsup ^8 | 各 packages 库构建（ESM + d.ts） | MIT |
| vitest ^4 | 单元/集成测试 | MIT |
| @vitest/coverage-v8 | 测试覆盖率 | MIT |
| fast-check ^4 | 属性测试（undo/redo、snapshot 往返等恒等性质） | MIT |
| @playwright/test | E2E 测试 | Apache-2.0 |
| prettier ^3 | 代码格式化 | MIT |
| eslint 系（demo） | 代码检查 | MIT |
| yaml ^2 | 脚本辅助（许可证/清单解析） | ISC |
| @types/node, @types/react, @types/react-dom | 类型声明 | MIT |

## 经评审的例外（传递依赖，构建期，不修改、不分发其源码）

| 依赖 | 来源 | 许可证 | 评审结论 |
|---|---|---|---|
| caniuse-lite | browserslist（babel/vite 传递依赖） | CC-BY-4.0 | 数据文件而非代码；允许 |
| lightningcss | Vite CSS 压缩（可选传递依赖） | MPL-2.0 | 弱文件级 copyleft；作为工具原样使用允许 |

## 政策

- 禁止：GPL / AGPL / SSPL / BUSL / Commons-Clause / UNKNOWN，以及任何需要商业许可的生产组件（如 `@univerjs-pro/*`）。
- 新增依赖前必须确认许可证在 allowlist 内，否则 CI `license-check` job 失败。
- 本清单在新增直接依赖时同步更新。
