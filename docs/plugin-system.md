# 插件系统

## 设计原则

- `plugin-api` **只依赖 shared**：插件看不见 runtime、React 或完整 `OpenSheetAPI`，只看见窄接口注册表。
- 插件通过注册表**贡献能力**，由 runtime 在装配时排入真实系统（command bus、公式引擎、UI）。
- 一期不做插件市场、远程下载、沙箱、签名、动态权限。

## 接口

```ts
interface OpenSheetPlugin {
  id: string;
  setup(context: OpenSheetPluginContext): void | Promise<void>;
  dispose?(): void | Promise<void>;
}

interface OpenSheetPluginContext {
  commands: CommandRegistry;    // 注册命令贡献 { id, description }
  functions: FunctionRegistry;  // 注册公式函数 { name, minArgs, maxArgs }
  menus: MenuRegistry;          // 注册菜单/工具栏/右键项
  hooks: {
    onBeforeCommand(cb): Unsubscribe;
    onAfterCommand(cb): Unsubscribe;
    onWorkbookLoaded(cb): Unsubscribe;
  };
}
```

`PluginHost`（plugin-api 内置）保存贡献清单并派发 hook；runtime 通过 `listCommandContributions()` 等 drain 方法把贡献装入真实注册表。

## 一期内置插件（与核心同仓库构建，不走动态加载）

| 插件 | 落地里程碑 | 贡献 |
|---|---|---|
| FormulaPlugin | M3 | 28 个公式函数 + derived 重算钩子 |
| ClipboardPlugin | M2 | 复制/剪切/粘贴命令 + TSV 互转 |
| FormattingPlugin | M2 | 样式命令 + 工具栏项 |
| SortFilterPlugin | M4 | 排序/筛选/去重命令 + 菜单项 |
| CSVPlugin | M5 | CSV 导入导出命令 |

内置插件与核心同仓库构建的唯一目的，是验证"内核不硬编码功能、全部能力可经注册表注入"这一架构假设。

## 边界保障

`pnpm check:boundaries`（CI）同时校验两层：

1. 每个包的 `package.json` workspace 依赖 ⊆ 允许矩阵；
2. `src/` 中实际 `import "@opensheet/x"` 语句 ⊆ 允许矩阵（测试文件豁免）。

规则示例：`core`/`plugin-api` 仅允许 `shared`；`renderer-canvas` 禁止 `react`；仅 `runtime` 可依赖全部功能包。
