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
  commands: CommandRegistry; // 注册元数据或可执行原子命令
  functions: FunctionRegistry; // 注册元数据或纯同步公式函数
  menus: MenuRegistry; // 注册菜单/工具栏/右键项
  hooks: {
    onBeforeCommand(cb): Unsubscribe;
    onAfterCommand(cb): Unsubscribe;
    onWorkbookLoaded(cb): Unsubscribe;
  };
}
```

`PluginHost`（plugin-api 内置）保存 contribution 清单并派发观察 hook；runtime 将可执行 contribution 装入真实 CommandBus 与每个 Workbook 的 FormulaEngine。

### 可执行命令

命令 handler 只获得只读 physical worksheet context，并返回既有 `PluginOperation[]`。Runtime 将返回的 operations 作为一个 `atomic: true` CommandBus batch 执行，因此 validation、formula refresh、rollback、History 与 Undo/Redo 继续使用核心路径。原子性只涵盖返回的 Workbook operations，不涵盖插件自行产生的 HTTP、storage 等外部副作用。

内置 CommandBus ID 不能被覆盖；插件卸载后，其命令 contribution 立即失效。before/after hooks 只观察一个插件命令，不展开其内部 operations。

### 公式函数

插件公式函数必须纯同步。函数名不区分大小写，不能覆盖内置 registry 函数或 lazy special forms `IF`、`AND`、`OR`。参数为 scalar `CellValue` 或 lazy range iterable；函数直接返回 `CellValue`。普通异常与非法返回值会变为稳定 `#VALUE!`，非有限 number 变为 `#NUM!`，已知 `CellError` 原样传播。

已安装函数会同步到现有和后创建的 Workbook；卸载后相关公式重新计算为 `#NAME?`。函数仍受 FormulaEngine 原有的 per-formula/per-transaction evaluation budget 约束。

### 观察 hooks

hooks 是观察通道，不是 middleware：不能取消核心操作，也不能将已提交操作伪装为失败。一个 observer 抛错会被隔离，不会阻止其他 observer 或 Runtime/Core 操作。

## 边界保障

`pnpm check:boundaries`（CI）同时校验两层：

1. 每个包的 `package.json` workspace 依赖 ⊆ 允许矩阵；
2. `src/` 中实际 `import "@opensheet/x"` 语句 ⊆ 允许矩阵（测试文件豁免）。

规则示例：`core`/`plugin-api` 仅允许 `shared`；`renderer-canvas` 禁止 `react`；仅 `runtime` 可依赖全部功能包。
