# OpenSheet 架构

> 当前状态：M0、M1 已完成（内核骨架 + Canvas 渲染）。M2–M6 见文末里程碑表。

## 分层

```text
React UI（toolbar / formula bar / sheet tabs / dialogs）        packages/react
  ↓ 订阅 ChangeEvent，派发意图
Interaction Controller（键鼠 / 剪贴板 / DOM 编辑器）            packages/renderer-canvas (M1)
  ↓ 全部翻译为 Command
Command Bus（校验 → 事务 → journal → 历史 → 事件缓冲）          packages/commands
  ↓
Workbook Core（稀疏数据模型，纯 TS，无 DOM/React）              packages/core
  ↓
Formula（M3，derived 通道）/ History / Snapshot                 packages/{formula-engine,history}
  ↓ ChangeEvent（合并后单事件）
Canvas Renderer（Content + Overlay 双 Canvas，Dirty Region）    packages/renderer-canvas (M1)
```

核心规则：

1. `core` 只依赖 `shared`，不 import commands/history/formula/plugin（由 `pnpm check:boundaries` 在 CI 强制）。
2. Renderer 只读 core（只读视图接口，M1）。
3. UI 不改数据，一切修改经 Command Bus。
4. 外部 SDK 与内部 UI 操作同路径（`applyOperations` 即事务门面）。
5. `plugin-api` 只依赖 `shared`；插件只见窄接口注册表，不见 runtime/React/OpenSheetAPI。
6. `runtime` 是唯一组合根：装配 core + commands + history + 插件，导出 `createOpenSheet()`。

## 无环包依赖

```text
shared ← core ← commands ← history
              ← formula-engine / clipboard / import-export
shared ← plugin-api
core+plugin-api+shared ← renderer-canvas
renderer-canvas+plugin-api+runtime+shared ← react
全部功能包 ← runtime ← apps/*
```

## 事务语义（ADR-0003）

```text
beginBatch
→ 逐条执行命令，journal 记录逆补丁，ChangeEvent 入缓冲，不重算公式，不通知外界
成功 → beforeCommit（公式 derived 重算，M3）→ endBatch(true)：每 sheet+source 一条合并事件 → 一条 History
失败 → 逆序回放 journal → endBatch(false)：缓冲丢弃、无历史、外界无感知
```

## 里程碑

| 里程碑 | 内容                                                                                              | E2E 场景             |
| ------ | ------------------------------------------------------------------------------------------------- | -------------------- |
| M0 ✅  | 骨架、shared、core、commands+事务、history、plugin-api、runtime 骨架、benchmark、ADR              | —                    |
| M1 ✅  | 双 Canvas 渲染内核（Content+Overlay）、前缀和二分定位、虚拟化、冻结四分区、选区与键盘导航、滚动条 | —                    |
| M2 ✅  | DOM 编辑器、剪贴板、行列操作、样式、Undo/Redo UI、Snapshot 持久化                                 | 1,2,3,5,6,7,12,13,14 |
| M3 ✅  | 公式引擎（28 函数、依赖图、环检测、derived 通道）                                                 | 4                    |
| M4 ✅  | 排序、筛选、查找、去重（追加行移出本轮；契约见 `docs/m4-data-operations.md`）                     | 8,9                  |
| M5 ✅  | CSV Worker、完整 SDK API、插件生命周期与可执行命令/公式函数                                       | 10,11,15,16,17       |
| M6     | 性能报告、发布、npm 构建                                                                          | 全量回归             |

性能目标与设计约束见 `docs/performance-report.md`（M6 填实测数据）。
