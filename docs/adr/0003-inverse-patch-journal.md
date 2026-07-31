# ADR-0003: 事务 = inverse patch journal + 事件缓冲 + 提交后重算

## 状态

已接受（M0 实现）

## 背景

`applyOperations(atomic=true)` 需要：任一子操作失败时完全回滚；外部观察者（renderer、React、插件 hook）绝不能看到事务中间状态；公式重算结果也不能泄漏部分状态。

## 决策

1. **Inverse patch journal**：每个命令执行时返回 `JournalEntry { undo, redo, affected, approxBytes }`，闭包捕获逆补丁（旧值/旧结构）。无法生成逆操作的命令（如未来的 sort/dedupe）只允许**深拷贝受影响 Range**，禁止全表快照。废弃早期"复制 Map 引用作为 before snapshot"方案。
2. **事件缓冲**：`Workbook.beginBatch()` 开启缓冲；事务内所有 `emit` 暂存。成功时 `endBatch(true)` 合并为**每个 sheet+source 一条** `ChangeEvent` 派发；失败时 `endBatch(false)` 整体丢弃。
3. **公式提交时机**：`beforeCommit` 钩子在事件合并前、事务仍开启时运行。M3 公式引擎的 derived 重算在此执行，derived 事件并入同一批缓冲。derived 变更**永不进入 Undo 历史**（bus 不推送 `source === "derived"` 的 journal）。
4. **失败路径**：逆序回放 journal → 丢弃缓冲 → 不写 History → 返回 `ApplyOperationsError { failedOperationIndex, errorCode }`。

## 后果

- 集成测试锁定语义：atomic 成功 = 单事件 + 单历史；atomic 失败 = 零事件 + 零历史 + 状态全等（含 fast-check 属性测试）。
- 每个命令的逆补丁正确性是核心质量风险，由 execute→undo→redo 全等属性测试持续看护。
