# opensheet-core Rust Crate 实现规划

## 已确认决策

| 决策项 | 结论 |
|---|---|
| 架构 | **原生 Rust**，Tauri workspace member 直接依赖，无 WASM 桥 |
| 发布方式 | Monorepo workspace 内部成员，暂不发布 crates.io |
| 插件系统 | v1 仅函数注册（`dyn Fn` trait），命令插件 v2 再做 |
| 公式跨 Sheet | 同 TS 版，M3 阶段仅同 Sheet |
| 实现顺序 | 自底而上：数据模型 → 命令 → 公式 → 数据操作 → CSV → 组合根 |
| WASM 支持 | 预留 `wasm` feature flag，v1 不实现 |

> [!NOTE]
> Tauri 本身是 Rust 运行时，`opensheet-core` 作为 workspace member 被 `src-tauri` 原生引用，性能最优，无序列化桥开销。

---

## 目标

将 OpenSheet 的**无渲染核心**封装为一个独立的 Rust crate `opensheet-core`，可被：
- Tauri 应用直接引用（原生调用，零序列化桥损耗）
- 其他 Rust 项目作为库依赖
- 未来可选编译为 WASM（`wasm-bindgen` feature flag）

渲染层（`renderer-canvas`、`react`）**不在此范围内**，仍由前端承担。

---

## 范围确认：哪些模块进 Rust

| TS 包 | 进 Rust crate？ | 说明 |
|---|---|---|
| `shared` | ✅ 是 | 所有基础类型 |
| `core` | ✅ 是 | Workbook / Worksheet / CellStore |
| `commands` | ✅ 是 | Command Bus + 事务 |
| `history` | ✅ 是 | Undo/Redo |
| `formula-engine` | ✅ 是 | Lexer/Parser/AST/DependencyGraph/Evaluate |
| `import-export` | ✅ 是 | CSV（用 `csv` crate）|
| `plugin-api` | ⚠️ 部分 | 仅函数注册（`dyn Fn` trait），命令插件留到 v2 |
| `renderer-canvas` | ❌ 否 | 浏览器 Canvas API |
| `react` | ❌ 否 | React 组件层 |
| `runtime` | ✅ 适配 | 提供等价的 `OpenSheet` struct 组合根 |

---

## Crate 架构设计

```
opensheet-core/         ← workspace root（独立 git repo 或 workspace member）
├── Cargo.toml
├── src/
│   ├── lib.rs          ← 公开 re-exports
│   ├── address.rs      ← CellAddress / CellRef / col_to_name / parse_address
│   ├── range.rs        ← Range struct + 工具函数
│   ├── cell.rs         ← CellValue / CellData / CellStyle / CellError
│   ├── errors.rs       ← SheetError (thiserror)
│   ├── events.rs       ← ChangeEvent / ChangeListener trait
│   ├── cell_store/
│   │   ├── mod.rs
│   │   └── sparse.rs   ← HashMap<(u32,u32), CellData>，实现 CellStore trait
│   ├── workbook.rs     ← Workbook struct（含事务批处理、事件总线）
│   ├── worksheet.rs    ← Worksheet struct（行列元数据、冻结、filter状态）
│   ├── styles.rs       ← StyleTable（HashMap<StyleId, CellStyle>）
│   ├── snapshot.rs     ← WorkbookSnapshot V1/V2（serde JSON round-trip）
│   ├── commands/
│   │   ├── mod.rs      ← Command trait + CommandBus
│   │   ├── cell.rs     ← SetCell / ClearCell / SetStyle
│   │   ├── sheet.rs    ← AddSheet / RemoveSheet / RenameSheet
│   │   ├── rows_cols.rs← InsertRow / DeleteRow / InsertCol / DeleteCol
│   │   └── data_ops.rs ← Sort / Filter / Dedupe / Find 命令
│   ├── history.rs      ← UndoRedo（条数上限 + 内存上限）
│   ├── formula/
│   │   ├── mod.rs
│   │   ├── lexer.rs    ← Token / Lexer
│   │   ├── parser.rs   ← Pratt parser → Expr AST
│   │   ├── ast.rs      ← Expr enum / FormulaDependencies
│   │   ├── dependency.rs ← DependencyGraph（HashMap + petgraph）
│   │   ├── evaluate.rs ← 公式求值器（lazy range budget）
│   │   ├── functions.rs ← 28 个内置函数
│   │   └── rewrite.rs  ← 行列偏移重写（InsertRow 后公式地址更新）
│   ├── data_ops/
│   │   ├── sort.rs
│   │   ├── filter.rs
│   │   ├── dedupe.rs
│   │   └── find.rs
│   ├── csv/
│   │   ├── import.rs
│   │   └── export.rs
│   ├── plugin.rs       ← FunctionPlugin trait / FunctionRegistry
│   └── opensheet.rs    ← OpenSheet struct（组合根，等价 createOpenSheet()）
```

---

## 关键技术决策

### 1. 依赖库

```toml
[dependencies]
thiserror     = "2"          # SheetError 派生
serde         = { version = "1", features = ["derive"] }
serde_json    = "1"          # Snapshot JSON
uuid          = { version = "1", features = ["v4"] }   # workbook/sheet ID
csv           = "1"          # CSV 导入导出
indexmap      = "2"          # 有序 Map（StyleTable 确定性序列化）

# 可选：图算法（DependencyGraph 的拓扑排序）
# petgraph 功能完备但偏重；也可内联轻量 DFS（当前 TS 就是内联 DFS）
# 推荐内联，保持零额外依赖

[features]
default  = []
wasm     = ["wasm-bindgen"]  # 未来预留
```

### 2. 事件总线

```rust
// 用 Arc<Mutex<Vec<Box<dyn Fn(&ChangeEvent) + Send + Sync>>>> 实现 on_change
// Tauri command handler 持有 Arc<Mutex<OpenSheet>>，事件回调可发 Tauri emit
pub trait ChangeListener: Send + Sync {
    fn on_change(&self, event: &ChangeEvent);
}
```

### 3. CellStore

```rust
// 稀疏 HashMap，row/col 上限 u32，key = (row, col)
pub struct SparseCellStore {
    cells: HashMap<(u32, u32), CellData>,
}
```

### 4. 公式引擎 borrow 挑战

求值时需要同时读 worksheet（取单元格值）又写单元格（缓存结果），这在 Rust 中是经典的 `&mut` 冲突：

**解决方案**：求值器接受 `&dyn CellReader`（不可变读），返回 `Vec<(CellAddress, CellValue)>` patch 列表，然后 apply 到 worksheet。求值过程全程只读，避免 borrow 冲突。

### 5. Tauri 集成模式

```rust
// src-tauri/src/main.rs
use std::sync::Mutex;
use opensheet_core::OpenSheet;

#[tauri::command]
fn apply_operations(
    state: tauri::State<Mutex<OpenSheet>>,
    ops: Vec<OperationInput>,
) -> Result<(), String> {
    let mut sheet = state.lock().unwrap();
    sheet.apply_operations(ops).map_err(|e| e.to_string())
}
```

---

## 分阶段实现计划

### Phase 1 — 基础类型与数据模型（约 1.5 周）
- `address.rs`、`range.rs`、`cell.rs`、`errors.rs`
- `cell_store/sparse.rs`
- `workbook.rs`、`worksheet.rs`、`styles.rs`
- `snapshot.rs`（serde JSON 序列化/反序列化，兼容 V1/V2）
- 单元测试：地址解析、稀疏存储、Snapshot round-trip

### Phase 2 — 命令总线 + History（约 1 周）
- `commands/mod.rs`（Command trait + inverse patch + 事务）
- 全部 command 实现（cell / sheet / rows_cols）
- `history.rs`（条数 + 内存双上限）
- 单元测试：事务回滚、undo/redo

### Phase 3 — 公式引擎（约 2 周，最复杂）
- `formula/lexer.rs`、`formula/parser.rs`（Pratt 解析）
- `formula/ast.rs`、`formula/dependency.rs`（DFS 拓扑排序）
- `formula/evaluate.rs`（含 lazy range budget）
- `formula/functions.rs`（28 个函数）
- `formula/rewrite.rs`（行列偏移重写）
- 单元测试：all TS formula test cases 翻译为 Rust 测试

### Phase 4 — 数据操作 + CSV（约 1 周）
- `data_ops/{sort,filter,dedupe,find}.rs`
- `commands/data_ops.rs`
- `csv/{import,export}.rs`

### Phase 5 — 插件 API + 组合根（约 0.5 周）
- `plugin.rs`（函数注册 `dyn Fn`）
- `opensheet.rs`（`OpenSheet` 组合根，公开 SDK API）

### Phase 6 — Tauri 示例集成（约 0.5 周）
- 在现有 `apps/demo` 旁新增 `apps/tauri-demo/`
- `src-tauri/` 引用 `opensheet-core`
- 前端通过 Tauri IPC 调用，渲染层仍用 React + Canvas

---

## 与现有 TS 版本的共存策略

```
monorepo/
├── packages/           ← 现有 TS 包（Web 用）
├── crates/
│   └── opensheet-core/ ← 新 Rust crate（Tauri/服务端用）
└── apps/
    ├── demo/           ← 现有 Web demo（纯 TS）
    └── tauri-demo/     ← 新 Tauri demo（Rust core + React UI）
```

两套并行维护，TS 版不废弃，Rust 版作为高性能替代后端。

---

## Tauri 集成架构图

```
 monorepo/
 ├── Cargo.toml               ← workspace root
 ├── crates/
 │   └── opensheet-core/      ← 本 crate
 ├── packages/                ← 现有 TS 包（Web 用，不废弃）
 └── apps/
     ├── demo/                ← 现有 Web demo（纯 TS）
     └── tauri-demo/
         ├── src/             ← React UI（调用 Tauri IPC）
         └── src-tauri/
             ├── Cargo.toml   ← 依赖 opensheet-core
             └── src/
                 └── main.rs  ← Tauri commands 封装
```

```rust
// apps/tauri-demo/src-tauri/src/main.rs 示例
use std::sync::Mutex;
use opensheet_core::OpenSheet;

#[tauri::command]
fn apply_operations(
    state: tauri::State<Mutex<OpenSheet>>,
    ops: Vec<serde_json::Value>,
) -> Result<serde_json::Value, String> {
    let mut sheet = state.lock().unwrap();
    sheet.apply_operations_json(ops).map_err(|e| e.to_string())
}
```

---

## 验证计划

- 将现有 TS 版所有 `__tests__` 翻译为 Rust `#[cfg(test)]` 模块
- `cargo bench`：与 TS benchmark 结果对比（目标：SUM 900k cells < 50ms）
- Snapshot round-trip 测试：Rust 读 TS 产生的 JSON，TS 读 Rust 产生的 JSON
