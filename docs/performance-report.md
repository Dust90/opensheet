# M6 性能报告

基线日期：2026-08-12。以下数据由本地 Chromium / Playwright 单 worker 运行；它们是可复现的 release baseline，不代表跨机器或远端 CI 的 SLA。

## 复现环境与方法

- 数据集：除 Formula 外均为 100,000 rows × 20 columns（2,000,000 cells）；Formula 为默认预算内的 100,000 × 9 lazy aggregate（900,000 reads）。
- 浏览器：Playwright Chromium headless，deviceScaleFactor=1。
- Grid/Filter/Sort 使用 120 帧稳定态绘制样本，排除 warmup。
- 每条基准会断言结果的关键内容，不能仅因 Promise 返回或初始状态而通过。

| 路径                          | 命令               |          本次结果 |         Gate | 结果 |
| ----------------------------- | ------------------ | ----------------: | -----------: | ---- |
| Grid 2M scroll paint          | pnpm bench:grid    |         p95 6.4ms |        ≤50ms | ✅   |
| Formula lazy SUM 900k         | pnpm bench:formula |            90.2ms |      ≤1000ms | ✅   |
| Filter apply                  | pnpm bench:filter  |            36.2ms |       ≤500ms | ✅   |
| Filter clear                  | pnpm bench:filter  |            22.1ms |       ≤100ms | ✅   |
| Filter stable paint           | pnpm bench:filter  |         p95 6.3ms |        ≤25ms | ✅   |
| Sort 100k×20 two-key          | pnpm bench:sort    |           718.0ms |      ≤2000ms | ✅   |
| Sort undo / redo              | pnpm bench:sort    | 514.6ms / 519.5ms | ≤1500ms each | ✅   |
| Sort stable paint             | pnpm bench:sort    |         p95 5.4ms |        ≤25ms | ✅   |
| Find dense / next             | pnpm bench:find    | 182.4ms / 173.8ms |  ≤500ms each | ✅   |
| Find sparse                   | pnpm bench:find    |             0.1ms |       ≤500ms | ✅   |
| Dedupe 100k×20 (50k removed)  | pnpm bench:dedupe  |           765.7ms |      ≤1500ms | ✅   |
| Dedupe undo / redo            | pnpm bench:dedupe  | 407.9ms / 446.8ms | ≤1500ms each | ✅   |
| CSV Worker import 100k×20     | pnpm bench:csv     |          1165.1ms |      ≤5000ms | ✅   |
| CSV used-range export 100k×20 | pnpm bench:csv     |           499.2ms |      ≤2500ms | ✅   |

原始 JSON 会写入 test-results/*-perf.json，便于重新比较本机结果。

## 解释与限制

- CSV import 使用 Browser Worker + staged Worksheet；CSV export 目前在主线程构造 used-range rows、CSV string 与 Blob。若未来超大导出成为 UI 响应瓶颈，应单独设计 streaming 或 Worker export，而不是改变现有 CSV v1 语义。
- Dedupe 的 Undo/Redo 性能基准通过显式较大的 History budget 测量算法路径；普通运行仍受默认 maxMemoryBytes 的 retention policy 约束。
- 指标不应作为绝对跨设备承诺。发布前应在目标 CI/发布设备重新运行同一命令，并保留生成的原始 JSON。

## 发布验证

M6.1 提供 pnpm check:pack。

它会从 pnpm pack 的 @injoysai/opensheet 依赖闭包 tarball 创建临时消费者，以本地 tarball overrides 离线安装，执行 Vite production build，断言 csv.worker-*.js 已输出，并在 Chromium 中真实执行 CSV import/export。该检查避免 monorepo source alias 掩盖发布后的 Worker URL 或 package export 问题。
