# ADR-0005: CellStore 冻结为分块存储（128×128 chunk）

## 状态

已接受（M0，benchmark 驱动决策）

## 背景

`CellStore` 接口隔离了稀疏存储实现，三个候选：字符串 key `"r:c"`、数字 key `row*2^20+col`、分块 `Map<chunkKey, Map<innerKey>>`。冻结前必须在 2,000,000 有值单元格（100,000 行 × 20 列）场景下实测。

## 测量方法

`pnpm bench:cell-store`：每个候选在**独立 node 进程**运行（`--expose-gc --max-old-space-size=6144`），避免 GC 状态互相污染。测量写入 2M 单元格、100 万次确定性随机读、全 Range 遍历、Snapshot 式序列化耗时，以及 GC 后 RSS / heapUsed。该 benchmark 不属于常规 CI。

原始数据：`docs/benchmarks/cell-store-2026-07-31.json`（已入库；`test-results/` 副本被 gitignore）。

## 结果（2026-07-31，node v22.22.2，darwin x64，M 系列 Mac）

| 实现 | 写入 2M | 1M 随机读 | 全 Range 遍历 | 序列化 | RSS | heapUsed |
|---|---|---|---|---|---|---|
| string-key | 2621.9 ms | 1060.6 ms | 327.1 ms | 3887.6 ms | 243.8 MB | 167.2 MB |
| number-key | 2548.1 ms | 589.4 ms | 652.6 ms | 3681.9 ms | 219.8 MB | 151.3 MB |
| **chunked** | **904.8 ms** | **422.8 ms** | **63.6 ms** | 3717.3 ms | 223.7 MB | **150.9 MB** |

## 决策

冻结默认实现为 **chunked（128×128）**：

- 写入快 2.8×（小 Map 再哈希成本低、缓存局部性好）；
- 随机读最快；
- Range 遍历快 5–10×（渲染与排序的主路径）；
- 内存与 number-key 持平，显著优于 string-key；
- 序列化三者相当（瓶颈在对象分配而非容器）。

## 后果

- `Worksheet` 默认注入 `chunkedCellStoreFactory`，仍可通过 `WorksheetInit.storeFactory` 替换（用于对比测试）。
- 数字 key 方案保留了 2^20 列上限的隐性约束，chunked 方案同样受 chunk key 空间约束（chunk 列索引 < 2^20，即列 < 2^27），远超市面表格上限，可接受。

## 冻结范围说明

本次 benchmark 仅覆盖"窄而长、连续密集、全范围遍历"场景；未覆盖稀疏随机分布、超宽表（1000+ 列）、小视口遍历、频繁删除、插入行列重建、超大逻辑表少量非空等场景。因此结论措辞为：**chunked 冻结为 M0/M1 默认实现**；M6 性能阶段可在保持 `CellStore` 接口不变的前提下，针对上述场景补测并重新评估 chunk size 甚至实现本身。
