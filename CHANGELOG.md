# Changelog

本文按增量方式记录面向发布的用户可见变更。当前所有 workspace package 仍标记为 private；0.1.0 是待发布的首个版本，并不表示已经发布到 npm registry。

## [0.1.0] - Pending release

### Added

- Workbook Core、事务 Command Bus、History、Snapshot 和公式引擎。
- Canvas Grid、编辑、剪贴板、样式、冻结窗格、筛选、排序、查找和去重。
- CSV RFC 4180 codec、增量解析、浏览器 Worker 导入，以及 used-range CSV 导出。
- 公开 Runtime SDK、插件生命周期、观察 hooks、原子插件命令和同步插件公式函数。

### Verification

- Workspace typecheck、依赖边界、许可证检查、生产构建和单 worker E2E 已通过。
- 发布 tarball consumer smoke check、全量性能基准及 CSV 100k x 20 import/export gate 已完成。
