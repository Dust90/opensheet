# M6 发布清单

M6 的本地工程验证已完成。真正发布前仍需由 release owner 决定 registry、package 可见性和版本策略；当前 workspace package 均为 private 且版本为 0.1.0。

## 已完成的工程验证

- pnpm typecheck、pnpm check:boundaries、pnpm licenses:check 和 pnpm build
- 单 worker Playwright E2E
- docs/performance-report.md 中记录的 Grid、Formula、Filter、Sort、Find、Dedupe 和 CSV 基准
- pnpm check:pack：从 pnpm pack tarball 创建临时 consumer，离线安装、Vite 构建并在 Chromium 中执行真实 CSV Worker import/export

## Release owner 决策

1. 确定发布 registry 与 package scope。
2. 决定哪些 package 对外发布，并将对应 package 的 private 配置改为目标可见性。
3. 确定首发版本、dist-tag 和 Git tag 策略。
4. 确认 npm token、2FA、provenance 和组织访问权限。
5. 确定 changelog 日期、GitHub Release 文案与发布审批人。

## 发布前最后命令

在以上决策完成且 package metadata 已更新后，依次运行：

```sh
pnpm typecheck
pnpm check:boundaries
pnpm licenses:check
pnpm build
pnpm test:e2e --workers=1
pnpm check:pack
npm publish --dry-run
```

发布后应从 registry 创建新的 consumer 项目，重复 CSV Worker import/export smoke test，确认 tarball 与 registry metadata 均可用。
