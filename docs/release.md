# M6 发布清单

M6 的本地工程验证已完成。发布闭包已固定为 npmjs.org 上的 public `@injoysai` package；尚未执行真正 publish。

## 发布包

以下 8 个包均为版本 0.1.0、MIT、`https://registry.npmjs.org/` 和 `access: public`：

- `@injoysai/opensheet`
- `@injoysai/opensheet-shared`
- `@injoysai/opensheet-core`
- `@injoysai/opensheet-commands`
- `@injoysai/opensheet-history`
- `@injoysai/opensheet-formula-engine`
- `@injoysai/opensheet-import-export`
- `@injoysai/opensheet-plugin-api`

根 workspace、Demo、React、Canvas Renderer 和 Clipboard 保持 `private: true`，不会参与发布。

## 已完成的工程验证

- pnpm typecheck、pnpm check:boundaries、pnpm licenses:check 和 pnpm build
- 单 worker Playwright E2E
- docs/performance-report.md 中记录的 Grid、Formula、Filter、Sort、Find、Dedupe 和 CSV 基准
- pnpm check:pack：从 pnpm pack tarball 创建临时 consumer，离线安装、Vite 构建并在 Chromium 中执行 Formula、Filter、Sort、Dedupe、插件和真实 CSV Worker import/export

## Release owner 决策

1. 确认 npm token、2FA、provenance 和组织访问权限在正式 publish 时仍有效。
2. 决定是否先发布 `0.1.0-next.0`。若采用 prerelease，必须先将所有发布包改为该版本；同一版本不能先以 `next` 发布后再重新发布为 `latest`。
3. 确定 changelog 日期、GitHub Release 文案与发布审批人。

## 发布前最后命令

在以上决策完成后，依次运行：

```sh
pnpm typecheck
pnpm check:boundaries
pnpm licenses:check
pnpm build
pnpm test:e2e --workers=1
pnpm check:pack
pnpm release:dry-run
```

`pnpm release:dry-run` 按固定 allowlist 和依赖顺序逐一执行 8 个公开包的 `npm publish --dry-run --access public --tag next --registry https://registry.npmjs.org/`；它不会发布 root、Demo、React、Renderer 或 Clipboard。

发布后应从 registry 创建新的 consumer 项目，重复 CSV Worker import/export smoke test，确认 tarball 与 registry metadata 均可用。
