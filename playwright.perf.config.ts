// Config for the reproducible grid perf benchmark (pnpm bench:grid).
// Kept separate so `pnpm test:e2e` never picks up perf specs.

import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests/performance",
  timeout: 300_000,
  workers: 1,
  retries: 0,
  use: {
    baseURL: "http://127.0.0.1:5173",
    // Full chromium binary (new headless) — the headless shell build is not
    // installed on macOS 13; keeps perf numbers comparable across runners.
    channel: "chromium",
  },
  webServer: {
    command: "pnpm --filter @opensheet/demo dev --host 127.0.0.1",
    url: "http://127.0.0.1:5173",
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
