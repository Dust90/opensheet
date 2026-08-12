import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 30_000,
  use: {
    baseURL: "http://127.0.0.1:5173",
    trace: "on-first-retry",
    // Playwright 1.55 headless defaults to the headless shell build, which is
    // not installed on macOS 13 runners; the full chromium binary (new headless
    // mode) is used instead — identical behavior for these tests.
    channel: "chromium",
  },
  webServer: {
    command: "pnpm --filter @injoysai/opensheet-demo dev --host 127.0.0.1",
    url: "http://127.0.0.1:5173",
    reuseExistingServer: !process.env.CI,
    timeout: 120_000
  }
});
