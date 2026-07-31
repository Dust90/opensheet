import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const r = (path: string) => fileURLToPath(new URL(path, import.meta.url));

export default defineConfig({
  resolve: {
    alias: {
      "@opensheet/shared": r("./packages/shared/src/index.ts"),
      "@opensheet/core": r("./packages/core/src/index.ts"),
      "@opensheet/commands": r("./packages/commands/src/index.ts"),
      "@opensheet/history": r("./packages/history/src/index.ts"),
      "@opensheet/formula-engine": r("./packages/formula-engine/src/index.ts"),
      "@opensheet/clipboard": r("./packages/clipboard/src/index.ts"),
      "@opensheet/import-export": r("./packages/import-export/src/index.ts"),
      "@opensheet/plugin-api": r("./packages/plugin-api/src/index.ts"),
      "@opensheet/renderer-canvas": r("./packages/renderer-canvas/src/index.ts"),
      "@opensheet/react": r("./packages/react/src/index.ts"),
      "@opensheet/runtime": r("./packages/runtime/src/index.ts"),
    },
  },
  test: {
    include: ["packages/**/*.test.ts", "tests/integration/**/*.test.ts"],
    coverage: {
      provider: "v8",
      reporter: ["text", "html"],
    },
  },
});
