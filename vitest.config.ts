import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const r = (path: string) => fileURLToPath(new URL(path, import.meta.url));

export default defineConfig({
  resolve: {
    alias: {
      "@injoysai/opensheet-shared": r("./packages/shared/src/index.ts"),
      "@injoysai/opensheet-core": r("./packages/core/src/index.ts"),
      "@injoysai/opensheet-commands": r("./packages/commands/src/index.ts"),
      "@injoysai/opensheet-history": r("./packages/history/src/index.ts"),
      "@injoysai/opensheet-formula-engine": r("./packages/formula-engine/src/index.ts"),
      "@injoysai/opensheet-clipboard": r("./packages/clipboard/src/index.ts"),
      "@injoysai/opensheet-import-export": r("./packages/import-export/src/index.ts"),
      "@injoysai/opensheet-plugin-api": r("./packages/plugin-api/src/index.ts"),
      "@injoysai/opensheet-renderer-canvas": r("./packages/renderer-canvas/src/index.ts"),
      "@injoysai/opensheet-react": r("./packages/react/src/index.ts"),
      "@injoysai/opensheet": r("./packages/runtime/src/index.ts"),
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
