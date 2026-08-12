import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const r = (path: string) => fileURLToPath(new URL(path, import.meta.url));

// Resolve workspace packages to their TypeScript sources in dev, so the demo
// runs without requiring `pnpm build` of every library package first.
export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      "@injoysai/opensheet-shared": r("../../packages/shared/src/index.ts"),
      "@injoysai/opensheet-core": r("../../packages/core/src/index.ts"),
      "@injoysai/opensheet-commands": r("../../packages/commands/src/index.ts"),
      "@injoysai/opensheet-history": r("../../packages/history/src/index.ts"),
      "@injoysai/opensheet-import-export": r("../../packages/import-export/src/index.ts"),
      "@injoysai/opensheet-plugin-api": r("../../packages/plugin-api/src/index.ts"),
      "@injoysai/opensheet-formula-engine": r("../../packages/formula-engine/src/index.ts"),
      "@injoysai/opensheet-renderer-canvas": r("../../packages/renderer-canvas/src/index.ts"),
      "@injoysai/opensheet": r("../../packages/runtime/src/index.ts"),
    },
  },
});
