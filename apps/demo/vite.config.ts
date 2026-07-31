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
      "@opensheet/shared": r("../../packages/shared/src/index.ts"),
      "@opensheet/core": r("../../packages/core/src/index.ts"),
      "@opensheet/commands": r("../../packages/commands/src/index.ts"),
      "@opensheet/history": r("../../packages/history/src/index.ts"),
      "@opensheet/plugin-api": r("../../packages/plugin-api/src/index.ts"),
      "@opensheet/renderer-canvas": r("../../packages/renderer-canvas/src/index.ts"),
      "@opensheet/runtime": r("../../packages/runtime/src/index.ts"),
    },
  },
});
