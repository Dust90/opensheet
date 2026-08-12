// `pnpm release:dry-run` — validate exactly the public package allowlist
// against npmjs.org without publishing anything.

import { execFileSync } from "node:child_process";

const publicPackages = [
  "shared",
  "core",
  "formula-engine",
  "commands",
  "history",
  "import-export",
  "plugin-api",
  "runtime",
];

for (const directory of publicPackages) {
  console.log(`\nDry-running packages/${directory}`);
  execFileSync(
    "npm",
    [
      "publish",
      "--dry-run",
      "--access",
      "public",
      "--tag",
      "next",
      "--registry",
      "https://registry.npmjs.org/",
    ],
    {
      cwd: new URL(`../packages/${directory}/`, import.meta.url),
      stdio: "inherit",
    },
  );
}

console.log("\nRelease dry-run passed for 8 public packages.");
