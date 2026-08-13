// `pnpm release:dry-run` — validate exactly the public package allowlist
// against npmjs.org without publishing anything. Use pnpm publish here:
// pnpm rewrites workspace: dependencies in the packed manifest; npm publish
// does not.

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
    "pnpm",
    [
      "publish",
      "--dry-run",
      "--access",
      "public",
      "--tag",
      "next",
      "--registry",
      "https://registry.npmjs.org/",
      "--no-git-checks",
    ],
    {
      cwd: new URL(`../packages/${directory}/`, import.meta.url),
      stdio: "inherit",
    },
  );
}

console.log("\nRelease dry-run passed for 8 public packages.");
