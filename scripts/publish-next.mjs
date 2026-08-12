// `pnpm release:publish-next -- 0.1.0-next.0` — publish the fixed
// public package allowlist in dependency order. The explicit version argument
// prevents an accidental release of a different manifest version.

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const registry = "https://registry.npmjs.org/";
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

const manifests = publicPackages.map((directory) => ({
  directory,
  manifest: JSON.parse(
    readFileSync(
      new URL(`../packages/${directory}/package.json`, import.meta.url),
      "utf8",
    ),
  ),
}));
const version = manifests[0].manifest.version;
const confirmation = process.argv[2];

if (!/^\d+\.\d+\.\d+-next\.\d+$/.test(version)) {
  throw new Error(`Expected a next prerelease version, received ${version}`);
}
if (confirmation !== version) {
  throw new Error(
    `Refusing publish. Run: pnpm release:publish-next -- ${version}`,
  );
}

for (const { manifest } of manifests) {
  if (manifest.version !== version) {
    throw new Error(
      `${manifest.name} is ${manifest.version}; expected ${version}`,
    );
  }
  for (const [name, dependencyVersion] of Object.entries(
    manifest.dependencies ?? {},
  )) {
    if (!name.startsWith("@injoysai/opensheet")) continue;
    if (dependencyVersion !== `workspace:${version}`) {
      throw new Error(
        `${manifest.name} depends on ${name}@${dependencyVersion}; expected workspace:${version}`,
      );
    }
  }
}

const worktree = execFileSync("git", ["status", "--porcelain"], {
  cwd: root,
  encoding: "utf8",
}).trim();
if (worktree !== "") {
  throw new Error(
    "Refusing publish from a dirty worktree. Commit or stash first.",
  );
}

execFileSync(process.execPath, ["scripts/publish-dry-run.mjs"], {
  cwd: root,
  stdio: "inherit",
});

for (const { directory, manifest } of manifests) {
  console.log(`\nPublishing ${manifest.name}@${version}`);
  execFileSync(
    "npm",
    ["publish", "--access", "public", "--tag", "next", "--registry", registry],
    {
      cwd: new URL(`../packages/${directory}/`, import.meta.url),
      stdio: "inherit",
    },
  );
}

console.log(`\nPublished 8 packages at ${version} with the next tag.`);
