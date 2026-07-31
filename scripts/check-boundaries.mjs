// Dependency boundary enforcement (architecture guard, ADR-0002).
// Two layers, both must hold:
//   1. package.json dependencies of each package ⊆ allowlist
//   2. actual `import ... from "@opensheet/x"` statements in src/ ⊆ allowlist
//
// TypeScript's typechecker does NOT prevent architectural reverse imports;
// this script does. Runs in CI (`pnpm check:boundaries`).

import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, relative } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));

/** Allowlist: package → @opensheet/* packages it may depend on. */
const ALLOWED = {
  shared: [],
  core: ["shared"],
  commands: ["shared", "core"],
  history: ["shared", "commands"],
  "formula-engine": ["shared", "core"],
  clipboard: ["shared", "core"],
  "import-export": ["shared"],
  "plugin-api": ["shared"],
  "renderer-canvas": ["shared", "core", "plugin-api"],
  react: ["shared", "renderer-canvas", "plugin-api", "runtime"],
  runtime: ["shared", "core", "commands", "history", "plugin-api", "formula-engine", "clipboard", "import-export"],
};

const violations = [];

for (const [pkg, allowed] of Object.entries(ALLOWED)) {
  const dir = join(root, "packages", pkg);
  const manifest = JSON.parse(readFileSync(join(dir, "package.json"), "utf8"));

  // Layer 1: manifest
  for (const section of ["dependencies", "devDependencies", "peerDependencies"]) {
    for (const dep of Object.keys(manifest[section] ?? {})) {
      if (dep.startsWith("@opensheet/")) {
        const name = dep.slice("@opensheet/".length);
        if (!allowed.includes(name)) {
          violations.push(`${pkg}/package.json ${section}: "${dep}" is not allowed (allowed: ${allowed.join(", ") || "none"})`);
        }
      }
    }
  }

  // Layer 2: source imports (test files are not shipped and may import
  // anything in the workspace for fixtures)
  const importRe = /(?:import|export)[^"']*from\s+["'](@opensheet\/[a-z-]+)["']/g;
  for (const file of walk(join(dir, "src"))) {
    if (file.includes("__tests__") || /\.(test|spec)\.tsx?$/.test(file)) continue;
    const text = readFileSync(file, "utf8");
    for (const match of text.matchAll(importRe)) {
      const name = match[1].slice("@opensheet/".length);
      if (!allowed.includes(name)) {
        violations.push(`${relative(root, file)}: imports "${match[1]}" which is not allowed for ${pkg}`);
      }
    }
  }
}

function* walk(dir) {
  let entries;
  try {
    entries = readdirSync(dir);
  } catch {
    return; // package without src yet
  }
  for (const entry of entries) {
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) yield* walk(full);
    else if (/\.(ts|tsx)$/.test(entry)) yield full;
  }
}

if (violations.length > 0) {
  console.error("Dependency boundary violations:\n");
  for (const v of violations) console.error(`  ✗ ${v}`);
  process.exit(1);
}
console.log(`Boundary check passed (${Object.keys(ALLOWED).length} packages).`);
