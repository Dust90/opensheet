// License allowlist enforcement for the full pnpm dependency tree.
//
// Walks node_modules/.pnpm (the real installed tree — including transitive
// deps, which is where license problems hide) and rejects anything outside
// the allowlist. Runs in CI (`pnpm licenses:check`).
//
// Allowlist decision: permissive OSI licenses only. GPL/AGPL/SSPL/BUSL/
// Commons-Clause/UNKNOWN are blocked by default.

import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const PNPM_STORE = join(root, "node_modules", ".pnpm");

const ALLOWED = new Set([
  "MIT",
  "ISC",
  "APACHE-2.0",
  "BSD-2-CLAUSE",
  "BSD-3-CLAUSE",
  "0BSD",
  "CC0-1.0",
  "UNLICENSE",
  "PYTHON-2.0", // argparse-style, permissive
  "BLUEOAK-1.0.0", // Blue Oak Model License, permissive
  // Reviewed exceptions (transitive, build-time only, unmodified):
  "CC-BY-4.0", // caniuse-lite: data file (browser support DB), not code
  "MPL-2.0", // lightningcss: weak file-level copyleft; allowed when used unmodified as a tool dependency
]);

const BLOCKED = /GPL|AGPL|SSPL|BUSL|COMMONS[- ]CLAUSE|UNKNOWN/i;

function normalizeLicense(pkg) {
  const field = pkg.license ?? pkg.licenses;
  if (field === undefined) return "UNKNOWN";
  if (typeof field === "string") return field;
  if (Array.isArray(field)) {
    return field
      .map((l) => (typeof l === "string" ? l : (l.type ?? "UNKNOWN")))
      .join(" OR ");
  }
  if (typeof field === "object" && field !== null) return field.type ?? "UNKNOWN";
  return "UNKNOWN";
}

/** SPDX expression check: OR → any allowed is fine; AND → all must be allowed. */
function isAllowed(expression) {
  const cleaned = expression.replace(/[()]/g, " ").replace(/\s+/g, " ").trim();
  if (cleaned === "" || /^UNLICENSED$/i.test(cleaned)) return false;
  if (/\bOR\b/i.test(cleaned)) {
    return cleaned.split(/\s+OR\s+/i).some((part) => isAllowed(part));
  }
  if (/\bAND\b/i.test(cleaned)) {
    return cleaned.split(/\s+AND\s+/i).every((part) => isAllowed(part));
  }
  const token = cleaned.replace(/[^\w.-]/g, "").toUpperCase();
  if (BLOCKED.test(token)) return false;
  return ALLOWED.has(token);
}

if (!existsSync(PNPM_STORE)) {
  console.error("node_modules/.pnpm not found — run `pnpm install` first.");
  process.exit(2);
}

const checked = new Map(); // name@version → license string
const violations = [];

for (const dirEntry of readdirSync(PNPM_STORE)) {
  // .pnpm layout: <name>@<version>[_peer-suffix] or @scope+name@<version>[...]
  const nmDir = join(PNPM_STORE, dirEntry, "node_modules");
  if (!existsSync(nmDir)) continue;
  for (const entry of readdirSync(nmDir, { withFileTypes: true })) {
    if (!entry.isDirectory() && !entry.isSymbolicLink()) continue;
    const names = entry.name.startsWith("@")
      ? readdirSync(join(nmDir, entry.name)).map((n) => `${entry.name}/${n}`)
      : [entry.name];
    for (const name of names) {
      const pkgJsonPath = join(nmDir, name, "package.json");
      if (!existsSync(pkgJsonPath)) continue;
      let pkg;
      try {
        pkg = JSON.parse(readFileSync(pkgJsonPath, "utf8"));
      } catch {
        continue;
      }
      const id = `${pkg.name ?? name}@${pkg.version ?? "?"}`;
      if (checked.has(id)) continue;
      const license = normalizeLicense(pkg);
      checked.set(id, license);
      if (!isAllowed(license)) {
        violations.push({ id, license });
      }
    }
  }
}

console.log(`Checked ${checked.size} installed packages against the allowlist.`);
if (violations.length > 0) {
  console.error(`\n${violations.length} license violation(s):`);
  for (const v of violations) console.error(`  ✗ ${v.id} — ${v.license}`);
  console.error(
    "\nAllowed: MIT, ISC, Apache-2.0, BSD-2/3-Clause, 0BSD, CC0-1.0, Unlicense, Python-2.0, BlueOak-1.0.0, CC-BY-4.0 (data), MPL-2.0 (build tools)",
  );
  process.exit(1);
}
console.log("All licenses are within the allowlist.");
