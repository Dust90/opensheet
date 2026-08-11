// `pnpm check:pack` — prove published runtime tarballs work in a fresh Vite
// consumer, including the separately emitted CSV Worker asset.

import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "@playwright/test";

const root = fileURLToPath(new URL("..", import.meta.url));
const demoRequire = createRequire(
  new URL("../apps/demo/package.json", import.meta.url),
);
const { build, createServer } = await import(demoRequire.resolve("vite"));
const temp = mkdtempSync(join(tmpdir(), "opensheet-pack-smoke-"));
const tarballs = join(temp, "tarballs");
const consumer = join(temp, "consumer");
mkdirSync(consumer);

const runtimePackages = [
  "shared",
  "core",
  "commands",
  "history",
  "formula-engine",
  "import-export",
  "plugin-api",
  "runtime",
];

let browser;
let server;
try {
  execFileSync(
    "pnpm",
    ["--filter", "@opensheet/runtime...", "-r", "run", "build"],
    {
      cwd: root,
      stdio: "inherit",
    },
  );
  execFileSync(
    "pnpm",
    [
      "--filter",
      "@opensheet/runtime...",
      "-r",
      "pack",
      "--pack-destination",
      tarballs,
    ],
    {
      cwd: root,
      stdio: "inherit",
    },
  );

  const packed = readdirSync(tarballs).filter((name) => name.endsWith(".tgz"));
  assert.equal(
    packed.length,
    runtimePackages.length,
    "runtime pack should include exactly its workspace dependency closure",
  );
  const localPackages = Object.fromEntries(
    runtimePackages.map((directory) => {
      const manifest = JSON.parse(
        readFileSync(join(root, "packages", directory, "package.json"), "utf8"),
      );
      const tarball = packed.find((name) =>
        name.startsWith(`opensheet-${directory}-`),
      );
      assert(
        tarball !== undefined,
        `Missing packed tarball for ${manifest.name}`,
      );
      return [manifest.name, `file:${join(tarballs, tarball)}`];
    }),
  );
  writeFileSync(
    join(consumer, "package.json"),
    JSON.stringify(
      {
        name: "opensheet-pack-smoke",
        private: true,
        type: "module",
        dependencies: localPackages,
      },
      null,
      2,
    ),
  );
  const overrides = Object.entries(localPackages)
    .map(([name, target]) => `  '${name}': '${target}'`)
    .join("\n");
  writeFileSync(
    join(consumer, "pnpm-workspace.yaml"),
    `overrides:\n${overrides}\n`,
  );
  writeFileSync(
    join(consumer, "index.html"),
    '<div id="app">loading</div><script type="module" src="/main.js"></script>',
  );
  writeFileSync(
    join(consumer, "main.js"),
    `
import { createOpenSheet } from "@opensheet/runtime";

const api = createOpenSheet();
const workbook = api.createWorkbook({ name: "Pack smoke" });
window.__opensheetPackSmoke = { api, workbook };
document.querySelector("#app").textContent = "ready";
`,
  );

  execFileSync("pnpm", ["install", "--offline"], {
    cwd: consumer,
    stdio: "inherit",
  });

  process.chdir(consumer);
  await build({ logLevel: "error" });
  const assets = readdirSync(join(consumer, "dist", "assets"));
  assert(
    assets.some(
      (name) => name.startsWith("csv.worker-") && name.endsWith(".js"),
    ),
    "Vite build must emit csv.worker asset",
  );

  server = await createServer({
    logLevel: "error",
    server: { host: "127.0.0.1", port: 0 },
  });
  await server.listen();
  const url = server.resolvedUrls?.local[0];
  assert(url !== undefined, "Vite did not provide a local URL");

  browser = await chromium.launch({ channel: "chromium" });
  const page = await browser.newPage();
  await page.goto(url);
  await page.waitForFunction(
    () => document.querySelector("#app")?.textContent === "ready",
  );
  const result = await page.evaluate(async () => {
    const { api } = window.__opensheetPackSmoke;
    const imported = await api.importCSV({
      file: new File(["Name,Amount\r\nAda,10"], "sales.csv", {
        type: "text/csv",
      }),
    });
    const exported = await api.exportCSV({ sheetId: imported.sheetId });
    return {
      imported,
      values: api.readRange({ sheetId: imported.sheetId, range: "A1:B2" }),
      csv: await exported.text(),
    };
  });
  assert.deepEqual(result.imported.rowCount, 2);
  assert.deepEqual(result.imported.columnCount, 2);
  assert.deepEqual(result.values, [
    ["Name", "Amount"],
    ["Ada", "10"],
  ]);
  assert.equal(result.csv, "Name,Amount\r\nAda,10");
  console.log(
    "Pack smoke passed: tarball consumer built and executed CSV Worker import/export.",
  );
} finally {
  await browser?.close();
  await server?.close();
  process.chdir(root);
  rmSync(temp, { recursive: true, force: true });
}
