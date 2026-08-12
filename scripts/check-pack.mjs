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
    ["--filter", "@injoysai/opensheet...", "-r", "run", "build"],
    {
      cwd: root,
      stdio: "inherit",
    },
  );
  execFileSync(
    "pnpm",
    [
      "--filter",
      "@injoysai/opensheet...",
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
      const tarballPrefix = `${manifest.name.replace(/^@/, "").replace("/", "-")}-`;
      const tarball = packed.find((name) => name.startsWith(tarballPrefix));
      assert(
        tarball !== undefined,
        `Missing packed tarball for ${manifest.name}`,
      );
      const packedManifest = JSON.parse(
        execFileSync(
          "tar",
          ["-xOf", join(tarballs, tarball), "package/package.json"],
          {
            encoding: "utf8",
          },
        ),
      );
      const packedFiles = execFileSync(
        "tar",
        ["-tf", join(tarballs, tarball)],
        {
          encoding: "utf8",
        },
      );
      assert(
        packedFiles.split("\n").includes("package/LICENSE"),
        `${manifest.name} tarball must include MIT LICENSE`,
      );
      assert.equal(
        packedManifest.private,
        undefined,
        `${manifest.name} must be publishable`,
      );
      assert.equal(
        packedManifest.publishConfig?.registry,
        "https://registry.npmjs.org/",
        `${manifest.name} must target npmjs.org`,
      );
      for (const [dependency, version] of Object.entries(
        packedManifest.dependencies ?? {},
      )) {
        if (!dependency.startsWith("@injoysai/opensheet")) continue;
        assert.equal(
          version,
          "0.1.0",
          `${manifest.name} tarball must pin ${dependency} to 0.1.0`,
        );
      }
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
import { createOpenSheet } from "@injoysai/opensheet";

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
    const { api, workbook } = window.__opensheetPackSmoke;
    const sheetId = workbook.activeSheetId;

    await api.applyOperations({
      workbookId: workbook.id,
      sheetId,
      atomic: true,
      operations: [
        { type: "cell.set", range: "A1", value: 2 },
        { type: "cell.set", range: "A2", value: 1 },
        { type: "cell.set", range: "A3", value: 2 },
        { type: "formula.set", range: "B1", formula: "=A1*2" },
        {
          type: "filter.apply",
          spec: {
            range: { startRow: 0, startCol: 0, endRow: 2, endCol: 0 },
            hasHeader: false,
            conditions: [{ columnOffset: 0, operator: "equals", value: 2 }],
          },
        },
      ],
    });
    const filterRows = Array.from(
      api.getFilterProjectionState(sheetId).visibleRows ?? [],
    );

    await api.applyOperations({
      workbookId: workbook.id,
      sheetId,
      atomic: true,
      operations: [{ type: "filter.clear" }],
    });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId,
      atomic: true,
      operations: [
        {
          type: "range.sort",
          spec: {
            range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
            hasHeader: false,
            keys: [{ columnOffset: 0, direction: "asc" }],
          },
        },
      ],
    });
    const sorted = api.readRange({ sheetId, range: "A1:B3" });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId,
      atomic: true,
      operations: [
        {
          type: "range.dedupe",
          spec: {
            range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
            hasHeader: false,
            keyColumnOffsets: [0],
            keep: "first",
          },
        },
      ],
    });
    const deduped = api.readRange({ sheetId, range: "A1:B3" });

    await api.usePlugin({
      id: "pack-smoke-plugin",
      setup(context) {
        context.commands.registerCommand({
          id: "pack-smoke.mark",
          execute: () => [
            { type: "cell.set", range: "D1", value: "plugin-command" },
          ],
        });
        context.functions.registerFunction({
          name: "PACK_DOUBLE",
          minArgs: 1,
          maxArgs: 1,
          execute: ([value]) =>
            typeof value === "number" ? value * 2 : { type: "#VALUE!" },
        });
      },
    });
    await api.executePluginCommand({
      workbookId: workbook.id,
      sheetId,
      commandId: "pack-smoke.mark",
      payload: null,
    });
    await api.applyOperations({
      workbookId: workbook.id,
      sheetId,
      atomic: true,
      operations: [
        { type: "formula.set", range: "C1", formula: "=PACK_DOUBLE(A2)" },
      ],
    });
    const plugin = api.readRange({ sheetId, range: "C1:D1" });

    const imported = await api.importCSV({
      file: new File(["Name,Amount\r\nAda,10"], "sales.csv", {
        type: "text/csv",
      }),
    });
    const exported = await api.exportCSV({ sheetId: imported.sheetId });
    return {
      filterRows,
      sorted,
      deduped,
      plugin,
      imported,
      values: api.readRange({ sheetId: imported.sheetId, range: "A1:B2" }),
      csv: await exported.text(),
    };
  });
  assert.deepEqual(result.filterRows, [0, 2]);
  assert.deepEqual(result.sorted, [
    [1, null],
    [2, 4],
    [2, null],
  ]);
  assert.deepEqual(result.deduped, [
    [1, null],
    [2, 4],
    [null, null],
  ]);
  assert.deepEqual(result.plugin, [[4, "plugin-command"]]);
  assert.deepEqual(result.imported.rowCount, 2);
  assert.deepEqual(result.imported.columnCount, 2);
  assert.deepEqual(result.values, [
    ["Name", "Amount"],
    ["Ada", "10"],
  ]);
  assert.equal(result.csv, "Name,Amount\r\nAda,10");
  console.log(
    "Pack smoke passed: tarball consumer executed Formula, Filter, Sort, Dedupe, plugin, and CSV Worker flows.",
  );
} finally {
  await browser?.close();
  await server?.close();
  process.chdir(root);
  rmSync(temp, { recursive: true, force: true });
}
