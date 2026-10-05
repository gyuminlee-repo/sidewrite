import { test } from "node:test";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import * as path from "node:path";
import { tmpdir } from "node:os";
import { fileURLToPath } from "node:url";
import { runInNewContext } from "node:vm";

const here = fileURLToPath(new URL("../scripts/e2e/", import.meta.url));
const source = fs.readFileSync(
  process.env.E2E_SETUP_SOURCE || path.join(here, "run-acceptance.mjs"),
  "utf8",
);
const start = source.indexOf("async function setup() {\n");
const end = source.indexOf("  // --resident:", start);
assert.ok(start >= 0 && end > start, "setup copy fragment boundaries exist");
const fragment = source.slice(start + "async function setup() {\n".length, end);

function checkDependencies(tools, name, visited = new Set()) {
  if (visited.has(name)) return;
  visited.add(name);
  const deployed = fs.readFileSync(path.join(tools, name));
  assert.deepEqual(deployed, fs.readFileSync(path.join(here, name)), `${name}: exact copy`);
  if (!name.endsWith(".ps1")) return;
  for (const line of deployed.toString().split(/\r?\n/)) {
    if (!/^\s*\.\s+/.test(line)) continue;
    const dependency = line.match(/^\s*\. \(Join-Path \$PSScriptRoot "([\w.-]+\.ps1)"\)\s*$/);
    assert.ok(dependency, `${name}: unsupported dot-source: ${line}`);
    assert.ok(
      fs.existsSync(path.join(tools, dependency[1])),
      `${name} requires deployed ${dependency[1]}`,
    );
    checkDependencies(tools, dependency[1], visited);
  }
}

for (const stage of ["s2", "s3", "cite", "parts"]) {
  test(`setup deploys complete Word helper dependencies for ${stage}`, (t) => {
    const evidence = process.env.E2E_SETUP_EVIDENCE_DIR;
    if (evidence) fs.mkdirSync(evidence, { recursive: true });
    const repo = fs.mkdtempSync(path.join(evidence || tmpdir(), `setup-${stage}-`));
    if (!evidence) t.after(() => fs.rmSync(repo, { recursive: true }));
    const tools = path.join(repo, "scripts", "e2e");
    // Execute only setup's deployment prefix, before any services or Word calls.
    runInNewContext(
      fragment,
      { fs, path, HERE: here, TOOLS: tools, STAGE: stage, opt: {} },
      {
        timeout: 1000,
      },
    );
    for (const name of [
      "pane-driver.cjs",
      "word-com.ps1",
      "word-rot.ps1",
      "word-launch.ps1",
      "word-shot.ps1",
    ]) {
      assert.ok(fs.existsSync(path.join(tools, name)), `${stage}: ${name} deployed`);
    }
    assert.equal(fs.existsSync(path.join(tools, "marked-diff-word.ps1")), stage === "cite");
    for (const name of fs.readdirSync(tools)) checkDependencies(tools, name);
    t.diagnostic(`dependency closure verified in ${tools}`);
  });
}
