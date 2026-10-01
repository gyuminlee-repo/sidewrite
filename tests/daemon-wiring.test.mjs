// index.mjs starts servers when it is imported, so its wiring is checked on
// its source text: the few facts the provider work depends on.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const src = readFileSync(new URL("../daemon/index.mjs", import.meta.url), "utf8");
const code = src
  .split("\n")
  .filter((l) => !l.trim().startsWith("//"))
  .join("\n");

test("the HTTP server, which carries /mcp, listens on loopback only", () => {
  assert.match(code, /http\.listen\(HTTP_PORT, "127\.0\.0\.1"/);
  assert.equal(/http\.listen\(HTTP_PORT,\s*\(/.test(code), false, "no listen on every interface");
});

test("/mcp is handled before the CORS headers, OPTIONS and the bridge token of the pane files", () => {
  const mcp = code.indexOf('mcpPath.startsWith("/mcp/")');
  assert.ok(mcp > 0);
  for (const later of [
    "Access-Control-Allow-Origin",
    'req.method === "OPTIONS"',
    '"/bridge-token"',
  ]) {
    assert.ok(code.indexOf(later) > mcp, later);
  }
  // A path under /mcp never falls through to the file server.
  assert.match(code.slice(mcp, mcp + 600), /res\.writeHead\(404\)\.end\(\);\s*return;/);
});

test("the endpoint serves the review service's recorder and the daemon's own review turns", () => {
  assert.match(code, /getReviewTurn: \(key\) => reviewTurnByKey\.get\(key\) \?\? null/);
  assert.match(code, /trackedEditDeps: \(key\) => review\.trackedEditDeps\(key\)/);
});

test("judgments go through the router; stop and shutdown reach a CLI turn", () => {
  assert.match(code, /oneshotFor: \(key\) => oneshotRouter\.forKey\(key\)/);
  assert.match(code, /oneshotModelFor: \(key\) => oneshotRouter\.modelLabel\(key\)/);
  assert.match(code, /if \(cliTurns\.stop\(key\)\)/);
  assert.match(code, /cliTurns\.shutdown\(\)/);
  assert.match(code, /process\.on\("SIGTERM", shutdown\)/);
});

test("no message registers the office tools with agy outside a turn, and no flag skips an approval", () => {
  assert.equal(/agy_mcp_connect|agyMcpAddArgs/.test(code), false);
  assert.equal(
    /dangerously|skip-permissions|accept-edits|approval_policy|bypass/i.test(code),
    false,
  );
});

test("a provider key is passed to the key store and nowhere else", () => {
  const uses = code.split("\n").filter((l) => /msg\.key\b/.test(l));
  assert.deepEqual(
    uses.map((l) => l.trim()),
    ['await providerKeys.set(String(msg.id ?? ""), { key: msg.key, base_url: msg.base_url });'],
  );
});
