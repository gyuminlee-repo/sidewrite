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
  // One handle per job: provider and model are fixed when the job starts.
  assert.match(code, /oneshotJob: \(key\) => oneshotRouter\.pin\(key\)/);
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

// Codex review of stage 2.5: the daemon uses a provider only once the pane
// confirmed it in this connection.
test("judgments and review turns wait for the pane's confirmed provider; hello starts over", () => {
  // The router asks the selection, which waits and throws: never the sync default.
  assert.match(code, /getChoice: \(key\) => selection\.ready\(key, "oneshot"\)/);
  assert.match(code, /turnGate: \(key\) => selection\.ready\(key, "turn"\)/);
  // A new connection has no confirmed choice.
  const hello = code.slice(code.indexOf("async function onPaneConnect"));
  assert.match(hello.slice(0, 300), /selection\.begin\(key\)/);
  // The stored choices are restored against the probe, by their own message.
  assert.match(code, /restore_providers: async/);
  assert.match(code, /selection\.restore\(\s*key,/);
});

// Stage 3: the response letter service.
test("the response service is wired to the same one-shot router and its handlers reach the bridge", () => {
  const at = code.indexOf("const response = createResponseService({");
  assert.ok(at > 0);
  const block = code.slice(at, code.indexOf("});", at));
  assert.match(block, /getContext: \(key\) => bridge\.getContext\(key\)/);
  // The job handle of the pane's provider; the service names its Claude default.
  assert.match(block, /oneshotJob: \(key, options\) => oneshotRouter\.pin\(key, options\)/);
  assert.match(block, /sendEvent: \(key, obj\) => bridge\.sendToTaskpane\(obj, key\)/);
  // No runner of its own: a draft never goes around the pane's choice.
  assert.equal(/makeOneshot|runOneshot|ONESHOT_MODEL/.test(block), false);
  assert.match(code, /\.\.\.review\.handlers,\s*\.\.\.response\.handlers,/);
});

test("a Review refresh marks saved responses stale with the rule of the Response tab", () => {
  const at = code.indexOf("const review = createReviewService({");
  const block = code.slice(at, code.indexOf("\n});", at));
  assert.match(block, /afterRefresh: \(rev, snap\) => markStale\(rev, snap\)/);
  assert.match(code, /import \{ markStale \} from "\.\/response-quotes\.mjs";/);
});

test("set_model: the model the session runs on, read after the wait, decides whether it restarts", () => {
  const at = code.indexOf("set_model: async");
  const body = code.slice(at, at + 1400);
  // The session's own model, read once the list is known.
  assert.match(body, /const after = await modelCache\.get\(\);/);
  assert.match(body, /running: sessionFor\(key\)\?\.model \?\? null/);
  assert.ok(body.indexOf("await modelCache.get()") < body.indexOf("sessionFor(key)?.model"));
  assert.match(body, /if \(restart\)/);
  assert.equal(/modelCache\.current\(\)/.test(body), false, "no guess from the list as it was");
  // The restart that corrects the stored model on the first connection keeps
  // a first message that is still waiting; a choice made later does not.
  assert.match(
    body,
    /restartSession\(key, host, \{\s*reason: "model_changed",\s*keepQueued: prev === undefined,?\s*\}\)/,
  );
  // The session records the model it is started on, where query() gets it.
  assert.match(code, /session\.model = modelArgFor\(key\);/);
  assert.match(code, /model: session\.model,/);
  // Its definition, the session start and getModel: query() reads it nowhere else.
  assert.equal((code.match(/modelArgFor\(key\)/g) ?? []).length, 3);
});

test("a restart that keeps waiting input does not drop it or end its turn", () => {
  assert.match(code, /q\.latest = \{[^}]*keepQueued[^}]*\}/);
  assert.match(code, /keepQueued: req\.keepQueued/);
  const at = code.indexOf("async function startSessionForFolder");
  const body = code.slice(at, at + 900);
  assert.match(body, /const kept = bridge\.clearUserMessages\(key, \{ keepQueued \}\);/);
  // A message still waiting belongs to the next session: its turn is not over.
  assert.match(
    body,
    /if \(!kept\) endTurnState\(key, \{ ok: false, error: "The session was restarted\." \}\);/,
  );
  const restart = code.slice(code.indexOf("async function restartSession"));
  assert.match(restart.slice(0, 700), /\{ replay: false, keepQueued \}/);
});
