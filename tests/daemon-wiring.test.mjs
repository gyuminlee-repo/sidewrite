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

// The office tools are reached through the pane's own bridge and nothing
// else: the HTTP server serves the pane files and has no tool endpoint.
test("the HTTP server has no MCP route and the daemon starts no CLI with the office tools", () => {
  assert.equal(/["'`]\/mcp\b/.test(code), false, "no route under /mcp");
  assert.equal(/mcp_servers|StreamableHTTP/.test(code), false);
  const imports = [...code.matchAll(/from "\.\/([\w/-]+)\.mjs"/g)].map((m) => m[1]);
  assert.ok(imports.includes("proposal-turns"));
  assert.deepEqual(
    imports.filter((m) => /turns|mcp/.test(m)),
    ["proposal-turns"],
    "the proposal runner is the only turn runner next to the Claude session",
  );
});

test("judgments go through the router; shutdown stops a proposal turn", () => {
  // One handle per job: provider and model are fixed when the job starts.
  assert.match(code, /oneshotJob: \(key\) => oneshotRouter\.pin\(key\)/);
  assert.match(code, /proposalTurns\.shutdown\(\)\.then\(bye, bye\)/);
  assert.match(code, /process\.on\("SIGTERM", shutdown\)/);
});

test("no message registers the office tools with agy outside a turn, and no flag skips an approval", () => {
  assert.equal(/agy_mcp_connect|mcp["', ]+add/.test(code), false);
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

// Proposal turns (Codex): the review service decides per request whether the
// pane's provider proposes, the turn it queues carries the choice, and
// pushTurn starts the runner for it. A Stop reaches a request that has not
// become a turn yet, then the runner, then the Claude session.
test("a review turn of a proposing provider goes to the proposal runner; Stop and shutdown reach it", () => {
  assert.match(
    code,
    /const proposalTurns = proposalTurnsFor\(\{\s*bridge,\s*review,\s*router: oneshotRouter,\s*endTurn: \(key, outcome\) => turnState\.end\(key, PROPOSAL, outcome\),?\s*\}\);/,
  );
  // The mode is read from the registry, for the choice of the pane.
  const dep = code.slice(code.indexOf("proposalPreview: async (key, ref, source) =>"));
  assert.match(dep.slice(0, 700), /selection\.choiceFor\(key, "turn"\)/);
  assert.match(dep.slice(0, 700), /turnMode === "agent"\) return null;/);
  assert.match(dep.slice(0, 700), /turnMode !== "proposal"\) \{\s*throw new Error\(/);
  assert.match(
    dep.slice(0, 900),
    /proposalTurns\.preview\(key, \{ ref \}, \{ \.\.\.source, choice \}\)/,
  );
  const push = code.slice(code.indexOf("pushTurn: (key, host, text, turn) =>"));
  assert.match(push.slice(0, 1500), /if \(!turn\.choice\) \{/);
  assert.match(
    push.slice(0, 1500),
    /proposalTurns\s*\.start\(key, host, turn, \{\s*at: review\.at\(key\),\s*selected: turn\.selected,\s*choice: turn\.choice,?\s*\}\)/,
  );
  const stop = code.slice(code.indexOf("stop_agent: async"));
  assert.match(
    stop.slice(0, 900),
    /if \(review\.cancelStarting\(key\) \|\| proposalTurns\.stop\(key\)\) \{/,
  );
  assert.ok(stop.indexOf("review.cancelStarting(key)") < stop.indexOf("sessionFor(key)"));
  assert.match(code, /proposalTurns\.shutdown\(\)/);
});
