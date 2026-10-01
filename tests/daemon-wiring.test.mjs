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

test("the HTTP server listens on loopback only", () => {
  assert.match(code, /http\.listen\(HTTP_PORT, "127\.0\.0\.1"/);
  assert.equal(/http\.listen\(HTTP_PORT,\s*\(/.test(code), false, "no listen on every interface");
});

test("judgments go through the router", () => {
  // One handle per job: provider and model are fixed when the job starts.
  // (options: a job may name its role and Claude model; the paragraph check
  // of the Review tab asks for those of the edits.)
  assert.equal(
    code.match(/oneshotJob: \(key, options\) => oneshotRouter\.pin\(key, options\)/g).length,
    2,
  );
  assert.match(code, /process\.on\("SIGTERM", shutdown\)/);
});

// Codex and agy are for judgments only: the daemon has no runner that gives a
// CLI the office tools, no HTTP route for such a run, and a review turn goes
// to the pane's Claude session and nowhere else.
test("a review turn goes to the Claude session only: no CLI runner and no tool route for one", () => {
  assert.equal(/createServer\(async \(req, res\) => \{\s*const \w+Path/.test(code), false);
  assert.equal(
    /startsWith\("\/\w+\/"\)/.test(code),
    false,
    "no route of its own on the HTTP server",
  );
  const at = code.indexOf("pushTurn: (key, host, text, turn) => {");
  assert.ok(at > 0);
  const body = code.slice(at, code.indexOf("oneshot: makeOneshot()", at));
  assert.match(body, /bridge\.pushUserMessage\(text, key\);/);
  assert.equal(/\.start\(/.test(body), false, "nothing else is started for a review turn");
  assert.equal(/choice/.test(body), false);
});

test("no message registers the office tools with agy outside a turn, and no flag skips an approval", () => {
  assert.equal(/agy_mcp_connect|"mcp",\s*"add"/.test(code), false);
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
  assert.match(code, /getChoice: \(key, role\) => selection\.ready\(key, role\)/);
  // Writing without a model of its own: the best Claude model of the account.
  assert.match(
    code,
    /claudeDefault: \(role\) =>\s*role === "writing" \? bestClaudeModel\(modelCache\.current\(\)\) : ONESHOT_MODEL/,
  );
  // The stored writing choice is restored with the other two.
  assert.match(
    code,
    /\{ turn: pick\(msg\.turn\), writing: pick\(msg\.writing\), oneshot: pick\(msg\.oneshot\) \}/,
  );
  // A session with no model chosen runs on the best one, not on sonnet.
  assert.match(
    code,
    /return resolveModelChoice\(modelByKey\.get\(key\), modelCache\.current\(\)\);/,
  );
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
  const body = code.slice(at, at + 1800);
  // The session's own model, read once the list is known.
  assert.match(body, /const after = await modelCache\.get\(\);/);
  assert.match(
    body,
    // (A session left on another model because this very choice declines the
    // manuscript counts as running on the choice; another choice restarts it.)
    /const chosen = resolveModelChoice\(requested, after\);[\s\S]{0,400}running: staysOnFallback\(declinedByKey\.get\(key\), chosen\)\s*\? chosen\s*: \(sessionFor\(key\)\?\.fellBackTo \?\? sessionFor\(key\)\?\.model \?\? null\)/,
  );
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
  // Its definition, the session start, getModel, getRunModel, staysOn and one log line: query() reads it nowhere else.
  assert.equal((code.match(/modelArgFor\(key\)/g) ?? []).length, 6);
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

// What runs is told to the pane as it runs, not as it was chosen.
test("the pane is told what each request ran on: one-shots by the router, the session by its init", () => {
  assert.match(
    code,
    /onUsed: \(key, used\) =>\s*bridge\.sendToTaskpane\(\{ type: "model_used", \.\.\.used \}, key\)/,
  );
  // The session: the model query() was given and the one the CLI says it runs.
  const at = code.indexOf('event: "session_init"');
  const ev = code.slice(at, at + 200);
  assert.match(ev, /model: msg\.model,/);
  assert.match(ev, /asked: session\?\.model \?\? null,/);
  // No option of the session or of a one-shot lowers thinking or effort.
  assert.equal(/\b(thinking|effort|maxThinkingTokens)\s*:/.test(code), false);
});

test("a refusal fallback of the CLI is logged, told to the pane, and recorded as the model the session runs on", () => {
  const at = code.indexOf('msg.subtype === "model_refusal_fallback"');
  assert.ok(at > 0);
  const body = code.slice(at, at + 1200);
  assert.match(body, /const f = refusalFallback\(msg\);/);
  assert.match(body, /console\.log\(f\.log\);/);
  // set_model compares the model the session runs on with the one the user
  // chose: with the fallback recorded, choosing the model again restarts the
  // session on it. session.model stays what was asked for.
  assert.match(body, /session\.fellBackTo = msg\.fallback_model;/);
  // The daemon itself names no fallback model.
  assert.equal(/fallbackModel/.test(code), false);
});

// After a refusal fallback the CLI keeps the session on the other model. The
// retry inside that turn stays (and is shown); the next turn starts on the
// model the user chose again: the session is restarted, resuming, before the
// new message is queued, and the old session is given no further message.
test("the turn after a refusal fallback goes back to the chosen model, and its message is not lost", () => {
  const at = code.indexOf("async function ensureLoopForMessage(key, host, { idle = true } = {})");
  assert.ok(at > 0);
  const body = code.slice(at, at + 1100);
  assert.match(
    body,
    /if \(live && !live\.settled\) \{\s*if \(!live\.fellBackTo \|\| !idle \|\| staysOn\(key\)\) return;/,
  );
  // First the old session's wait for a message ends (no await before it: the
  // caller queues the message right after this function returns its promise).
  const restore = body.slice(body.indexOf("live.fellBackTo"));
  assert.ok(
    restore.indexOf("bridge.clearUserMessages(key, { keepQueued: true })") <
      restore.indexOf('restartSession(key, host, { reason: "model_restored", keepQueued: true })'),
  );
  assert.equal(/await/.test(restore.slice(0, restore.indexOf("restartSession("))), false);
  assert.match(
    body,
    /\[agent\] model restored to \$\{modelArgFor\(key\)\} after the fallback to \$\{live\.fellBackTo\}/,
  );
  // A chat message sent while a turn runs does not restart under that turn.
  assert.match(
    code,
    /onUserMessage: \(key, host\) => \{\s*const idle = !turnBusy\.has\(key\);\s*turnBusy\.add\(key\);\s*return ensureLoopForMessage\(key, host, \{ idle \}\);/,
  );
});

// What a session was given of the manuscript is remembered per pane, and
// forgotten when the session is one that did not get it: a start this turn
// did not cause, or a compaction of its history.
test("the manuscript a session has: kept for the session a turn starts, forgotten on any other start and on compaction", () => {
  assert.match(code, /const manuscriptByKey = new Map\(\);/);
  // A turn that starts or restarts the session takes nothing as seen.
  assert.match(
    code,
    /const startsSession = \(key\) => \{\s*const s = sessionFor\(key\);\s*return !s \|\| s\.settled \|\| \(!!s\.fellBackTo && !staysOn\(key\)\);\s*\};/,
  );
  assert.match(
    code,
    /manuscriptSeen: \(key\) => \(startsSession\(key\) \? null : \(manuscriptByKey\.get\(key\)\?\.hashes \?\? null\)\),/,
  );
  assert.match(
    code,
    /manuscriptGiven: \(key, hashes\) =>\s*manuscriptByKey\.set\(key, \{ hashes, fresh: startsSession\(key\) \}\),/,
  );
  const start = code.slice(code.indexOf("async function startSessionForFolder"));
  assert.match(
    start.slice(0, 600),
    /const given = manuscriptByKey\.get\(key\);\s*if \(given\?\.fresh\) given\.fresh = false;\s*else manuscriptByKey\.delete\(key\);/,
  );
  assert.match(
    code,
    /msg\.subtype === "compact_boundary"[\s\S]{0,200}manuscriptByKey\.delete\(session\.key\)/,
  );
});

// The chosen model declined two turns in a row (model-list.mjs declineStep):
// the session is not taken back to it before the next turn, the manuscript
// is not sent whole again, the pane is told with the notice and when it
// connects again, a turn the chosen model did answer clears the count, and a
// receipt names the model that ran.
test("a model that declines the manuscript twice: the session stays on the other model", () => {
  assert.match(code, /const declinedByKey = new Map\(\);/);
  assert.match(
    code,
    /const staysOn = \(key\) => staysOnFallback\(declinedByKey\.get\(key\), modelArgFor\(key\)\);/,
  );
  const at = code.indexOf('msg.subtype === "model_refusal_fallback"');
  const body = code.slice(at, at + 1200);
  assert.match(
    body,
    /declineStep\(declinedByKey\.get\(session\.key\), \{\s*asked: session\.model,\s*original: msg\.original_model,\s*fallback: msg\.fallback_model,\s*category: msg\.api_refusal_category \?\? null,\s*\}\)/,
  );
  assert.match(
    body,
    /\[agent\] model stays on \$\{d\.fallback\}: \$\{d\.original\} declined \$\{d\.count\} turns in a row \(category=\$\{d\.category \?\? "-"\}\)/,
  );
  assert.match(body, /bridge\.sendAssistantEvent\(\{ \.\.\.f\.event, stay \}, session\?\.key\);/);
  // A turn that ended on a session the CLI did not move: the count is gone.
  const result = code.slice(code.indexOf('case "result": {'));
  assert.match(
    result.slice(0, 2600),
    /if \(session\?\.key && !session\.fellBackTo\) declinedByKey\.delete\(session\.key\);/,
  );
  // Another model chosen: the record of the old one goes.
  const set = code.slice(code.indexOf("set_model: async"));
  assert.match(
    set.slice(0, 2600),
    /if \(prev !== undefined && prev !== model\) declinedByKey\.delete\(key\);/,
  );
  // A pane that connects again is told the state, without a chat notice.
  assert.match(
    set.slice(0, 2600),
    /bridge\.sendAssistantEvent\(\s*\{\s*\.\.\.refusalFallback\(\{[\s\S]{0,260}\}\)\.event,\s*stay: true,\s*quiet: true,\s*\},\s*key,\s*\);/,
  );
  assert.match(code, /declinedByKey\.delete\(key\);/);
  // Receipts of a review turn name the model the session runs on.
  assert.match(
    code,
    /getRunModel: \(key\) =>\s*sessionFor\(key\)\?\.fellBackTo \?\? sessionFor\(key\)\?\.ranModel \?\? modelArgFor\(key\),/,
  );
  assert.match(code, /if \(session\) session\.ranModel = msg\.model;/);
});
