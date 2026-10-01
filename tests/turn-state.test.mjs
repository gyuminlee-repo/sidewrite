// daemon/turn-state.mjs: which turn of a pane runs and who runs it. A pane
// has one turn at a time, run by the pane's Claude session (a chat message
// or a Claude review turn) or by the proposal runner (a review turn of
// Codex). Each ends only its own: a restart of the Claude session while a
// proposal turn runs must not tell the task pane that the review turn is
// over, because the runner goes on writing to Word.
import { test } from "node:test";
import assert from "node:assert/strict";
import { createTurnState, CLAUDE, PROPOSAL } from "../daemon/turn-state.mjs";

function make() {
  const ended = [];
  const s = createTurnState({ onReviewTurnEnded: (...a) => ended.push(a) });
  return { s, ended };
}

test("a restart of the Claude session does not end a proposal turn; the runner does", () => {
  const { s, ended } = make();
  const turn = { ref: "R1.1", selected: [], choice: { id: "codex", model: null } };
  assert.equal(s.begin("k", PROPOSAL, turn), true);
  assert.deepEqual([s.busy("k"), s.owner("k")], [true, PROPOSAL]);
  // set_model, set_context, set_cwd, a stream that ended: all end the turn
  // of the session, which this one is not.
  assert.equal(s.end("k", CLAUDE, { ok: false, error: "The session was restarted." }), false);
  assert.deepEqual(ended, [], "the task pane is told nothing");
  assert.equal(s.busy("k"), true, "the pane stays busy: the runner still writes to Word");
  assert.equal(s.end("k", PROPOSAL, { ok: true }), true);
  assert.deepEqual(ended, [["k", turn, { ok: true }]]);
  assert.equal(s.busy("k"), false);
  // Ended once: a second end says nothing.
  assert.equal(s.end("k", PROPOSAL, { ok: true }), false);
  assert.equal(ended.length, 1);
});

test("a Claude review turn is ended by its session and not by the runner", () => {
  const { s, ended } = make();
  assert.equal(s.begin("k", CLAUDE, { ref: "R1.2" }), true);
  assert.equal(s.end("k", PROPOSAL, { ok: true }), false);
  assert.equal(s.busy("k"), true);
  const restarted = { ok: false, error: "The session was restarted." };
  assert.equal(s.end("k", CLAUDE, restarted), true);
  assert.deepEqual(ended, [["k", { ref: "R1.2" }, restarted]]);
});

test("a chat turn has no review turn; a pane that is busy keeps the turn it has", () => {
  const { s, ended } = make();
  assert.equal(s.begin("k", CLAUDE), true);
  assert.equal(s.reviewTurn("k", CLAUDE), null);
  // A second begin (a message that came while a turn runs) changes nothing.
  assert.equal(s.begin("k", PROPOSAL, { ref: "R1.1" }), false);
  assert.equal(s.owner("k"), CLAUDE);
  assert.equal(s.end("k", CLAUDE), true);
  assert.deepEqual(ended, [], "no review turn ended");
  // Panes are apart.
  s.begin("a", PROPOSAL, { ref: "R1.1" });
  assert.deepEqual([s.busy("a"), s.busy("b")], [true, false]);
});

test("the review turn is shown to its owner only: the session's tools see no proposal turn", () => {
  const { s } = make();
  const turn = { ref: "R1.1", choice: { id: "codex" } };
  s.begin("k", PROPOSAL, turn);
  assert.equal(s.reviewTurn("k", PROPOSAL), turn);
  // office-tools.mjs asks for the session: no turn, so office_tracked_edit
  // of the Claude session is refused while Codex's turn runs.
  assert.equal(s.reviewTurn("k", CLAUDE), null);
  assert.equal(s.reviewTurn("other", CLAUDE), null);
});

// index.mjs starts servers when imported: its use of the turn state is read
// from its source.
test("index.mjs: the session ends Claude's turns, the runner its own, and nothing else keeps turn state", async () => {
  const { readFileSync } = await import("node:fs");
  const code = readFileSync(new URL("../daemon/index.mjs", import.meta.url), "utf8")
    .split("\n")
    .filter((l) => !l.trim().startsWith("//"))
    .join("\n");
  assert.equal(/turnBusy|reviewTurnByKey/.test(code), false, "the two maps are gone");
  assert.match(
    code,
    /function endTurnState\(key, outcome = \{ ok: true \}\) \{\s*turnState\.end\(key, CLAUDE, outcome\);\s*\}/,
  );
  assert.match(code, /endTurn: \(key, outcome\) => turnState\.end\(key, PROPOSAL, outcome\),/);
  assert.match(code, /turnState\.begin\(key, turn\.choice \? PROPOSAL : CLAUDE, turn\);/);
  assert.match(code, /getReviewTurn: \(\) => turnState\.reviewTurn\(key, CLAUDE\),/);
  // A proposal turn that could not start is ended as the runner's.
  const push = code.slice(code.indexOf("pushTurn: (key, host, text, turn) =>"));
  assert.match(push.slice(0, 1600), /turnState\.end\(key, PROPOSAL, \{ ok: false, error \}\);/);
});
