// Known answers for the pure parts of the acceptance runner (scripts/e2e).
import { test } from "node:test";
import assert from "node:assert/strict";
import { buildComPlan, judge, minimalFind, rewriteCommand } from "../scripts/e2e/judge.mjs";
import { formatRationaleComment } from "../daemon/review-tools.mjs";

const rationale = { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] };
const change = (id, extra = {}) => ({
  id,
  kind: "replace",
  find: "x",
  text: "y",
  word_comment_id: "1",
  rationale,
  ...extra,
});

test("buildComPlan: decisions per comment, first R1.3 change edited, comment text as Word holds it", () => {
  const revision = {
    comments: [
      { ref: "R1.1", changes: [change("R1.1-1")] },
      { ref: "R1.2", changes: [change("R1.2-1"), change("R1.2-2")] },
      {
        ref: "R1.3",
        changes: [change("R1.3-1", { kind: "insert" }), change("R1.3-2"), change("R1.3-3")],
      },
      { ref: "R1.4", changes: [change("R1.4-1", { kind: "delete", lost_facts: ["250 mL"] })] },
      { ref: "R1.5", changes: [change("R1.5-1", { word_comment_id: null })] },
    ],
  };
  const { ops, skipped } = buildComPlan(revision);
  assert.deepEqual(
    ops.map((o) => `${o.change_id}:${o.op}`),
    [
      "R1.1-1:accept",
      "R1.2-1:reject",
      "R1.2-2:reject",
      "R1.3-1:accept_edit",
      "R1.3-2:accept",
      "R1.3-3:accept",
      "R1.4-1:reject",
    ],
  );
  assert.deepEqual(skipped, ["R1.5-1"]);
  const r14 = ops.find((o) => o.change_id === "R1.4-1");
  assert.equal(r14.comment_text, formatRationaleComment("R1.4", rationale, { lost: ["250 mL"] }));
  assert.match(r14.comment_text, /^\[R1\.4\] w \/ y \/ a \[q1\]\n/);
});

test("minimalFind: replaces sharing an edge word are not minimal", () => {
  const yes = [
    { kind: "replace", find: "degrees", text: "°C" },
    { kind: "replace", find: "cell-free extract", text: "clarified lysate" },
    { kind: "insert", find: "30 degrees", text: "30 degrees" },
    { kind: "delete", find: "a b", text: "" },
    { kind: "replace", find: "reading", text: "measuring" },
  ];
  const no = [
    { kind: "replace", find: "30 degrees", text: "30 °C" },
    { kind: "replace", find: "the extract was", text: "the lysate was" },
    { kind: "replace", find: "cell-free extract", text: "clarified extract" },
    { kind: "replace", find: "a novel substrate", text: "a substrate" },
    { kind: "replace", find: "x y", text: "x z" },
  ];
  assert.ok(yes.every(minimalFind));
  assert.ok(no.every((c) => !minimalFind(c)));
});

test("rewriteCommand: longest prefix wins, other text kept", () => {
  const map = {
    "/plan/s2test": "/run/e2e-1",
    "cd /plan/repo ": "cd /wt ",
    "/tmp/daemon-s2.log": "/tmp/d-e2e.log",
    "/tmp/daemon-s2-knob.log": "/tmp/d-e2e-knob.log",
    "/tmp/s2-": "/run/e2e-1/ok/s2-",
  };
  assert.equal(
    rewriteCommand(
      "cd /plan/repo && jq . /plan/s2test/x.json /tmp/daemon-s2.log /tmp/daemon-s2-knob.log && test -f /tmp/s2-a.ok",
      map,
    ),
    "cd /wt && jq . /run/e2e-1/x.json /tmp/d-e2e.log /tmp/d-e2e-knob.log && test -f /run/e2e-1/ok/s2-a.ok",
  );
  // "/plan/repo" without "cd " is not a repo path to rewrite.
  assert.equal(rewriteCommand("ls /plan/repository", map), "ls /plan/repository");
});

test("judge: pane texts against the procedure wording", () => {
  const ok = judge("refresh", {
    message:
      "Accepted 0, rejected 0, edited 0, pending 4, not found 0. Changes 4, expected tool comments 4, found 4.",
  });
  assert.equal(ok.pass, true);
  const bad = judge("refresh", {
    message:
      "Accepted 0, rejected 0, edited 0, pending 4, not found 0. Changes 4, expected tool comments 4, found 3 (MISMATCH: missing R1.2-1, unrecorded -)",
  });
  assert.equal(bad.pass, false);
  assert.equal(
    judge("clean-tool", { message: "Tool comments 7 -> 0, your comments 1 -> 1 (unchanged)" }).pass,
    true,
  );
  assert.equal(
    judge("clean-tool", { message: "Tool comments 7 -> 0, your comments 1 -> 0 (MISMATCH)" }).pass,
    false,
  );
  const del = (offer, message) => judge("delete-all", { offer, message }).pass;
  assert.equal(
    del(
      "Back up and delete 2 comments (tool 0, yours 2)",
      "Backed up 2, deleted 2 of 2. Backup: /x",
    ),
    true,
  );
  assert.equal(
    del(
      "Back up and delete 3 comments (tool 1, yours 2)",
      "Backed up 3, deleted 3 of 3. Backup: /x",
    ),
    false,
  );
  const prop = (samples, done) =>
    judge(
      "propose",
      { progress_samples: samples, message: "", message_error: false },
      { ref: "R1.1", turnCompleted: done },
    ).pass;
  assert.equal(
    prop(
      [
        "R1.1 running 0:01. 5 of 5 comments not done.",
        "R1.1 running 0:02. 5 of 5 comments not done.",
      ],
      true,
    ),
    true,
  );
  assert.equal(prop(["R1.1 running 0:01. 5 of 5 comments not done."], true), false);
  assert.equal(
    prop(
      [
        "R1.11 running 0:01. 5 of 5 comments not done.",
        "R1.11 running 0:02. 5 of 5 comments not done.",
      ],
      true,
    ),
    false,
  );
  assert.equal(
    prop(
      [
        "R1.1 running 0:01. 5 of 5 comments not done.",
        "R1.1 running 0:02. 5 of 5 comments not done.",
      ],
      false,
    ),
    false,
  );
  assert.equal(judge("split", { ok: false, error: "pane not reachable" }).pass, false);
});

test("badgesMatch: the three badges the procedure names", async () => {
  const { badgesMatch } = await import("../scripts/e2e/judge.mjs");
  const v = (a, b, c) => ({
    comments: [
      { ref: "R1.1", changes: a.map((badge) => ({ badge })) },
      { ref: "R1.2", changes: b.map((badge) => ({ badge })) },
      { ref: "R1.3", changes: c.map((badge) => ({ badge })) },
    ],
  });
  assert.equal(
    badgesMatch(v(["Accepted"], ["Rejected"], ["Edited by you", "Accepted"])).pass,
    true,
  );
  assert.equal(
    badgesMatch(v(["Accepted", "Pending"], ["Rejected"], ["Edited by you"])).pass,
    false,
  );
  assert.equal(badgesMatch(v([], ["Rejected"], ["Edited by you"])).pass, false);
  assert.equal(badgesMatch(v(["Accepted"], ["Rejected"], ["Accepted"])).pass, false);
});

// Codex review (P1): a refused quit has no `exited`, and the runner used to
// Stop-Process the Word it started although a document of the user was open
// in it. Only "quit ran and the process is still there" may be killed.
test("cleanupDecision: a refused quit is never force-killed, a quit that ran and left the process is", async () => {
  const { cleanupDecision } = await import("../scripts/e2e/judge.mjs");
  const runDir = "D:\\t\\e2e-1";
  const own = `${runDir}\\review_manuscript.docx`;
  const base = { pid: 4242, preexisting: [100, 200], runDir, alive: true };
  // Case 1: quit refused, a foreign document is open in the started Word.
  const refused = cleanupDecision({
    ...base,
    quit: {
      ok: false,
      refused: "foreign_documents",
      error: "documents outside the test folder are open",
      documents: [own, "C:\\Users\\me\\thesis.docx"],
      outside: ["C:\\Users\\me\\thesis.docx"],
    },
  });
  assert.equal(refused.action, "leave");
  assert.equal(refused.fail, true);
  assert.match(refused.reason, /thesis\.docx/);
  // Case 2: quit ran on documents of the run folder only, process still alive.
  const stuck = cleanupDecision({
    ...base,
    quit: { ok: true, documents: [own], pid: 4242, exited: false },
  });
  assert.deepEqual([stuck.action, stuck.fail], ["kill", false]);
  // Quit ran and the process went away: nothing to do.
  assert.equal(
    cleanupDecision({ ...base, quit: { ok: true, documents: [own], exited: true } }).action,
    "none",
  );
  // A quit that "ran" with a document outside the folder in its list is not
  // trusted (same prefix, other folder).
  assert.equal(
    cleanupDecision({
      ...base,
      quit: { ok: true, documents: [own, "D:\\t\\e2e-10\\x.docx"], exited: false },
    }).action,
    "leave",
  );
  // COM call failed: the open documents are unknown, so no kill while alive.
  const unknown = cleanupDecision({ ...base, quit: { ok: false, error: "bind failed" } });
  assert.deepEqual([unknown.action, unknown.fail], ["leave", true]);
  assert.equal(
    cleanupDecision({ ...base, alive: false, quit: { ok: false, error: "bind failed" } }).action,
    "none",
  );
  // Never a process that was there before the run, never without a pid.
  assert.equal(
    cleanupDecision({ ...base, pid: 100, quit: { ok: true, documents: [own], exited: false } })
      .action,
    "none",
  );
  assert.equal(cleanupDecision({ ...base, pid: undefined, quit: undefined }).action, "none");
});

// Codex review (P2): a failed verify_command and features_error did not
// change the exit code of the run.
test("exitDecision: failed verify commands and features_error fail the run; a conditional feature is skipped only when its condition was provably absent", async () => {
  const { exitDecision } = await import("../scripts/e2e/judge.mjs");
  const pass = [{ step: "split", result: "PASS" }];
  const turns = "[agent] turn complete (success) slash=false output=true\n";
  const limited = turns + "[agent] turn complete (usage_limit) slash=false output=false\n";
  const f = (id, exit) => ({ id, exit, command: "x" });
  const run = (over) =>
    exitDecision({
      failed: false,
      rows: pass,
      features: [],
      conditional: [],
      logs: turns,
      ...over,
    });

  assert.equal(run({ features: [f("S2-18", 0), f("S2-19", 0)] }).code, 0);
  // A verify command that fails.
  const bad = run({ features: [f("S2-18", 0), f("S2-20", 1)] });
  assert.equal(bad.code, 1);
  assert.deepEqual(
    bad.features.map((x) => x.status),
    ["PASS", "FAIL"],
  );
  assert.match(bad.reasons.join("\n"), /S2-20/);
  // The feature list could not be read or run.
  assert.equal(run({ featuresError: "ENOENT feature_list.json" }).code, 1);
  // S2-27 without --conditional counts.
  assert.equal(run({ features: [f("S2-27", 1)] }).code, 1);
  // --conditional S2-27, no usage-limit turn in this run's log: skipped.
  const skipped = run({ features: [f("S2-27", 1)], conditional: ["S2-27"] });
  assert.equal(skipped.code, 0);
  assert.equal(skipped.features[0].status, "SKIPPED-CONDITION-NOT-MET");
  // The condition happened in this run: the failure counts.
  const met = run({ features: [f("S2-27", 1)], conditional: ["S2-27"], logs: limited });
  assert.deepEqual([met.code, met.features[0].status], [1, "FAIL"]);
  // Absence is not proven by a log that shows no finished turn at all.
  for (const logs of ["", null, "[bridge] hello received\n"]) {
    assert.equal(run({ features: [f("S2-27", 1)], conditional: ["S2-27"], logs }).code, 1);
  }
  // A feature with no known condition cannot be declared conditional.
  const unknown = run({ features: [f("S2-30", 1)], conditional: ["S2-30"] });
  assert.deepEqual([unknown.code, unknown.features[0].status], [1, "FAIL"]);
  // A conditional feature that passes is a pass.
  assert.equal(
    run({ features: [f("S2-27", 0)], conditional: ["S2-27"], logs: limited }).features[0].status,
    "PASS",
  );
  // Steps and a thrown run still decide as before.
  assert.equal(run({ rows: [{ step: "split", result: "FAIL" }] }).code, 1);
  assert.equal(run({ failed: true }).code, 1);
});

// C1: after accept/reject and Refresh the rationale comment of the change is
// resolved. Read from Word over COM (Comment.Done), not from the pane.
test("resolvedMatch: comments of accepted and rejected changes are Done in Word, edited ones need not be", async () => {
  const { resolvedMatch } = await import("../scripts/e2e/judge.mjs");
  const text = (ref, id) =>
    formatRationaleComment(ref, { ...rationale, what: `w ${id}` }, { lost: [] });
  const ch = (ref, n, state, extra = {}) =>
    change(`${ref}-${n}`, { state, rationale: { ...rationale, what: `w ${ref}-${n}` }, ...extra });
  const revision = {
    comments: [
      { ref: "R1.1", changes: [ch("R1.1", 1, "accepted")] },
      { ref: "R1.2", changes: [ch("R1.2", 1, "rejected")] },
      { ref: "R1.3", changes: [ch("R1.3", 1, "user_edited"), ch("R1.3", 2, "pending")] },
    ],
  };
  const word = (done) => ({
    comments: [
      { text: text("R1.1", "R1.1-1"), done: done[0] },
      { text: text("R1.2", "R1.2-1"), done: done[1] },
      { text: text("R1.3", "R1.3-1"), done: false },
      { text: "Check the lag time against the raw data.", done: false },
    ],
  });
  const msg = "Accepted 1, rejected 1, edited 1, pending 1, not found 0. Changes 4, ...";
  const ok = resolvedMatch(revision, word([true, true]), msg);
  assert.equal(ok.pass, true);
  assert.match(ok.got, /R1\.1-1 done; R1\.2-1 done/);
  // One still open in Word.
  const open = resolvedMatch(revision, word([true, false]), msg);
  assert.equal(open.pass, false);
  assert.match(open.got, /R1\.2-1 open/);
  // Done could not be read (null), or the comment is gone.
  assert.equal(resolvedMatch(revision, word([true, null]), msg).pass, false);
  const gone = word([true, true]);
  gone.comments.splice(1, 1);
  assert.match(resolvedMatch(revision, gone, msg).got, /R1\.2-1 no comment/);
  // The pane said Word refused.
  assert.equal(
    resolvedMatch(
      revision,
      word([true, true]),
      msg + " 2 rationale comment(s) not marked resolved (Word: GeneralException; ...).",
    ).pass,
    false,
  );
  // Nothing accepted or rejected: nothing was shown.
  assert.equal(
    resolvedMatch({ comments: [revision.comments[2]] }, word([true, true]), msg).pass,
    false,
  );
});

test("judge refresh-mode14: counts, no mismatch, and the not-resolved note for every accepted change", () => {
  const head = (a) =>
    `Accepted ${a}, rejected 0, edited 0, pending 0, not found 0. Changes 1, expected tool comments 1, found 1.`;
  const note = (n) =>
    ` ${n} rationale comment(s) not marked resolved (Word: GeneralException; one cause is a document in compatibility mode).`;
  const j = (message) => judge("refresh-mode14", { message }).pass;
  assert.equal(j(head(1) + note(1)), true);
  assert.equal(j(head(1)), false); // no note: the resolve was not refused
  assert.equal(j(head(1) + note(2)), false); // not one note per handled change
  assert.equal(j(head(0) + note(1)), false); // nothing accepted
  assert.equal(j("Refresh failed: GeneralException"), false); // counts lost
  assert.equal(
    j(
      "Accepted 1, rejected 0, edited 0, pending 0, not found 0. Changes 1, expected tool comments 1, found 0 (MISMATCH: missing R1.1-1, unrecorded -)" +
        note(1),
    ),
    false,
  );
});

// Codex review 16: a change next to another tracked change in its paragraph.
// The author rejects the R1.4 deletion while a tracked insertion of their own
// is pending in the same paragraph; the change must read Rejected (not
// Pending) and its rationale comment must be resolved.
test("siblingMatch: the handled change shows its own state while a sibling revision is pending", async () => {
  const { siblingMatch, DEFAULT_DECISIONS } = await import("../scripts/e2e/judge.mjs");
  assert.equal(DEFAULT_DECISIONS["R1.4"], "reject");
  const HAND = "Checked: ";
  const ch = change("R1.4-1", { kind: "delete", state: "rejected", text: "" });
  const revision = { comments: [{ ref: "R1.4", changes: [ch] }] };
  const text = formatRationaleComment("R1.4", ch.rationale, { lost: [] });
  const view = (badge) => ({ comments: [{ ref: "R1.4", changes: [{ badge }] }] });
  const word = (done, revisions) => ({ comments: [{ text, done }], revisions });
  const pending = [{ type: 1, text: HAND }];
  const ok = siblingMatch(revision, view("Rejected"), word(true, pending), {
    ref: "R1.4",
    hand: HAND,
  });
  assert.equal(ok.pass, true, ok.got);
  // The bug: the pane says Pending and the comment stays open.
  const bug = siblingMatch(
    { comments: [{ ref: "R1.4", changes: [{ ...ch, state: "pending" }] }] },
    view("Pending"),
    word(false, pending),
    { ref: "R1.4", hand: HAND },
  );
  assert.equal(bug.pass, false);
  assert.match(bug.got, /badge Pending/);
  // Rejected in the file but the comment is still open.
  assert.equal(
    siblingMatch(revision, view("Rejected"), word(false, pending), { ref: "R1.4", hand: HAND })
      .pass,
    false,
  );
  // No sibling revision left in Word: the case was not shown.
  assert.equal(
    siblingMatch(revision, view("Rejected"), word(true, []), { ref: "R1.4", hand: HAND }).pass,
    false,
  );
  // A replace is judged by its text anyway: the case was not shown either.
  const replace = { comments: [{ ref: "R1.4", changes: [{ ...ch, kind: "replace" }] }] };
  assert.equal(
    siblingMatch(replace, view("Rejected"), word(true, pending), { ref: "R1.4", hand: HAND }).pass,
    false,
  );
});
