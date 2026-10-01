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
      "R1.4-1:accept",
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
