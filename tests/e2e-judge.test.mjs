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

// ---- Stage prefix, markers, provider and the new pane steps (stage 2.5) ----

test("featureIds and stageOf take any stage prefix", async () => {
  const { featureIds, stageOf } = await import("../scripts/e2e/judge.mjs");
  assert.deepEqual(featureIds("S25-30..S25-32"), ["S25-30", "S25-31", "S25-32"]);
  assert.deepEqual(featureIds("S2-18"), ["S2-18"]);
  assert.deepEqual(featureIds("S2-08..S2-10"), ["S2-08", "S2-09", "S2-10"]);
  assert.equal(stageOf("S25-30..S25-41"), "s25");
  assert.equal(stageOf("S2-18..S2-34"), "s2");
  assert.throws(() => featureIds("S2-18..S25-20"), /one stage/);
  assert.throws(() => featureIds("18..20"), /S<stage>-<n>/);
  assert.throws(() => featureIds("S25-41..S25-30"), /ends before it starts/);
});

test("rewriteVerify: markers of the stage go to the run's ok folder, plan paths to this run", async () => {
  const { rewriteVerify } = await import("../scripts/e2e/judge.mjs");
  const o = {
    runDir: "/r",
    prefix: "s25",
    planDir: "/tmp/s25-run",
    planRepo: "/plan/repo",
    repo: "/here",
    planLog: "/tmp/daemon-s25.log",
    log: "/logs/run.log",
  };
  assert.match(
    rewriteVerify("test -f /tmp/s25-x.ok", { runDir: "/r", prefix: "s25" }),
    /\/r\/ok\/s25-x\.ok/,
  );
  // The commands of the stage 2.5 feature list, as written there.
  assert.equal(
    rewriteVerify(
      'test -f /tmp/s25-ollama-split.ok && grep -q \'"model":"ollama:\' /tmp/s25-run/paper.decisions.jsonl',
      o,
    ),
    'test -f /r/ok/s25-ollama-split.ok && grep -q \'"model":"ollama:\' /r/paper.decisions.jsonl',
  );
  assert.equal(
    rewriteVerify(
      "grep -q '\\[models\\] source=sdk' /tmp/daemon-s25.log && test -f /tmp/s25-models.ok",
      o,
    ),
    "grep -q '\\[models\\] source=sdk' /logs/run.log && test -f /r/ok/s25-models.ok",
  );
  assert.equal(
    rewriteVerify("jq -e '.overlaps == 0' /tmp/s25-propose-all.evidence.json", o),
    "jq -e '.overlaps == 0' /r/ok/s25-propose-all.evidence.json",
  );
  assert.equal(
    rewriteVerify("cd /plan/repo && node --test tests/x.test.mjs", o),
    "cd /here && node --test tests/x.test.mjs",
  );
  // The knob log of the plan log is this run's knob log.
  assert.equal(
    rewriteVerify("grep -q x /tmp/daemon-s25-knob.log", { ...o, knobLog: "/logs/run-knob.log" }),
    "grep -q x /logs/run-knob.log",
  );
  // Stage 2 as before.
  assert.equal(
    rewriteVerify("test -f /tmp/s2-list-matches.ok", { runDir: "/r", prefix: "s2" }),
    "test -f /r/ok/s2-list-matches.ok",
  );
  assert.throws(
    () => rewriteVerify("test -f /tmp/s2-x.ok", { runDir: "/r", prefix: "S25" }),
    /prefix/,
  );
  assert.throws(() => rewriteVerify("true", { prefix: "s25" }), /runDir/);
});

test("a stale /tmp marker cannot satisfy a rewritten command", async (t) => {
  if (process.platform === "win32") return t.skip("verify commands are bash lines");
  const { rewriteVerify } = await import("../scripts/e2e/judge.mjs");
  const fs = await import("node:fs");
  const os = await import("node:os");
  const path = await import("node:path");
  const { spawnSync } = await import("node:child_process");
  const runDir = fs.mkdtempSync(path.join(os.tmpdir(), "e2e-judge-"));
  const name = `s25-unit-${process.pid}`;
  const stale = [`/tmp/${name}.ok`, `/tmp/${name}.evidence.json`];
  fs.writeFileSync(stale[0], "");
  fs.writeFileSync(stale[1], '{"document_changes":0}');
  t.after(() => {
    for (const f of stale) fs.rmSync(f, { force: true });
  });
  const exit = (cmd) => spawnSync("bash", ["-c", cmd], { stdio: "ignore" }).status;
  const written = `test -f /tmp/${name}.ok && test -s /tmp/${name}.evidence.json`;
  assert.equal(exit(written), 0, "as written, the command is satisfied by the stale files");
  const cmd = rewriteVerify(written, { runDir, prefix: "s25" });
  assert.equal(cmd.includes("/tmp/s25-"), false, cmd);
  assert.notEqual(exit(cmd), 0, "rewritten, it does not read them");
  // Only this run's own files pass it.
  fs.mkdirSync(path.join(runDir, "ok"));
  fs.writeFileSync(path.join(runDir, "ok", `${name}.ok`), "");
  fs.writeFileSync(path.join(runDir, "ok", `${name}.evidence.json`), "{}");
  assert.equal(exit(cmd), 0);
});

test("staleMarkers: only the stage's marker and evidence files of the ok folder", async () => {
  const { staleMarkers } = await import("../scripts/e2e/judge.mjs");
  assert.deepEqual(
    staleMarkers(
      [
        "s25-models.ok",
        "s25-providers.evidence.json",
        "s2-list-matches.ok",
        "s25-run",
        "s25-notes.txt",
        "xs25-a.ok",
        "s250-a.ok",
        "result.json",
      ],
      "s25",
    ),
    ["s25-models.ok", "s25-providers.evidence.json"],
  );
  assert.deepEqual(staleMarkers(["s2-a.ok", "s25-a.ok"], "s2"), ["s2-a.ok"]);
});

test("providersEvidence: one entry per provider from the judgment menu, with the reason", async () => {
  const { providersEvidence } = await import("../scripts/e2e/judge.mjs");
  const step = {
    ok: true,
    menus: {
      turn: [
        { id: "claude", text: "Claude (Agent SDK)", enabled: true, note: "" },
        {
          id: "codex",
          text: "Codex CLI (judgments only)",
          enabled: false,
          note: "Judgments only. x",
        },
      ],
      oneshot: [
        { id: "claude", text: "Claude (Agent SDK)", enabled: true, note: "" },
        { id: "codex", text: "Codex CLI", enabled: true, note: "" },
        { id: "openrouter", text: "OpenRouter (No API key)", enabled: false, note: "No API key" },
      ],
    },
    rows: [
      { id: "claude", state: "ok" },
      { id: "codex", state: "ok" },
      { id: "openrouter", state: "off" },
    ],
  };
  const e = providersEvidence(step);
  assert.equal(e.pass, true);
  assert.deepEqual(e.codex, {
    enabled: true,
    note: "",
    turn_enabled: false,
    turn_note: "Judgments only. x",
    state: "ok",
  });
  assert.deepEqual(e.openrouter, {
    enabled: false,
    note: "No API key",
    turn_enabled: false,
    turn_note: "",
    state: "off",
  });
  assert.equal(e.claude.turn_enabled, true);
  // A list that did not load, or Claude not selectable, is no pass.
  assert.equal(providersEvidence({ ok: false, error: "x" }).pass, false);
  const noClaude = { turn: [], oneshot: [{ id: "claude", enabled: false, note: "n" }] };
  assert.equal(providersEvidence({ ...step, menus: noClaude }).pass, false);
});

test("turnOverlaps and proposeAllEvidence: review turns of the log, one at a time", async () => {
  const { turnOverlaps, proposeAllEvidence } = await import("../scripts/e2e/judge.mjs");
  const log = [
    "[review] prompt R1.3 instructions=L1 forbid=L2 omitted_forbid=- paragraphs=0 (preview)",
    "[review] prompt R1.1 instructions=L1 forbid=L2 omitted_forbid=- paragraphs=0",
    "[review] turn done R1.1",
    "[review] prompt R1.2 instructions=L1 forbid=L2 omitted_forbid=- paragraphs=0",
    "[review] turn done R1.2 failed: x",
  ].join("\n");
  assert.deepEqual(turnOverlaps(log), { started: ["R1.1", "R1.2"], overlaps: 0 });
  const bad = [
    "[review] prompt R1.1 instructions=- forbid=- omitted_forbid=- paragraphs=0",
    "[review] prompt R1.2 instructions=- forbid=- omitted_forbid=- paragraphs=0",
    "[review] turn done R1.1",
    "[review] turn done R1.2",
  ].join("\n");
  assert.equal(turnOverlaps(bad).overlaps, 1);

  const ok = proposeAllEvidence({ message: "Propose all: 2 of 2 proposed.", log });
  assert.deepEqual(
    { pass: ok.pass, overlaps: ok.overlaps, proposed: ok.proposed, open_before: ok.open_before },
    { pass: true, overlaps: 0, proposed: 2, open_before: 2 },
  );
  assert.deepEqual(ok.started, ["R1.1", "R1.2"]);
  const stopped = proposeAllEvidence({
    message: "Propose all stopped at R1.2: The turn was stopped. 1 of 2 proposed.",
    log,
  });
  assert.equal(stopped.pass, false);
  assert.equal(stopped.proposed, 1);
  const none = proposeAllEvidence({ message: "No open comment to propose changes for.", log });
  assert.equal(none.pass, false);
  assert.equal(
    proposeAllEvidence({ message: "Propose all: 2 of 2 proposed.", log: bad }).pass,
    false,
  );
});

test("judge: the stage 2.5 pane steps", () => {
  const sel = (r, ctx) => judge("select-provider", r, ctx);
  const chosen = { ok: true, selected: { id: "ollama", model: "qwen3.5:9b" }, message: "" };
  assert.equal(sel(chosen, { id: "ollama" }).pass, true);
  const refused = sel(
    {
      ok: true,
      selected: { id: "claude", model: null },
      message: "",
      refused: "Codex CLI (judgments only)",
    },
    { id: "codex" },
  );
  assert.equal(refused.pass, false);
  assert.match(refused.got, /Codex CLI \(judgments only\)/);

  const loaded = {
    ok: true,
    message: "Loaded fake_review.txt. Split into comments next.",
    chars: 9,
  };
  assert.equal(judge("load-review-file", loaded).pass, true);
  const wrongType = { ok: true, message: "x.pdf: choose a .txt or .md file.", chars: 0 };
  assert.equal(judge("load-review-file", wrongType).pass, false);

  const exported = { ok: true, message: "Exported 3 ledger rule(s) to D:\\r\\paper.ledger.json" };
  assert.equal(judge("export-ledger", exported, { count: 3 }).pass, true);
  assert.equal(judge("export-ledger", exported, { count: 2 }).pass, false);

  const added = { ok: true, message: "Added 3 ledger rule(s), skipped 0 already in the ledger." };
  assert.equal(judge("import-ledger", added, { added: 3, skipped: 0 }).pass, true);
  assert.equal(judge("import-ledger", added, { added: 0, skipped: 3 }).pass, false);

  const suggested = {
    ok: true,
    candidates: [{ label: 'forbid "novel"', checked: true }],
    message: "Added 1 ledger rule(s), skipped 0 already in the ledger.",
  };
  assert.equal(judge("suggest-ledger", suggested).pass, true);
  assert.equal(
    judge("suggest-ledger", { ok: true, candidates: [], message: "0 candidate rule(s)" }).pass,
    false,
  );

  const auto = { ok: true, changed: true, after: "Auto-refreshed 10:00:00: Changes 1" };
  assert.equal(judge("focus-refresh", auto).pass, true);
  assert.equal(judge("focus-refresh", { ok: true, changed: false, after: "" }).pass, false);
});

test("exitDecision: a CLI-turn feature is skipped only when no CLI review turn started", async () => {
  const { exitDecision } = await import("../scripts/e2e/judge.mjs");
  const features = [{ id: "S25-33", exit: 1 }];
  const claudeOnly = "[agent] turn complete (success) slash=false output=true\n";
  const d = exitDecision({ rows: [], features, conditional: ["S25-33"], logs: claudeOnly });
  assert.equal(d.features[0].status, "SKIPPED-CONDITION-NOT-MET");
  assert.equal(d.code, 0);
  const ran = claudeOnly + "[cli-turn] start provider=codex model=default ref=R1.1\n";
  const f = exitDecision({ rows: [], features, conditional: ["S25-33"], logs: ran });
  assert.equal(f.features[0].status, "FAIL");
  assert.equal(f.code, 1);
  // Not named with --conditional: it counts.
  assert.equal(exitDecision({ rows: [], features, logs: claudeOnly }).code, 1);
});

// ---- Stage 2.5 flow: evidence written next to each marker (run-acceptance.mjs) ----
// The pane outputs below are written by hand in the form pane-driver.cjs
// returns; none was recorded from Word yet.

test("defaultFeatures and docStem follow the stage", async () => {
  const { defaultFeatures, docStem } = await import("../scripts/e2e/judge.mjs");
  assert.equal(defaultFeatures("s2"), "S2-18..S2-34");
  assert.equal(defaultFeatures("s25"), "S25-30..S25-41");
  assert.throws(() => defaultFeatures("s9"), /--features/);
  // The stage 2.5 feature list names paper.decisions.jsonl and paper.ledger.json.
  assert.equal(docStem("s2"), "review_manuscript");
  assert.equal(docStem("s25"), "paper");
});

test("modelsEvidence: a model outside the three aliases and the list from the account", async () => {
  const { modelsEvidence } = await import("../scripts/e2e/judge.mjs");
  const step = {
    ok: true,
    title: "Models your Claude account can use",
    options: [
      { value: "sonnet", label: "Sonnet" },
      { value: "opus", label: "Opus" },
      { value: "haiku", label: "Haiku" },
      { value: "claude-fable-5", label: "Fable · claude-fable-5" },
    ],
  };
  const log = "[daemon] Ready\n[models] source=sdk count=11\n";
  const e = modelsEvidence({ step, log });
  assert.deepEqual([e.pass, e.outside_aliases, e.source_sdk], [true, ["claude-fable-5"], true]);
  assert.equal(
    modelsEvidence({ step, log: "[models] source=fallback count=3 error=x\n" }).pass,
    false,
  );
  const aliases = { ...step, options: step.options.slice(0, 3) };
  assert.equal(modelsEvidence({ step: aliases, log }).pass, false);
  assert.equal(modelsEvidence({ step: { ok: false, error: "x" }, log }).pass, false);
  assert.equal(judge("models", step).pass, true);
  assert.equal(
    judge("models", { ...step, title: "Model list unavailable (x); aliases only" }).pass,
    false,
  );
});

test("ollamaSplitEvidence: Ollama chosen, the split done, a receipt with an ollama model", async () => {
  const { ollamaSplitEvidence } = await import("../scripts/e2e/judge.mjs");
  const select = { ok: true, selected: { id: "ollama", model: "qwen3.5:9b" }, message: "" };
  const split = {
    ok: true,
    message: "Split by regex: 7 comment(s). Edit, then Save.",
    draft_refs: [],
  };
  const decisions =
    '{"kind":"requirement_extract","model":"ollama:qwen3.5:9b"}\n{"kind":"ledger_check","model":"haiku"}\n';
  const e = ollamaSplitEvidence({ select, split, decisions });
  assert.deepEqual([e.pass, e.models], [true, ["ollama:qwen3.5:9b", "haiku"]]);
  assert.equal(
    ollamaSplitEvidence({ select, split, decisions: '{"model":"haiku"}\n' }).pass,
    false,
  );
  assert.equal(ollamaSplitEvidence({ select, split, decisions: "" }).pass, false);
  const refused = {
    ok: true,
    selected: { id: "claude", model: null },
    refused: "Ollama (local) (Ollama is not running)",
  };
  const r = ollamaSplitEvidence({ select: refused, split, decisions });
  assert.equal(r.pass, false);
  assert.match(r.reason, /Ollama is not running/);
});

test("cliTurnEvidence: a CLI turn that started, used the endpoint and left a change with its comment", async () => {
  const { cliTurnEvidence } = await import("../scripts/e2e/judge.mjs");
  const select = { ok: true, selected: { id: "codex", model: "gpt-5.5" }, message: "" };
  const log = [
    "[cli-turn] start provider=codex model=gpt-5.5 ref=R1.1",
    "[mcp-http] POST route=ab key=k ref=R1.1",
    "[cli-turn] done provider=codex code=0 tools=2",
  ].join("\n");
  const revision = {
    comments: [{ ref: "R1.1", changes: [{ id: "R1.1-1", word_comment_id: "12" }] }],
  };
  const e = cliTurnEvidence({ provider: "codex", ref: "R1.1", select, log, revision });
  assert.deepEqual(
    [e.pass, e.started, e.done, e.endpoint_calls, e.changes],
    [true, true, true, 1, 1],
  );
  // No rationale comment, no change, or no finished turn: no pass.
  const bare = { comments: [{ ref: "R1.1", changes: [{ id: "R1.1-1", word_comment_id: null }] }] };
  assert.equal(
    cliTurnEvidence({ provider: "codex", ref: "R1.1", select, log, revision: bare }).pass,
    false,
  );
  assert.equal(
    cliTurnEvidence({
      provider: "codex",
      ref: "R1.1",
      select,
      log: log.replace("code=0", "code=1"),
      revision,
    }).pass,
    false,
  );
  assert.equal(
    cliTurnEvidence({ provider: "antigravity", ref: "R1.1", select, log, revision }).pass,
    false,
    "the log lines of another provider do not count",
  );
  // The provider is not offered for review turns: nothing ran, and the reason is kept.
  const refused = {
    ok: true,
    selected: { id: "claude", model: null },
    refused: "Codex CLI (judgments only): Judgments only. Codex CLI will not call the office tools",
  };
  const n = cliTurnEvidence({ provider: "codex", ref: "R1.1", select: refused, log: "", revision });
  assert.deepEqual([n.pass, n.started], [false, false]);
  assert.match(n.reason, /judgments only/);
});

test("guardEvidence: the forbid rule holds on a CLI turn; the count the feature reads is at the top", async () => {
  const { guardEvidence } = await import("../scripts/e2e/judge.mjs");
  const select = { ok: true, selected: { id: "codex", model: "gpt-5.5" }, message: "" };
  const log = [
    "[cli-turn] start provider=codex model=gpt-5.5 ref=R1.3",
    "[review] blocked R1.3 by=L2",
    "[cli-turn] done provider=codex code=0 tools=3",
  ].join("\n");
  const clean = {
    comments: [{ ref: "R1.3", changes: [{ id: "R1.3-1", text: "followed by optical density" }] }],
  };
  const e = guardEvidence({
    provider: "codex",
    ref: "R1.3",
    forbidden: "\\bnovel",
    select,
    log,
    revision: clean,
  });
  assert.deepEqual([e.pass, e.document_changes, e.blocked], [true, 0, true]);
  const leaked = {
    comments: [{ ref: "R1.3", changes: [{ id: "R1.3-1", text: "a novel substrate" }] }],
  };
  const bad = guardEvidence({
    provider: "codex",
    ref: "R1.3",
    forbidden: "\\bnovel",
    select,
    log,
    revision: leaked,
  });
  assert.deepEqual([bad.pass, bad.document_changes], [false, 1]);
  const refused = { ok: true, selected: { id: "claude" }, refused: "Codex CLI (judgments only)" };
  const n = guardEvidence({
    provider: "codex",
    ref: "R1.3",
    forbidden: "\\bnovel",
    select: refused,
    log: "",
    revision: clean,
  });
  assert.deepEqual([n.pass, n.document_changes], [false, null]);
});

test("stopEvidence: Stop ended the CLI turn, the pane is ready and no CLI process is left", async () => {
  const { stopEvidence } = await import("../scripts/e2e/judge.mjs");
  const select = { ok: true, selected: { id: "codex", model: "gpt-5.5" }, message: "" };
  const step = { ok: true, stopped: true, agent_status: "Ready", seconds: 12 };
  const log =
    "[cli-turn] start provider=codex model=gpt-5.5 ref=R1.4\n[cli-turn] done provider=codex code=- tools=0\n";
  assert.equal(stopEvidence({ select, step, log, processes: 0 }).pass, true);
  assert.equal(stopEvidence({ select, step, log, processes: 1 }).pass, false);
  assert.equal(
    stopEvidence({ select, step: { ...step, stopped: false }, log, processes: 0 }).pass,
    false,
  );
  assert.equal(
    stopEvidence({ select, step: { ...step, agent_status: "Working" }, log, processes: 0 }).pass,
    false,
  );
  assert.equal(judge("propose-stop", step).pass, true);
  assert.equal(
    judge("propose-stop", { ok: true, stopped: false, reason: "the turn ended before Stop" }).pass,
    false,
  );
});

test("focusEvidence, suggestEvidence and cleanupEvidence", async () => {
  const { focusEvidence, suggestEvidence, cleanupEvidence } =
    await import("../scripts/e2e/judge.mjs");
  const focus = {
    ok: true,
    changed: true,
    after:
      "Auto-refreshed 10:00:00: Changes 5, expected tool comments 5, found 5. Outside changes: 1.",
    external: ["X1: a sentence was deleted. Lost: 0.82; 0.91"],
  };
  assert.equal(focusEvidence(focus).pass, true);
  assert.equal(focusEvidence({ ...focus, external: [] }).pass, false);
  assert.equal(focusEvidence({ ...focus, changed: false }).pass, false);

  const suggest = {
    ok: true,
    candidates: [
      { label: "Forbidden, must not match /cell-free extract/i: Please do not ...", checked: true },
    ],
    message: "Added 1 ledger rule(s), skipped 0 already in the ledger.",
    ledger: [
      "L1 instruction a",
      "L2 forbid b",
      "L3 instruction c",
      'L4 forbid Please do not use "cell-free extract"',
    ],
  };
  assert.equal(suggestEvidence({ step: suggest, pick: "cell-free extract", before: 3 }).pass, true);
  assert.equal(
    suggestEvidence({ step: suggest, pick: "cell-free extract", before: 4 }).pass,
    false,
  );
  assert.equal(
    suggestEvidence({ step: { ...suggest, candidates: [] }, pick: "x", before: 3 }).pass,
    false,
  );

  const cleanup = { own_word_gone: true, preexisting_alive: true };
  assert.equal(cleanupEvidence({ cleanup, listening: 0 }).pass, true);
  assert.equal(cleanupEvidence({ cleanup, listening: 2 }).pass, false);
  assert.equal(
    cleanupEvidence({ cleanup: { ...cleanup, own_word_gone: false }, listening: 0 }).pass,
    false,
  );
  assert.equal(
    cleanupEvidence({ cleanup: { ...cleanup, preexisting_alive: false }, listening: 0 }).pass,
    false,
  );
  assert.equal(
    cleanupEvidence({
      cleanup: { ...cleanup, left_running: [{ pid: 1, reason: "x" }] },
      listening: 0,
    }).pass,
    false,
  );
  assert.equal(
    cleanupEvidence({ cleanup: { ...cleanup, daemon_error: "x" }, listening: 0 }).pass,
    false,
  );
  assert.equal(cleanupEvidence({ cleanup: { kept: { word_pid: 1 } }, listening: 2 }).pass, false);
});

test("the stage 2.5 feature commands pass on the evidence the builders write, and fail on a bare run", async (t) => {
  if (process.platform === "win32") return t.skip("verify commands are bash lines");
  const j = await import("../scripts/e2e/judge.mjs");
  const fs = await import("node:fs");
  const os = await import("node:os");
  const path = await import("node:path");
  const { spawnSync } = await import("node:child_process");
  if (spawnSync("jq", ["--version"]).status !== 0) return t.skip("jq is not installed");
  const run = fs.mkdtempSync(path.join(os.tmpdir(), "e2e-s25-"));
  const ok = path.join(run, "ok");
  fs.mkdirSync(ok);
  const log = path.join(run, "daemon.log");
  const opts = {
    runDir: run,
    prefix: "s25",
    planDir: "/tmp/s25-run",
    planLog: "/tmp/daemon-s25.log",
    log,
  };
  const exit = (cmd) =>
    spawnSync("bash", ["-c", j.rewriteVerify(cmd, opts)], { stdio: "ignore" }).status;
  // As the runner's marker(): evidence always, .ok only on a pass.
  const marker = (name, evidence) => {
    fs.writeFileSync(path.join(ok, `${name}.evidence.json`), JSON.stringify(evidence));
    if (evidence.pass) fs.writeFileSync(path.join(ok, `${name}.ok`), "");
  };
  // The commands of the feature list that read an evidence file with jq.
  const S31 =
    "jq -e '.claude.enabled and .codex.enabled and (.openrouter.enabled|not) and (.openrouter.note|test(\"No API key\"))' /tmp/s25-providers.evidence.json";
  const S34 =
    "test -f /tmp/s25-codex-guard.ok && jq -e '.document_changes == 0' /tmp/s25-codex-guard.evidence.json";
  const S37 =
    "jq -e '.overlaps == 0 and .proposed == .open_before' /tmp/s25-propose-all.evidence.json";
  const S38 =
    "test -f /tmp/s25-focus-refresh.ok && grep -q '\\[review\\] external X' /tmp/daemon-s25.log";
  fs.writeFileSync(log, "");
  for (const c of [S31, S34, S37, S38]) assert.notEqual(exit(c), 0, c);

  const menu = (id, enabled, note = "") => ({ id, text: id, enabled, note });
  marker(
    "s25-providers",
    j.providersEvidence({
      ok: true,
      menus: {
        turn: [menu("claude", true)],
        oneshot: [
          menu("claude", true),
          menu("codex", true),
          menu("openrouter", false, "No API key"),
        ],
      },
      rows: [],
    }),
  );
  assert.equal(exit(S31), 0);
  const turnLog =
    "[review] prompt R1.1 x\n[review] turn done R1.1\n[review] prompt R1.2 x\n[review] turn done R1.2\n";
  marker(
    "s25-propose-all",
    j.proposeAllEvidence({ message: "Propose all: 2 of 2 proposed.", log: turnLog }),
  );
  assert.equal(exit(S37), 0);
  const select = { ok: true, selected: { id: "codex", model: "gpt-5.5" }, message: "" };
  marker(
    "s25-codex-guard",
    j.guardEvidence({
      provider: "codex",
      ref: "R1.3",
      forbidden: "\\bnovel",
      select,
      log: "[cli-turn] start provider=codex model=gpt-5.5 ref=R1.3\n[cli-turn] done provider=codex code=0 tools=1\n",
      revision: { comments: [{ ref: "R1.3", changes: [] }] },
    }),
  );
  assert.equal(exit(S34), 0);
  marker(
    "s25-focus-refresh",
    j.focusEvidence({
      ok: true,
      changed: true,
      after: "Auto-refreshed 1:00:00: x",
      external: ["X1: y"],
    }),
  );
  fs.writeFileSync(log, "[review] external X1 kind=delete paragraph=4 lost=2\n");
  assert.equal(exit(S38), 0);
});

test("the rules review of the stage 2.5 run yields the candidate its suggest step ticks", async () => {
  const { readFileSync } = await import("node:fs");
  const { suggestLedgerRegex } = await import("../daemon/ledger-suggest.mjs");
  const { candidateLabel } = await import("../taskpane/shared/review-view.js");
  const text = readFileSync(new URL("./fixtures/fake_review_rules.txt", import.meta.url), "utf8");
  const labels = suggestLedgerRegex(text).map(candidateLabel);
  assert.equal(labels.length, 1);
  assert.match(labels[0], /^Forbidden, must not match \/cell-free extract\/i: /);
});

// Codex review of stage 2.5: Ollama stored as the judgment provider, the
// daemon restarted mid-run, Split clicked at once. log: the daemon log from
// the restart on; decisions: the receipts written after it.
test("restartSplitEvidence: after a daemon restart the split used the stored provider and no Claude judgment was made", async () => {
  const { restartSplitEvidence } = await import("../scripts/e2e/judge.mjs");
  const split = {
    ok: true,
    message: "Split by regex: 7 comment(s). Edit, then Save.",
    draft_refs: [],
  };
  const log = [
    "[daemon] Ready",
    "[providers] restore turn=ready oneshot=ready provider=ollama",
    "[providers] judgment provider=ollama model=ollama:qwen3.5:9b",
  ].join("\n");
  const decisions = '{"kind":"requirement_extract","model":"ollama:qwen3.5:9b"}\n';
  const e = restartSplitEvidence({ provider: "ollama", split, log, decisions });
  assert.deepEqual(
    [e.pass, e.restored, e.judgments, e.receipt_models],
    [true, true, { ollama: 1 }, ["ollama:qwen3.5:9b"]],
  );
  // A judgment by Claude after the restart, in the log or in a receipt, fails it.
  const leaked = `${log}\n[providers] judgment provider=claude model=haiku`;
  assert.equal(
    restartSplitEvidence({ provider: "ollama", split, log: leaked, decisions }).pass,
    false,
  );
  assert.equal(
    restartSplitEvidence({
      provider: "ollama",
      split,
      log,
      decisions: decisions + '{"kind":"comment_split","model":"haiku"}\n',
    }).pass,
    false,
  );
  // The split must have run on the provider, and the stored choice must have been restored.
  assert.equal(
    restartSplitEvidence({ provider: "ollama", split, log: "[daemon] Ready", decisions: "" }).pass,
    false,
  );
  const refused = log.replace("oneshot=ready provider=ollama", "oneshot=refused provider=-");
  assert.equal(
    restartSplitEvidence({ provider: "ollama", split, log: refused, decisions }).pass,
    false,
  );
  const failed = {
    ok: true,
    message: "The task pane has not confirmed its provider for judgments (waited 45 s).",
    draft_refs: [],
  };
  assert.equal(
    restartSplitEvidence({ provider: "ollama", split: failed, log, decisions }).pass,
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

// Follow-up review: every tracked run the pane reads from the OOXML must be
// found in the paragraph text Word reports, special characters included.
test("trackedPlaced: every tracked run of every paragraph is found in Word's paragraph text", async () => {
  const { trackedPlaced } = await import("../scripts/e2e/judge.mjs");
  const p = (index, text, tracked, counts = {}) => ({
    index,
    text,
    w_ins: 0,
    w_del: 0,
    ...counts,
    tracked,
  });
  const run = (kind, text, start) => ({ kind, text, start, marks: 1 });
  const good = [
    p(0, "Title", []),
    p(1, "Case 0 before non\u001elinear after.", [run("ins", "non\u001elinear", 14)], { w_ins: 1 }),
    p(2, "Case 1 before left\tright after.", [run("del", "left\tright", 14)], { w_del: 1 }),
  ];
  const ok = trackedPlaced(good, 2);
  assert.equal(ok.pass, true, ok.got);
  // The hyphen was lost in the extraction: the run is not in the text.
  const lost = trackedPlaced(
    [good[0], p(1, good[1].text, [run("ins", "nonlinear", 14)], { w_ins: 1 }), good[2]],
    2,
  );
  assert.equal(lost.pass, false);
  assert.match(lost.got, /paragraph 1/);
  // A tracked mark with no run read from it.
  assert.equal(
    trackedPlaced([good[0], p(1, good[1].text, [], { w_ins: 1 }), good[2]], 2).pass,
    false,
  );
  // Fewer tracked paragraphs than the document was built with.
  assert.equal(trackedPlaced(good, 3).pass, false);
  assert.equal(trackedPlaced([], 0).pass, false);
});

// The first run against Word (2026-10-01): the split step came back on the
// line the pane shows while a click is held after a reconnect, and the
// suggest step on the late result of that split. Neither is the result of
// the step; the pane steps now wait for their own message.
test("a held click or another action's late message is not the result of the step", async () => {
  const { JUDGES, restartSplitEvidence, suggestEvidence } =
    await import("../scripts/e2e/judge.mjs");
  const refs = ["E.1", "R1.1", "R1.2", "R1.3", "R1.4", "R1.5"];
  const held = {
    ok: true,
    message: "Waiting for the daemon to confirm the stored provider choice...",
    draft_refs: refs, // the rows of the split before the restart
  };
  const log = [
    "[providers] restore turn=ready oneshot=ready provider=ollama",
    "[providers] judgment provider=ollama model=ollama:qwen3.5:9b",
  ].join("\n");
  const decisions =
    '{"kind":"requirement_extract","model":"ollama:qwen3.5:9b","output":{"requests":null}}\n';
  assert.equal(JUDGES.split(held).pass, false);
  // Read at that moment: the daemon has logged nothing yet.
  assert.equal(
    restartSplitEvidence({ provider: "ollama", split: held, log: "", decisions: "" }).pass,
    false,
  );
  // Even with the log and the receipt of a split that ended later.
  assert.equal(
    restartSplitEvidence({ provider: "ollama", split: held, log, decisions }).pass,
    false,
  );

  // After the wait: the message of the split itself. A requirement extraction
  // that failed is a warning in it and a receipt of that provider all the same.
  const done = {
    ok: true,
    message:
      "Split by regex. Text before the first reviewer (usually the editor letter) is kept as E.1. Delete it if it asks for nothing. requirement extraction failed: reply was not JSON; add them by hand",
    draft_refs: refs,
  };
  assert.equal(JUDGES.split(done).pass, true);
  const e = restartSplitEvidence({ provider: "ollama", split: done, log, decisions });
  assert.deepEqual(
    [e.pass, e.restored, e.judgments, e.receipt_models],
    [true, true, { ollama: 1 }, ["ollama:qwen3.5:9b"]],
  );

  // The suggest step that read the split's late message: no candidate yet.
  const early = {
    ok: true,
    suggest_message: done.message,
    candidates: [],
    message: "",
    ledger: [],
  };
  assert.equal(suggestEvidence({ step: early, pick: "cell-free extract", before: 3 }).pass, false);
  const waited = {
    ok: true,
    suggest_message: "1 candidate rule(s) by regex. Tick the ones to keep, then Add selected.",
    candidates: [{ label: 'Forbid "cell-free extract"', checked: true }],
    message: "Added 1 ledger rule(s), skipped 0 already in the ledger.",
    ledger: ["L1", "L2", "L3", 'L4 forbid "cell-free extract"'],
  };
  assert.equal(suggestEvidence({ step: waited, pick: "cell-free extract", before: 3 }).pass, true);
});

// The patterns the pane steps wait for are the messages review-pane.js writes.
test("the pane steps wait for the message their own action ends with", async () => {
  const fs = await import("node:fs");
  const read = (p) => fs.readFileSync(new URL(p, import.meta.url), "utf8");
  const driver = read("../scripts/e2e/pane-driver.cjs");
  const pane = read("../taskpane/shared/review-pane.js");
  const waits = [...driver.matchAll(/waitMessage\(page, [^,]+, "((?:[^"\\]|\\.)*)"\)/g)].map((m) =>
    JSON.parse(`"${m[1]}"`),
  );
  assert.deepEqual(waits, ["^Split by ", "candidate rule\\(s\\) by "]);
  assert.match("Split by regex: 5 comment(s). Edit, then Save.", new RegExp(waits[0]));
  assert.match("1 candidate rule(s) by regex. Tick the ones to keep", new RegExp(waits[1]));
  const HOLD = "Waiting for the daemon to confirm the stored provider choice...";
  for (const w of waits) assert.doesNotMatch(HOLD, new RegExp(w));
  assert.ok(pane.includes("say(`Split by ${r.method}"));
  assert.ok(pane.includes("candidate rule(s) by ${r.method}."));
  assert.ok(pane.includes(`const PROVIDER_WAIT = "${HOLD}";`));
});

// The account's model list comes from outside and can be late once (seen on
// 2026-10-01: "Model list unavailable (timed out after 20000 ms); aliases
// only"; the pane fell back to the aliases, as it should). The `models` step
// asks once more before it is judged. Only the step does: the pane's own
// fallback is unchanged.
test("models step: an unavailable list is asked for once more, after the daemon's fallback has run out", async () => {
  const { createRequire } = await import("node:module");
  const { steps, MODELS_RETRY_AFTER_MS } = createRequire(import.meta.url)(
    "../scripts/e2e/pane-driver.cjs",
  );
  // The daemon keeps a failed list for 60 s (model-list.mjs fallbackTtlMs).
  const list = (await import("node:fs")).readFileSync(
    new URL("../daemon/model-list.mjs", import.meta.url),
    "utf8",
  );
  assert.match(list, /fallbackTtlMs = 60_000/);
  assert.ok(MODELS_RETRY_AFTER_MS > 60_000 && MODELS_RETRY_AFTER_MS <= 70_000);

  const LATE = "Model list unavailable (timed out after 20000 ms); aliases only";
  const GOOD = "Models your Claude account can use";
  const menu = (title, values) => ({
    title,
    selected: "sonnet",
    options: values.map((v) => ({ value: v, label: v })),
  });
  const fakePage = (readings) => {
    const did = [];
    return {
      did,
      waitForFunction: async () => void did.push("wait"),
      waitForTimeout: async (ms) => void did.push(`sleep ${ms}`),
      reload: async () => void did.push("reload"),
      evaluate: async () => ({ message: "", comments: [] }),
      $eval: async (sel) => (sel === "#composer-model" ? readings.shift() : ""),
    };
  };

  // In time: one reading, no reload.
  const ok = fakePage([menu(GOOD, ["sonnet", "claude-fable-5"])]);
  const first = await steps.models(ok);
  assert.equal(first.title, GOOD);
  assert.equal("retried" in first, false);
  assert.equal(ok.did.includes("reload"), false);

  // Late once: wait out the fallback, reload the pane, back to the Review tab, read again.
  const late = fakePage([
    menu(LATE, ["opus", "sonnet", "haiku"]),
    menu(GOOD, ["sonnet", "claude-fable-5"]),
  ]);
  const second = await steps.models(late);
  assert.deepEqual([second.title, second.retried, second.first_title], [GOOD, true, LATE]);
  assert.deepEqual(
    late.did.filter((d) => d !== "wait"),
    [`sleep ${MODELS_RETRY_AFTER_MS}`, "reload"],
  );
  assert.ok(late.did.indexOf("reload") < late.did.lastIndexOf("wait"), "waits for the pane again");

  // Late twice: one retry, and the step says what it saw.
  const twice = fakePage([menu(LATE, ["opus"]), menu(LATE, ["opus"])]);
  const third = await steps.models(twice);
  assert.deepEqual([third.title, third.retried], [LATE, true]);
  assert.equal(twice.did.filter((d) => d === "reload").length, 1);

  // The judgments of both outcomes.
  const { JUDGES, modelsEvidence } = await import("../scripts/e2e/judge.mjs");
  const log =
    "[models] source=fallback count=3 error=timed out after 20000 ms\n[models] source=sdk count=11\n";
  assert.equal(JUDGES.models({ ok: true, ...second }).pass, true);
  const e = modelsEvidence({ step: { ok: true, ...second }, log });
  assert.deepEqual([e.pass, e.retried, e.first_title], [true, true, LATE]);
  assert.equal(JUDGES.models({ ok: true, ...third }).pass, false);
  assert.equal(
    modelsEvidence({ step: { ok: true, ...third }, log: log.split("\n")[0] }).pass,
    false,
  );
  const plain = modelsEvidence({ step: { ok: true, ...first }, log });
  assert.deepEqual([plain.pass, plain.retried], [true, false]);
});

// ---- stage 3: the response letter steps -------------------------------------------
const S3 = await import("../scripts/e2e/judge.mjs");

test("stage 3 has its feature range, keeps the stage 2 manuscript name and maps its markers", () => {
  assert.equal(S3.defaultFeatures("s3"), "S3-30..S3-40");
  assert.equal(S3.docStem("s3"), "review_manuscript");
  assert.deepEqual(S3.featureIds("S3-30..S3-32"), ["S3-30", "S3-31", "S3-32"]);
  assert.deepEqual(S3.featureIds("S3-41"), ["S3-41"]);
  assert.deepEqual(S3.markerMap("s3", "/r/ok"), { "/tmp/s3-": "/r/ok/s3-" });
  // The author's decisions of the stage 3 run: R1.4 (the deleted sentence) is
  // accepted, so the letter has a removal to cite.
  assert.equal(S3.S3_DECISIONS["R1.4"], "accept");
  assert.equal(S3.S3_DECISIONS["R1.2"], "reject");
});

test("every response step has a judge that reads the pane", () => {
  for (const s of [
    "open-response",
    "load-letter",
    "apply-proposal",
    "response-decline",
    "response-draft",
    "response-draft-all",
    "response-edit",
    "response-refresh",
    "response-locate",
    "response-check",
    "response-export",
    "response-export-text",
    "response-set-lang",
  ])
    assert.equal(typeof S3.JUDGES[s], "function", s);
  const J = S3.JUDGES;
  assert.equal(J["response-check"]({ summary: "All checks passed", checks: [] }).pass, true);
  assert.equal(
    J["response-check"]({ summary: "1 failed", checks: ["FAIL R1.1 lines: x"] }).pass,
    false,
  );
  // A warning is not "All checks passed" for the acceptance run.
  assert.equal(
    J["response-check"]({ summary: "All checks passed, 1 warning", checks: [] }).pass,
    false,
  );
  const card = (ref, more = {}) => ({ ref, text: "", declined: "", open_reqs: [], ...more });
  assert.equal(
    J["open-response"]({
      message: "Refreshed from the manuscript: 5 comments.",
      cards: [card("R1.1")],
    }).pass,
    true,
  );
  assert.equal(
    J["open-response"]({
      message: "Waiting for the daemon to confirm the stored provider choice...",
      cards: [card("R1.1")],
    }).pass,
    false,
  );
  assert.equal(
    J["response-locate"]({ message: "Line numbers updated: 3/3 quotes located." }).pass,
    true,
  );
  assert.equal(
    J["response-locate"]({ message: "Line numbers updated: 2/3 quotes located." }).pass,
    false,
  );
  assert.equal(
    J["response-locate"]({ message: "Line numbers updated: 0/0 quotes located." }).pass,
    false,
  );
  const drafted = {
    message: "Drafted 4 responses.",
    cards: [card("R1.1", { text: "Done." }), card("R1.5", { declined: "Declined: none used" })],
  };
  assert.equal(J["response-draft-all"](drafted).pass, true);
  assert.equal(
    J["response-draft-all"]({ ...drafted, cards: [card("R1.1"), card("R1.5", { declined: "x" })] })
      .pass,
    false,
  );
  assert.equal(
    J["response-draft-all"]({ ...drafted, message: "Stopped at R1.2 after 1 draft: x" }).pass,
    false,
  );
  assert.equal(J["response-draft"]({ message: "Drafted R1.1." }, { ref: "R1.1" }).pass, true);
  assert.equal(J["response-draft"]({ message: "Drafted R1.2." }, { ref: "R1.1" }).pass, false);
  assert.equal(
    J["response-draft"](
      { message: "Drafted R1.1. 1 check fail for it: x.", message_error: true },
      { ref: "R1.1" },
    ).pass,
    false,
  );
  const declined = {
    message: "Declined q2 of R1.3.",
    cards: [card("R1.5", { declined: "Declined: none" }), card("R1.3")],
  };
  assert.equal(J["response-decline"](declined, { comments: ["R1.5"] }).pass, true);
  assert.equal(
    J["response-decline"](
      { ...declined, cards: [card("R1.5"), card("R1.3")] },
      { comments: ["R1.5"] },
    ).pass,
    false,
  );
  assert.equal(
    J["response-decline"](
      {
        ...declined,
        cards: [card("R1.5", { declined: "x" }), card("R1.3", { open_reqs: ["q2 ..."] })],
      },
      { comments: ["R1.5"] },
    ).pass,
    false,
  );
  assert.equal(J["response-export"]({ output: "Exported: /r/a.response.docx" }).pass, true);
  assert.equal(
    J["response-export"]({ output: "", message: "Not exported: 1 failed.", message_error: true })
      .pass,
    false,
  );
  assert.equal(J["response-edit"]({ message: "Saved R1.1." }).pass, true);
  assert.equal(
    J["response-set-lang"]({ message: "Saved for this paper: profile generic." }).pass,
    true,
  );
  assert.equal(
    J["load-letter"]({
      message:
        "Read the letter: 2 settings differ, 5 earlier answers found. Nothing is saved until you press Apply to this paper.",
      rows: ["a"],
    }).pass,
    true,
  );
  assert.equal(
    J["apply-proposal"]({
      message: "Applied to this paper: 2 settings, 5 earlier answers kept for drafting.",
    }).pass,
    true,
  );
});

test("openEvidence and proposalEvidence", () => {
  const revision = {
    comments: [
      { ref: "R1.1", text: "Use  Celsius.\n" },
      { ref: "R1.2", text: "b" },
    ],
  };
  const view = {
    profile: "generic",
    cards: [
      { ref: "R1.1", comment: "Use  Celsius.\n" },
      { ref: "R1.2", comment: "b" },
    ],
  };
  assert.deepEqual(S3.openEvidence(view, revision), {
    pass: true,
    cards: 2,
    verbatim: true,
    profile: "generic",
  });
  const other = { ...view, cards: [view.cards[0], { ref: "R1.2", comment: "B" }] };
  assert.deepEqual(
    [S3.openEvidence(other, revision).pass, S3.openEvidence(other, revision).verbatim],
    [false, false],
  );
  assert.equal(S3.openEvidence({ ...view, profile: "mine" }, revision).pass, false);
  assert.equal(S3.openEvidence({ profile: "generic", cards: [] }, { comments: [] }).pass, false);

  const rows = [
    "colors.response_body: (none) -> C00000 (6 paragraphs)",
    "colors.response_revised: (none) -> 0070C0 (7 paragraphs)",
  ];
  const after = { colors: { response_body: "C00000", response_revised: "0070C0" } };
  const e = S3.proposalEvidence({ rows, before: {}, after });
  assert.deepEqual([e.pass, e.before, e.after], [true, {}, after]);
  // Saved before Apply, a colour without its count, or nothing saved after Apply.
  assert.equal(S3.proposalEvidence({ rows, before: after, after }).pass, false);
  assert.equal(
    S3.proposalEvidence({ rows: ["colors.response_body: (none) -> C00000"], before: {}, after })
      .pass,
    false,
  );
  assert.equal(S3.proposalEvidence({ rows, before: {}, after: {} }).pass, false);
});

test("second reader: the margin number of the cited start line sits on the line that starts the quote", () => {
  const layout =
    "  8   Flasks were shaken at 200 rpm.\n  9   The incubation temperature was held at\n 10   30 °C. Samples were drawn daily.\n";
  const cits = [
    {
      id: "R1.1#0",
      quote: "The incubation temperature was held at 30 °C.",
      start: 9,
      end: 10,
      match_count: 1,
    },
  ];
  assert.deepEqual(S3.secondReader(layout, cits), {
    pass: true,
    unique: true,
    second_reader_agrees: true,
    rows: [
      { id: "R1.1#0", start: 9, line_text: "The incubation temperature was held at", agrees: true },
    ],
  });
  assert.equal(S3.secondReader(layout, [{ ...cits[0], start: 8 }]).second_reader_agrees, false);
  assert.equal(S3.secondReader(layout, [{ ...cits[0], start: 10 }]).second_reader_agrees, false);
  assert.equal(S3.secondReader(layout, [{ ...cits[0], match_count: 2 }]).unique, false);
  assert.equal(S3.secondReader(layout, []).pass, false);
  // A quote that starts in the middle of its line, a quote inside one line,
  // a word hyphenated at the line end, and a ligature of the PDF font.
  const mid =
    " 11   drawn daily. The plateau was reached on\n 12   day four in all flasks.\n 13   It was low. Growth was conﬁrmed by micro-\n 14   scopy at 600 nm.\n";
  const agree = (quote, start) =>
    S3.secondReader(mid, [{ id: "x", quote, start, end: start, match_count: 1 }]).rows[0].agrees;
  assert.equal(agree("The plateau was reached on day four in all flasks.", 11), true);
  assert.equal(agree("It was low.", 13), true);
  assert.equal(agree("Growth was confirmed by microscopy at 600 nm.", 13), true);
  assert.equal(agree("The plateau was reached on day four in all flasks.", 12), false);
  // A line that is not in the layout at all.
  assert.equal(agree("It was low.", 99), false);
});

test("line shift: citations below the inserted line move by exactly one, those above stay", () => {
  const before = [
    { id: "a", start: 9 },
    { id: "b", start: 40 },
    { id: "c", start: 55 },
  ];
  const moved = [
    { id: "a", start: 9 },
    { id: "b", start: 41 },
    { id: "c", start: 56 },
  ];
  assert.deepEqual(S3.lineShift(before, moved, 20, true), {
    pass: true,
    moved: ["b", "c"],
    kept: ["a"],
    wrong: [],
    stale_shown: true,
  });
  assert.equal(
    S3.lineShift(before, [moved[0], { id: "b", start: 42 }, moved[2]], 20, true).pass,
    false,
  );
  assert.equal(S3.lineShift(before, before, 20, true).pass, false);
  assert.equal(S3.lineShift(before, moved, 20, false).pass, false);
  // A citation that lost its line range, or an insertion line that was not found.
  assert.deepEqual(S3.lineShift(before, moved.slice(0, 2), 20, true).wrong, ["c"]);
  assert.equal(S3.lineShift(before, moved, NaN, true).pass, false);
});

test("stale, ledger, language and open-check evidence", () => {
  const v = (stale, summary, checks = []) => ({ cards: [{ ref: "R1.1", stale }], summary, checks });
  assert.deepEqual(
    S3.staleEvidence(
      v("Quote changed in the manuscript", "2 failed", [
        "FAIL R1.1 stale: quote_changed",
        "FAIL R1.1 quote_exact: x",
      ]),
      v("", "All checks passed"),
      "R1.1",
    ),
    { pass: true, stale_shown: true, failed_before: true, passed_after: true },
  );
  assert.equal(
    S3.staleEvidence(v("", "All checks passed"), v("", "All checks passed"), "R1.1").pass,
    false,
  );
  // The mark shown but no stale row of that comment, or still failing afterwards.
  assert.equal(
    S3.staleEvidence(
      v("Quote changed in the manuscript", "1 failed", ["FAIL R1.2 stale: quote_changed"]),
      v("", "All checks passed"),
      "R1.1",
    ).pass,
    false,
  );
  assert.equal(
    S3.staleEvidence(
      v("Quote changed in the manuscript", "1 failed", ["FAIL R1.1 stale: quote_changed"]),
      v("", "1 failed"),
      "R1.1",
    ).pass,
    false,
  );

  assert.deepEqual(
    S3.ledgerEvidence(
      { checks: ["FAIL R1.1 ledger: L4: unfortunately"], export_label: "Export anyway" },
      false,
      { summary: "All checks passed" },
    ),
    { pass: true, ledger_fail: 1, exported_while_failing: false, passed_after: true },
  );
  const failing = {
    checks: ["FAIL R1.1 ledger: L4: unfortunately"],
    export_label: "Export anyway",
  };
  assert.equal(S3.ledgerEvidence(failing, true, { summary: "All checks passed" }).pass, false);
  assert.equal(S3.ledgerEvidence(failing, false, { summary: "1 failed" }).pass, false);
  assert.equal(
    S3.ledgerEvidence({ ...failing, export_label: "Export letter (.docx)" }, false, {
      summary: "All checks passed",
    }).pass,
    false,
  );
  assert.equal(
    S3.ledgerEvidence({ checks: [], export_label: "Export anyway" }, false, {
      summary: "All checks passed",
    }).pass,
    false,
  );

  assert.deepEqual(
    S3.langEvidence(
      "[response] draft R1.1 lang=German model=x words=9 citations=1",
      "5. Write the rationale in German: what = ...",
      "German",
    ),
    {
      pass: true,
      response_prompt_lang: "German",
      rationale_prompt_lang: "German",
      korean_found: false,
    },
  );
  assert.equal(
    S3.langEvidence(
      "[response] draft R1.1 lang=English model=x words=9 citations=1",
      "5. Write the rationale in German: what = ...",
      "German",
    ).pass,
    false,
  );
  assert.equal(
    S3.langEvidence(
      "[response] draft R1.1 lang=German model=x words=9 citations=1",
      "5. Write the rationale in Korean: what = ...",
      "German",
    ).korean_found,
    true,
  );
  assert.equal(S3.langEvidence(undefined, undefined, "German").pass, false);

  const opened = { ok: true, opened: true, paragraphs: 31, repaired: false };
  assert.deepEqual(S3.openCheckEvidence(opened, 31, "/r/a.response.docx", "/r/a.response-2.docx"), {
    pass: true,
    opened: true,
    repaired: false,
    paragraphs_match: true,
    first_path: "/r/a.response.docx",
    second_path: "/r/a.response-2.docx",
    paragraphs: 31,
    model_paragraphs: 31,
  });
  assert.equal(
    S3.openCheckEvidence(opened, 31, "/r/a.response.docx", "/r/a.response.docx").pass,
    false,
  );
  assert.equal(S3.openCheckEvidence(opened, 30, "/r/a.response.docx", "/r/b.docx").pass, false);
  assert.equal(S3.openCheckEvidence(opened, 31, "/r/a.response.docx", null).pass, false);
  assert.equal(
    S3.openCheckEvidence({ ok: false, error: "x" }, 31, "/r/a.response.docx", "/r/b.docx").pass,
    false,
  );
  // The cleanup evidence says in one word whether both daemon ports are free.
  const cleanup = { own_word_gone: true, preexisting_alive: true };
  assert.equal(S3.cleanupEvidence({ cleanup, listening: 0 }).ports_free, true);
  assert.equal(S3.cleanupEvidence({ cleanup, listening: 1 }).ports_free, false);
});

test("exportEvidence: the colours the letter reader finds in the written letter", () => {
  const read = {
    ok: true,
    overrides: { colors: { response_body: "C00000", response_revised: "0070C0" } },
    evidence: { colors: { C00000: { paragraphs: 4 }, "0070C0": { paragraphs: 6 } } },
  };
  assert.deepEqual(
    S3.exportEvidence({ output: "Exported: /r/a.response.docx" }, read, "/r/a.response.docx", true),
    {
      pass: true,
      body: "C00000",
      revised: "0070C0",
      path: "/r/a.response.docx",
      colours: ["0070C0", "C00000"],
    },
  );
  // Not written, held, or a third colour in the letter.
  assert.equal(
    S3.exportEvidence({ output: "Exported: x" }, read, "/r/a.response.docx", false).pass,
    false,
  );
  assert.equal(S3.exportEvidence({ output: "" }, read, "/r/a.response.docx", true).pass, false);
  const third = {
    ...read,
    evidence: { colors: { ...read.evidence.colors, "00B050": { paragraphs: 1 } } },
  };
  assert.equal(
    S3.exportEvidence({ output: "Exported: x" }, third, "/r/a.response.docx", true).pass,
    false,
  );
});

test("insertTarget: the paragraph of the lowest citation, with one above it", () => {
  const cits = [
    { id: "R1.1#0", start: 9, paragraph_index: 4 },
    { id: "R1.3#0", start: 31, paragraph_index: 9 },
    { id: "R1.3#1", start: 33, paragraph_index: 9 },
  ];
  assert.equal(S3.insertTarget(cits), 9);
  // Every citation in one paragraph: nothing can stay above an insertion.
  assert.equal(S3.insertTarget(cits.slice(1)), null);
  assert.equal(S3.insertTarget([]), null);
});

test("the stage 3 feature commands pass on the evidence the builders write, and fail on a bare run", async (t) => {
  if (process.platform === "win32") return t.skip("verify commands are bash lines");
  const fs = await import("node:fs");
  const os = await import("node:os");
  const path = await import("node:path");
  const { spawnSync } = await import("node:child_process");
  if (spawnSync("jq", ["--version"]).status !== 0) return t.skip("jq is not installed");
  const run = fs.mkdtempSync(path.join(os.tmpdir(), "e2e-s3-"));
  const ok = path.join(run, "ok");
  fs.mkdirSync(ok);
  const log = path.join(run, "daemon.log");
  fs.writeFileSync(log, "");
  const opts = {
    runDir: run,
    prefix: "s3",
    planDir: "/tmp/s3-run",
    planLog: "/tmp/daemon-s3.log",
    log,
  };
  const exit = (cmd) =>
    spawnSync("bash", ["-c", S3.rewriteVerify(cmd, opts)], { stdio: "ignore" }).status;
  const marker = (name, evidence) => {
    fs.writeFileSync(path.join(ok, `${name}.evidence.json`), JSON.stringify(evidence));
    if (evidence.pass) fs.writeFileSync(path.join(ok, `${name}.ok`), "");
  };
  // The commands of the stage 3 feature list (the parts that read evidence).
  const C = {
    "S3-30":
      "test -f /tmp/s3-open.ok && jq -e '.pass == true and .cards == 5 and .verbatim == true and .profile == \"generic\"' /tmp/s3-open.evidence.json",
    "S3-31":
      'test -f /tmp/s3-proposal.ok && jq -e \'.pass == true and .before == {} and .after.colors.response_body == "C00000" and .after.colors.response_revised == "0070C0"\' /tmp/s3-proposal.evidence.json',
    "S3-33":
      "test -f /tmp/s3-lines.ok && jq -e '.pass == true and .unique == true and .second_reader_agrees == true and (.rows|length) > 0' /tmp/s3-lines.evidence.json",
    "S3-34":
      "test -f /tmp/s3-check.ok && jq -e '.pass == true and .summary == \"All checks passed\"' /tmp/s3-check.evidence.json && grep -q '\\[response\\] check fail=0' /tmp/daemon-s3.log",
    "S3-35":
      'test -f /tmp/s3-export.ok && jq -e \'.pass == true and .body == "C00000" and .revised == "0070C0"\' /tmp/s3-export.evidence.json',
    "S3-36":
      "test -f /tmp/s3-line-shift.ok && jq -e '.pass == true and (.moved|length) > 0 and (.kept|length) > 0 and (.wrong|length) == 0 and .stale_shown == true and .second.pass == true' /tmp/s3-line-shift.evidence.json",
    "S3-37":
      "test -f /tmp/s3-stale.ok && jq -e '.pass == true and .stale_shown == true and .failed_before == true and .passed_after == true' /tmp/s3-stale.evidence.json",
    "S3-38":
      "test -f /tmp/s3-ledger.ok && jq -e '.pass == true and .ledger_fail == 1 and .exported_while_failing == false and .passed_after == true' /tmp/s3-ledger.evidence.json",
    "S3-39":
      'test -f /tmp/s3-lang.ok && jq -e \'.pass == true and .response_prompt_lang == "German" and .rationale_prompt_lang == "German" and .korean_found == false\' /tmp/s3-lang.evidence.json',
    "S3-40":
      "test -f /tmp/s3-run/review_manuscript.response.R1.txt && test -f /tmp/s3-open-check.ok && jq -e '.pass == true and .opened == true and .repaired == false and .paragraphs_match == true and .second_path != .first_path' /tmp/s3-open-check.evidence.json",
    "S3-41":
      "test -f /tmp/s3-cleanup.ok && jq -e '.pass == true and .own_word_gone == true and .preexisting_alive == true and .ports_free == true' /tmp/s3-cleanup.evidence.json",
  };
  for (const [id, c] of Object.entries(C)) assert.notEqual(exit(c), 0, `${id} on a bare run`);

  const comments = ["a", "b", "c", "d", "e"].map((text, i) => ({ ref: `R1.${i + 1}`, text }));
  marker(
    "s3-open",
    S3.openEvidence(
      { profile: "generic", cards: comments.map((c) => ({ ref: c.ref, comment: c.text })) },
      { comments },
    ),
  );
  marker(
    "s3-proposal",
    S3.proposalEvidence({
      rows: [
        "colors.response_body: (none) -> C00000 (6 paragraphs)",
        "colors.response_revised: (none) -> 0070C0 (7 paragraphs)",
      ],
      before: {},
      after: { colors: { response_body: "C00000", response_revised: "0070C0" } },
    }),
  );
  const layout = "  9   The incubation temperature was held at\n 10   30 °C. Samples were drawn.\n";
  const cits = [
    {
      id: "R1.1#0",
      quote: "The incubation temperature was held at 30 °C.",
      start: 9,
      end: 10,
      match_count: 1,
    },
  ];
  const second = S3.secondReader(layout, cits);
  marker("s3-lines", second);
  const checked = { summary: "All checks passed", checks: [] };
  marker("s3-check", { pass: S3.JUDGES["response-check"](checked).pass, ...checked });
  fs.writeFileSync(log, "[response] check fail=0 warn=0\n");
  marker(
    "s3-export",
    S3.exportEvidence(
      { output: "Exported: /r/a.response.docx" },
      {
        overrides: { colors: { response_body: "C00000", response_revised: "0070C0" } },
        evidence: { colors: { C00000: {}, "0070C0": {} } },
      },
      "/r/a.response.docx",
      true,
    ),
  );
  const shift = S3.lineShift(
    [
      { id: "a", start: 9 },
      { id: "b", start: 40 },
    ],
    [
      { id: "a", start: 9 },
      { id: "b", start: 41 },
    ],
    20,
    true,
  );
  marker("s3-line-shift", { ...shift, inserted_at: 20, second, pass: shift.pass && second.pass });
  const v = (stale, summary, checks = []) => ({ cards: [{ ref: "R1.1", stale }], summary, checks });
  marker(
    "s3-stale",
    S3.staleEvidence(
      v("Quote changed in the manuscript", "1 failed", ["FAIL R1.1 stale: quote_changed: x"]),
      v("", "All checks passed"),
      "R1.1",
    ),
  );
  marker(
    "s3-ledger",
    S3.ledgerEvidence(
      { checks: ["FAIL R1.1 ledger: L4: Unfortunately"], export_label: "Export anyway" },
      false,
      { summary: "All checks passed" },
    ),
  );
  marker(
    "s3-lang",
    S3.langEvidence(
      "[response] draft R1.1 lang=German model=x words=9 citations=1",
      "5. Write the rationale in German: what = ...",
      "German",
    ),
  );
  fs.writeFileSync(path.join(run, "review_manuscript.response.R1.txt"), "x\n");
  marker(
    "s3-open-check",
    S3.openCheckEvidence(
      { ok: true, opened: true, paragraphs: 12, repaired: false },
      12,
      "/r/a.response.docx",
      "/r/a.response-2.docx",
    ),
  );
  marker(
    "s3-cleanup",
    S3.cleanupEvidence({ cleanup: { own_word_gone: true, preexisting_alive: true }, listening: 0 }),
  );
  for (const [id, c] of Object.entries(C)) assert.equal(exit(c), 0, id);
  fs.rmSync(run, { recursive: true, force: true });
});

// The patterns the response steps wait for are the messages the Response tab
// writes (response-pane.js, response-view.js), and none of them is the line
// shown while a click is held.
test("the response steps wait for the message their own action ends with, with no fixed sleep", async () => {
  const fs = await import("node:fs");
  const read = (p) => fs.readFileSync(new URL(p, import.meta.url), "utf8");
  const driver = read("../scripts/e2e/pane-driver.cjs");
  const at = driver.indexOf("// ---- stage 3: Response tab ----");
  assert.ok(at > 0);
  const part = driver.slice(at, driver.indexOf("\n  view,\n", at));
  assert.doesNotMatch(part, /sleep\(/);
  const waits = [
    ...part.matchAll(/waitResponse\(page, [^,]+, ((?:"(?:[^"\\]|\\.)*")|`[^`]*`)\)/g),
  ].map((m) => m[1]);
  assert.ok(waits.length >= 10, `${waits.length} waits`);
  assert.equal(
    (part.match(/waitResponse\(/g) ?? []).length,
    waits.length,
    "every wait has a pattern",
  );
  const HOLD = "Waiting for the daemon to confirm the stored provider choice...";
  // What the pane says when each action ends.
  const SAMPLES = [
    "Refreshed from the manuscript: 5 comments.",
    "Read the letter: 2 settings differ, 5 earlier answers found. Nothing is saved until you press Apply to this paper.",
    "Applied to this paper: 2 settings, 5 earlier answers kept for drafting.",
    "Saved for this paper: profile generic.",
    "Declined R1.5.",
    "Declined q2 of R1.3.",
    "Drafted R1.1.",
    "Drafted 4 responses.",
    "Nothing to draft: every comment has a response or is declined.",
    "Saved R1.1.",
    "Line numbers updated: 3/3 quotes located.",
    "Checked: All checks passed",
    "Exported: /r/a.response.docx",
    "Exported under a new name: /r/a.response-2.docx (/r/a.response.docx was already there or open in Word and was kept).",
    "Not exported: 1 failed. Fix them, or press Export anyway.",
  ];
  const fixed = waits.filter((w) => w.startsWith('"')).map((w) => new RegExp(JSON.parse(w)));
  for (const re of fixed) {
    assert.doesNotMatch(HOLD, re);
    assert.ok(
      SAMPLES.some((s) => re.test(s)),
      `${re} matches a message of the pane`,
    );
  }
  for (const s of SAMPLES)
    assert.ok(
      fixed.some((re) => re.test(s)) || /^(Declined|Drafted R|Saved R)/.test(s),
      `a step waits for: ${s}`,
    );
  // The words of those messages are in the pane's source.
  const pane =
    read("../taskpane/shared/response-pane.js") + read("../taskpane/shared/response-view.js");
  for (const words of [
    "Refreshed from the manuscript: ",
    "Read the letter: ",
    "Applied to this paper: ",
    "Saved for this paper: profile ",
    "Declined ${c.ref}.",
    "Declined ${q.id} of ${c.ref}.",
    "Drafted ${ref}.",
    'Drafted ${plural(done, "response")}.',
    "Nothing to draft",
    "Saved ${c.ref}.",
    "Line numbers updated: ",
    "Checked: ",
    "Exported: ",
    "Exported under a new name: ",
    "Not exported: ",
  ])
    assert.ok(pane.includes(words), words);
});

// The stage 3 flow of the runner cannot run here (it needs Word): what it asks
// of the driver, of the COM script and of the judges must at least exist.
test("the stage 3 flow uses only pane steps, COM actions and judges that exist, and writes every marker of its features", async () => {
  const fs = await import("node:fs");
  const read = (p) => fs.readFileSync(new URL(p, import.meta.url), "utf8");
  const runner = read("../scripts/e2e/run-acceptance.mjs");
  const driver = read("../scripts/e2e/pane-driver.cjs");
  const ps = read("../scripts/e2e/word-com.ps1");
  const flow = runner.slice(
    runner.indexOf("async function mainS3()"),
    runner.indexOf("// Close the Word this run started"),
  );
  assert.ok(flow.length > 1000);
  const names = (re) => [...new Set([...flow.matchAll(re)].map((m) => m[1]))];
  const steps = names(/\bpane(?:Try)?\(\s*"([a-z-]+)"/g);
  assert.ok(steps.length >= 15, steps.join(","));
  for (const s of steps)
    assert.match(driver, new RegExp(`\\n  (?:async )?"?${s}"?(?:\\(page|: )`), `pane step ${s}`);
  const actions = names(/\bcom\(\s*"([a-z-]+)"/g);
  assert.deepEqual(actions.sort(), [
    "apply",
    "close-other",
    "insert-paragraph",
    "open-check",
    "replace-text",
  ]);
  for (const a of actions) assert.ok(ps.includes(`  "${a}" {`), `COM action ${a}`);
  // Every step that is judged by its own name has a judge.
  for (const [, step, rest] of flow.matchAll(/\bnote\(\s*"([a-z-]+)",([^;]*);/g))
    if (!/judge: "none"/.test(rest))
      assert.equal(typeof S3.JUDGES[step], "function", `judge ${step}`);
  assert.deepEqual(names(/\bmarker\(\s*"(s3-[a-z-]+)"/g).sort(), [
    "s3-check",
    "s3-export",
    "s3-lang",
    "s3-ledger",
    "s3-line-shift",
    "s3-lines",
    "s3-open",
    "s3-open-check",
    "s3-proposal",
    "s3-stale",
  ]);
  // S3-41 is judged after the cleanup, by the marker every later stage writes.
  assert.match(runner, /marker\(\s*\`\$\{STAGE\}-cleanup\`/);
  assert.match(runner, /"after-cleanup": \{ type: "string" \}/);
  assert.match(runner, /STAGE === "s3" \? mainS3\(\)/);
  // No fixed wait in the stage 3 flow.
  assert.doesNotMatch(flow, /\bsleep\(/);
});

test("a check with warnings only passes the later steps of the stage 3 flow, not S3-34", () => {
  const J = S3.JUDGES;
  const warned = { summary: "All checks passed, 1 warning", checks: ["WARN R1.1 lost_facts: x"] };
  assert.equal(J["response-check"](warned).pass, false);
  assert.equal(J["response-check-no-fail"](warned).pass, true);
  assert.equal(J["response-check-no-fail"]({ summary: "All checks passed" }).pass, true);
  assert.equal(J["response-check-no-fail"]({ summary: "1 failed, 1 warning" }).pass, false);
  assert.equal(J["response-check-no-fail"]({ summary: "" }).pass, false);
});
