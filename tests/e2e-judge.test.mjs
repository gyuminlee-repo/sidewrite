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
