// S2-23 of the acceptance run without a model (scripts/e2e/run-acceptance.mjs
// forbidGuard): an item is put on file for R1.3 that names rule L1, "Allow
// exception" sends its edit through the guarded write, and L2 refuses it
// before Word. Here the same is done on the review service with a Word that
// is not there, and the evidence function is given both outcomes.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import { addBlocked, dropBlocked, loadRevision, saveRevision } from "../daemon/revision-store.mjs";
import {
  FORBID_GUARD_EDIT,
  FORBID_GUARD_SEED_RULE,
  forbidBlockEvidence,
} from "../scripts/e2e/judge.mjs";

const read = (p) => readFileSync(new URL(p, import.meta.url), "utf8");
const LEDGER = JSON.parse(read("./fixtures/review_ledger.json")).items;

async function setup() {
  const dir = await mkdtemp(join(tmpdir(), "forbid-guard-"));
  const doc = join(dir, "paper.docx");
  const lines = [];
  const calls = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: async (name, args) => {
      calls.push([name, args]);
      return { ok: true, word_comment_id: 5, paragraph_index: 2, before: "b", after: "a" };
    },
    pushTurn: () => {},
    oneshot: null,
    sendEvent: () => {},
    log: (s) => lines.push(s),
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q1", ...payload }, resolve, "k", "word"),
    );
  const file = join(dir, "paper.revision.json");
  return { dir, doc, file, ask, lines, calls, cleanup: () => rm(dir, { recursive: true }) };
}

// Word as word-com.ps1 state reports it: two tracked changes, none with the word.
const WORD = {
  ok: true,
  revisions: [
    { type: 1, text: " Celsius", start: 10, end: 18 },
    { type: 2, text: "a novel term the authors had", start: 30, end: 58 },
  ],
};

test("the ledger of the acceptance run: L2 forbids the word of the edit, L1 is another rule", () => {
  assert.deepEqual(
    LEDGER.map((i) => i.kind),
    ["instruction", "forbid", "instruction"],
  );
  const l2 = new RegExp(LEDGER[1].check.pattern, LEDGER[1].check.flags);
  assert.ok(l2.test(FORBID_GUARD_EDIT.replace));
  assert.ok(!l2.test(FORBID_GUARD_EDIT.find));
  assert.equal(FORBID_GUARD_SEED_RULE, "L1");
  assert.equal(FORBID_GUARD_EDIT.comment_ref, "R1.3");
});

test("an exception allowed for L1 does not lift L2: the edit with the forbidden word is refused before Word", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", {
      comments: [
        {
          ref: "R1.3",
          reviewer: 1,
          text: "Say how growth was followed; call the substrate novel.",
        },
      ],
    });
    for (const item of LEDGER) await t.ask("ledger_add", { item });
    // What the run does, with the functions it uses.
    const seeded = addBlocked(await loadRevision(t.file, t.doc), "R1.3", {
      args: FORBID_GUARD_EDIT,
      blocked_by: [FORBID_GUARD_SEED_RULE],
    });
    await saveRevision(t.file, seeded.revision);
    const seedId = seeded.blocked.id;
    const r = await t.ask("review_override", { ref: "R1.3", blocked_id: seedId });
    assert.deepEqual([r.ok, r.applied, r.result.blocked_by], [true, false, ["L2"]]);
    assert.equal(t.calls.length, 0);
    assert.ok(t.lines.includes("[review] blocked R1.3 by=L2"), t.lines.join("\n"));
    const revision = JSON.parse(await readFile(t.file, "utf8"));
    const blocked = revision.comments[0].blocked;
    assert.deepEqual(
      blocked.map((b) => [b.id === seedId, b.blocked_by, b.state]),
      [
        [true, ["L1"], "blocked"],
        [false, ["L2"], "blocked"],
      ],
    );
    assert.deepEqual(revision.comments[0].changes, []);
    // The pane shows the reason as an error (review-pane.js override, guard).
    const step = { message: r.result.message, message_error: true };
    const evidence = forbidBlockEvidence({
      step,
      log: t.lines.join("\n"),
      revision,
      before: seeded.revision,
      word: WORD,
      wordBefore: WORD,
      seedId,
    });
    assert.equal(evidence.pass, true, JSON.stringify(evidence.checks));
    assert.deepEqual(evidence.recorded, [{ id: blocked[1].id, blocked_by: ["L2"] }]);
    // The run takes its own item off the file; the daemon's record stays.
    await saveRevision(t.file, dropBlocked(await loadRevision(t.file, t.doc), "R1.3", seedId));
    const after = JSON.parse(await readFile(t.file, "utf8")).comments[0].blocked;
    assert.deepEqual(
      after.map((b) => b.blocked_by),
      [["L2"]],
    );
  } finally {
    await t.cleanup();
  }
});

test("forbidBlockEvidence: every way the guard could have failed is a fail", () => {
  const row = (id, by, more = {}) => ({
    id,
    args: FORBID_GUARD_EDIT,
    blocked_by: by,
    state: "blocked",
    ...more,
  });
  const good = {
    step: {
      message:
        "Not applied: the new text, alone or with the words next to it, matches forbidden rule(s) L2. Rephrase without it.",
      message_error: true,
    },
    log: "[agent] x\n[review] blocked R1.3 by=L2\n",
    revision: {
      comments: [
        {
          ref: "R1.1",
          changes: [{ id: "R1.1-1", text: " Celsius", override_of: [] }],
          blocked: [],
        },
        {
          ref: "R1.3",
          changes: [{ id: "R1.3-1", text: "an approved novel one", override_of: ["L2"] }],
          blocked: [row("R1.3-b0", ["L2"]), row("R1.3-b1", ["L1"]), row("R1.3-b2", ["L2"])],
        },
      ],
    },
    // Before the step: the seeded item, and an edit of a turn the guard
    // blocked earlier (not the record of this step).
    before: {
      comments: [{ ref: "R1.3", blocked: [row("R1.3-b0", ["L2"]), row("R1.3-b1", ["L1"])] }],
    },
    word: WORD,
    wordBefore: WORD,
    seedId: "R1.3-b1",
  };
  assert.deepEqual(forbidBlockEvidence(good).checks, {
    refused: true,
    logged: true,
    recorded: true,
    seed_kept: true,
    no_change: true,
    word_unchanged: true,
  });
  const failed = (patch) => {
    const e = forbidBlockEvidence({ ...good, ...patch });
    assert.equal(e.pass, false);
    return Object.entries(e.checks)
      .filter(([, v]) => !v)
      .map(([k]) => k);
  };
  // The edit went in: no error on the message line.
  assert.deepEqual(failed({ step: { message: "", message_error: false } }), ["refused"]);
  // Refused by something else than the rule.
  assert.deepEqual(
    failed({ step: { message: "Not applied: the task pane is older", message_error: true } }),
    ["refused"],
  );
  assert.deepEqual(failed({ step: { ok: false, error: "locator.click: Timeout" } }), ["refused"]);
  // The line of another comment, or of another rule, is not this block.
  assert.deepEqual(
    failed({ log: "[review] blocked R1.30 by=L2\n[review] blocked R1.3 by=L2,L3" }),
    ["logged"],
  );
  // The daemon recorded no blocked edit of its own.
  const only = structuredClone(good.revision);
  only.comments[1].blocked.pop();
  assert.deepEqual(failed({ revision: only }), ["recorded"]);
  // The seeded item is gone: it was applied.
  const gone = structuredClone(good.revision);
  gone.comments[1].blocked.splice(1, 1);
  assert.deepEqual(failed({ revision: gone }), ["seed_kept"]);
  // A change holds the word without an exception for the rule.
  const changed = structuredClone(good.revision);
  changed.comments[1].changes.push({
    id: "R1.3-2",
    text: "novel synthetic carbon source",
    override_of: ["L1"],
  });
  const e = forbidBlockEvidence({ ...good, revision: changed });
  assert.deepEqual([e.checks.no_change, e.changes_with_the_word], [false, ["R1.3-2"]]);
  // Word has one more tracked change, or an insertion with the word.
  const more = { ok: true, revisions: [...WORD.revisions, { type: 1, text: "x" }] };
  assert.deepEqual(failed({ word: more }), ["word_unchanged"]);
  const swapped = { ok: true, revisions: [WORD.revisions[0], { type: 1, text: "novel " }] };
  assert.deepEqual(failed({ word: swapped }), ["word_unchanged"]);
  assert.deepEqual(failed({ word: { ok: false, error: "x" } }), ["word_unchanged"]);
});

test("the run sends the edit after the model's turns and before the author's decisions in Word", () => {
  const runner = read("../scripts/e2e/run-acceptance.mjs");
  const main = runner.slice(runner.indexOf("// Stage 2 (procedure of the stage-2 plan"));
  const knob = main.indexOf('await propose("R1.3", "propose R1.3 (forbid list off)");');
  const guard = main.indexOf("await forbidGuard();");
  const word = main.indexOf('label: "refresh before Word"');
  assert.ok(knob > 0 && knob < guard && guard < word, `${knob} ${guard} ${word}`);
  assert.equal((runner.match(/await forbidGuard\(\);/g) ?? []).length, 1);
  // The knob still depends on what the model did, not on this step.
  assert.ok(main.indexOf('!opt["skip-knob"]') < guard);
  const fn = runner.slice(runner.indexOf("async function forbidGuard()"));
  const body = fn.slice(0, fn.indexOf("\n}\n"));
  assert.match(body, /blocked_by: \[FORBID_GUARD_SEED_RULE\]/);
  assert.match(
    body,
    /pane\("allow-exception", \{ ref: "R1\.3", blocked_by: FORBID_GUARD_SEED_RULE \}\)/,
  );
  assert.match(body, /marker\("s2-forbid-block", evidence\)/);
  assert.ok(body.indexOf('marker("s2-forbid-block"') < body.indexOf("dropBlocked("));
  // The step of the driver, and the words of the pane it looks for.
  const driver = read("../scripts/e2e/pane-driver.cjs");
  assert.match(driver, /\n  async "allow-exception"\(page, a\) \{/);
  const paneSrc = read("../taskpane/shared/review-pane.js");
  for (const words of ["Allow exception", "Apply despite ", "Blocked by "])
    assert.ok(paneSrc.includes(words) && driver.includes(words.trim()), words);
});
