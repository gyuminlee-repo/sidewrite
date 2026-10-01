// Journal guideline -> rule candidates that cite their sentence -> the ledger
// and the profile of the paper (daemon/aux-service.mjs). The model is a fake
// that labels sentences; the files are real, in a temporary folder.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createAuxService, AUX_JOBS } from "../daemon/aux-service.mjs";
import {
  loadRevision,
  revisionPathFor,
  decisionsPathFor,
  ledgerExportDoc,
  importLedgerItems,
  emptyRevision,
} from "../daemon/revision-store.mjs";
import { effectiveProfile } from "../daemon/journal-profile.mjs";
import { buildReviewPrompt } from "../daemon/review-prompt.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
const GUIDELINE =
  "The abstract must not exceed 250 words. Provide 4 to 6 keywords. We welcome submissions. Statistics symbols are set in italics.";
const LABELS = {
  sentences: [
    { s: 1, type: "count", rule: { target: "section_words", section: "abstract", max: 250 } },
    { s: 2, type: "count", rule: { target: "keywords", min: 4, max: 6 } },
    { s: 3, type: "other" },
    { s: 4, type: "format", rule: { setting: "italics.statistics" } },
  ],
};

async function setup({ reply = { ok: true, json: LABELS } } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "rules-"));
  const doc = join(dir, "paper.docx");
  await writeFile(doc, "docx");
  const asked = [];
  const events = [];
  const service = createAuxService({
    getContext: () => ({ activeDoc: doc }),
    oneshotFor: async () => ({ model: "haiku", run: async (req) => (asked.push(req), reply) }),
    sendEvent: (key, event) => events.push(event),
    log: () => {},
    env: ENV,
  });
  async function call(type, msg = {}) {
    let first;
    await service.handlers[type]({ request_id: "q", ...msg }, (r) => (first = r), "k", "word");
    if (!first.accepted) return first;
    const mine = () => events.find((e) => e.job_id === first.job_id);
    for (let i = 0; i < 100 && !mine(); i++) await new Promise((r) => setTimeout(r, 2));
    return mine().result;
  }
  const revision = () => loadRevision(revisionPathFor(doc), doc);
  return { dir, doc, asked, events, call, revision, done: () => rm(dir, { recursive: true }) };
}

test("extract is a job; every candidate cites a sentence of the guideline and nothing is stored", async () => {
  const t = await setup();
  try {
    assert.ok(AUX_JOBS.has("profile_rules_extract"));
    const r = await t.call("profile_rules_extract", { text: GUIDELINE });
    assert.equal(r.ok, true);
    assert.equal(r.model, "haiku");
    assert.deepEqual(r.warnings, []);
    assert.equal(r.candidates.length, 3);
    for (const c of r.candidates) {
      assert.equal(c.source.kind, "guideline");
      assert.ok(GUIDELINE.includes(c.source.sentence));
    }
    assert.deepEqual(
      r.candidates.map((c) => c.check.type),
      ["count", "count", "llm"],
    );
    assert.deepEqual(r.candidates[2].profile, { italics: { statistics: true } });
    assert.equal(t.asked.length, 1);
    assert.deepEqual((await t.revision()).ledger, []);
    // The receipt of the model's labelling: counts and kinds, no guideline text.
    const receipt = await readFile(decisionsPathFor(t.doc), "utf8");
    const row = JSON.parse(receipt.trim().split("\n").at(-1));
    assert.equal(row.kind, "guideline_rules");
    assert.deepEqual(row.output, {
      candidates: 3,
      kinds: ["count", "count", "format"],
      warnings: 0,
    });
    assert.ok(!receipt.includes("abstract must not"));
  } finally {
    await t.done();
  }
});

test("apply: the picked rules are in the ledger with their sentence, a format rule sets the profile", async () => {
  const t = await setup();
  try {
    const { candidates } = await t.call("profile_rules_extract", { text: GUIDELINE });
    const picked = [candidates[0], candidates[2]];
    const r = await t.call("profile_rules_apply", { text: GUIDELINE, candidates: picked });
    assert.deepEqual(
      {
        ok: r.ok,
        added: r.added,
        skipped: r.skipped,
        added_ids: r.added_ids,
        keys: r.profile_keys,
      },
      { ok: true, added: 2, skipped: 0, added_ids: ["L1", "L2"], keys: ["italics"] },
    );
    const rev = await t.revision();
    assert.deepEqual(
      rev.ledger.map((i) => [i.id, i.check.type, i.source]),
      [
        ["L1", "count", { kind: "guideline", sentence: "The abstract must not exceed 250 words." }],
        ["L2", "llm", { kind: "guideline", sentence: "Statistics symbols are set in italics." }],
      ],
    );
    assert.equal(rev.ledger[0].check.max, 250);
    assert.deepEqual(rev.profile_overrides, { italics: { statistics: true } });
    const { profile } = await effectiveProfile(rev, { docDir: t.dir, env: ENV });
    assert.equal(profile.italics.statistics, true);
    // The Review tab is told to load the ledger again.
    assert.deepEqual(t.events.at(-1), { type: "review_event", event: "ledger_changed" });
    // The same picks again: nothing is added twice.
    const again = await t.call("profile_rules_apply", { text: GUIDELINE, candidates: picked });
    assert.deepEqual([again.added, again.skipped], [0, 2]);
    assert.equal((await t.revision()).ledger.length, 2);
  } finally {
    await t.done();
  }
});

test("apply refuses a rule whose sentence is not in the guideline text, and saves nothing", async () => {
  const t = await setup();
  try {
    const { candidates } = await t.call("profile_rules_extract", { text: GUIDELINE });
    const forged = {
      ...candidates[0],
      source: { kind: "guideline", sentence: "The abstract must not exceed 900 words." },
    };
    for (const msg of [
      { text: GUIDELINE, candidates: [candidates[1], forged] },
      { text: "Another text.", candidates: [candidates[0]] },
      { text: GUIDELINE, candidates: [{ ...candidates[0], source: null }] },
    ]) {
      const r = await t.call("profile_rules_apply", msg);
      assert.equal(r.ok, false);
      assert.match(r.error, /cites a sentence that is not in the guideline text/);
    }
    const none = await t.call("profile_rules_apply", { text: GUIDELINE, candidates: [] });
    assert.match(none.error, /Pick at least one rule/);
    assert.deepEqual((await t.revision()).ledger, []);
  } finally {
    await t.done();
  }
});

test("no text, a model failure: an error or a warning, never a rule without a sentence", async () => {
  let t = await setup();
  try {
    const r = await t.call("profile_rules_extract", { text: "  " });
    assert.deepEqual([r.ok, r.error], [false, "Paste the guideline text first."]);
    assert.equal(t.asked.length, 0);
  } finally {
    await t.done();
  }
  t = await setup({ reply: { ok: false, error: "usage limit" } });
  try {
    const r = await t.call("profile_rules_extract", { text: GUIDELINE });
    assert.equal(r.ok, true);
    assert.deepEqual(r.candidates, []);
    assert.match(r.warnings[0], /guideline labelling failed: usage limit/);
  } finally {
    await t.done();
  }
});

test("a ledger export keeps the guideline sentence and an older export without it still imports", () => {
  const rule = {
    kind: "instruction",
    text: "Provide 4 to 6 keywords.",
    scope: "manuscript",
    check: { type: "count", target: "keywords", min: 4, max: 6 },
    source: { kind: "guideline", sentence: "Provide 4 to 6 keywords." },
  };
  const { revision } = importLedgerItems(emptyRevision({ doc_path: "/x.docx" }), [rule]);
  assert.deepEqual(ledgerExportDoc(revision).items, [rule]);
  const old = { kind: "forbid", text: "novel", scope: "manuscript", check: { type: "llm" } };
  const r = importLedgerItems(revision, [old]);
  assert.equal(r.added[0].source, null);
});

test("the policy pack a profile switches off is out of the review prompt", async () => {
  const t = await setup();
  try {
    const { profile } = await effectiveProfile(
      { journal_profile: null, profile_overrides: { policy_packs: { review_scope: false } } },
      { docDir: t.dir, env: ENV },
    );
    assert.deepEqual(profile.policy_packs, { review_scope: false });
    const comment = { ref: "R1.1", reviewer: 1, text: "Please state the unit." };
    const off = buildReviewPrompt({ comment, profile }).text;
    const on = buildReviewPrompt({ comment }).text;
    assert.match(on, /Do not delete or add anything the reviewer did not request/);
    assert.doesNotMatch(off, /Do not delete or add anything the reviewer did not request/);
  } finally {
    await t.done();
  }
});
