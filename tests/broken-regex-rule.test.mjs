// A regex rule that does not compile (H3). The input schema refuses it, a
// file that already holds one still loads, and a tracked edit is refused
// while an active forbid rule cannot be checked: a rule that cannot be read
// is not a rule that passed.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  LedgerItemSchema,
  emptyRevision,
  addLedgerItem,
  updateLedgerItem,
  loadRevision,
  saveRevision,
  replaceComments,
  importLedgerItems,
  ledgerExportDoc,
} from "../daemon/revision-store.mjs";
import * as ledgerCheck from "../daemon/ledger-check.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";
import { createReviewService } from "../daemon/review-service.mjs";

const regex = (pattern, flags = "", expect = "absent") => ({
  type: "regex",
  pattern,
  flags,
  expect,
});
const forbid = (id, pattern, over = {}) => ({
  id,
  kind: "forbid",
  text: `rule ${id}`,
  check: regex(pattern, "i"),
  ...over,
});
const NOVEL = forbid("L1", String.raw`\bnovel\b`);
const BROKEN = forbid("L2", "(");

test("schema: a pattern or flags that do not compile are refused on add and update; good ones pass", () => {
  const rev = emptyRevision("/m/paper.docx");
  const item = (check) => ({ kind: "forbid", text: "x", check });
  for (const check of [
    regex("("),
    regex("[a-"),
    regex("a", "ii"), // every letter is allowed, the pair is not
    regex(String.raw`\p{Nope}`, "u"),
  ]) {
    assert.throws(() => addLedgerItem(rev, item(check)), /regex|flags/i, JSON.stringify(check));
    assert.equal(LedgerItemSchema.safeParse({ id: "L1", ...item(check) }).success, false);
  }
  const added = addLedgerItem(rev, item(regex(String.raw`\bnovel\b`, "iu")));
  assert.equal(added.item.id, "L1");
  assert.equal(addLedgerItem(rev, item({ type: "llm" })).item.check.type, "llm");
  assert.throws(() => updateLedgerItem(added.revision, "L1", { check: regex("(") }), /regex/i);
  const kept = updateLedgerItem(added.revision, "L1", { check: regex("novelty") });
  assert.equal(kept.ledger[0].check.pattern, "novelty");
});

test("a file that already holds a broken rule still loads and saves, and a result can be stored on it", async () => {
  const dir = await mkdtemp(join(tmpdir(), "broken-rule-"));
  const path = join(dir, "paper.revision.json");
  try {
    const stored = { ...emptyRevision("/m/paper.docx"), ledger: [BROKEN, NOVEL], ledger_seq: 2 };
    await writeFile(path, JSON.stringify(stored), "utf8");
    const rev = await loadRevision(path, "/m/paper.docx");
    assert.deepEqual(
      rev.ledger.map((i) => i.check.pattern),
      ["(", String.raw`\bnovel\b`],
    );
    // ledger_check writes its error on the broken item with the other results.
    const next = updateLedgerItem(rev, "L2", { status: "error" });
    assert.equal(next.ledger[0].status, "error");
    await saveRevision(path, next);
    assert.equal(JSON.parse(await readFile(path, "utf8")).ledger[0].status, "error");
    // The way out: a check that compiles.
    assert.equal(updateLedgerItem(next, "L2", { check: regex("x") }).ledger[0].check.pattern, "x");
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("brokenForbidRules: the pre-check rules that do not compile, by id", () => {
  const { brokenForbidRules, checkRegexItem } = ledgerCheck;
  assert.equal(typeof brokenForbidRules, "function");
  const ledger = [
    NOVEL,
    BROKEN,
    forbid("L3", "(", { check: regex("(", "", "present") }), // must-keep: not pre-checked
    forbid("L4", "(", { kind: "instruction" }),
    forbid("L5", "(", { scope: "response" }),
    forbid("L6", "a", { check: regex("a", "ii") }),
    { id: "L7", kind: "forbid", text: "x", check: { type: "llm" } },
  ];
  const broken = brokenForbidRules(ledger);
  assert.deepEqual(
    broken.map((b) => b.id),
    ["L2", "L6"],
  );
  assert.ok(broken.every((b) => typeof b.error === "string" && b.error));
  assert.deepEqual(brokenForbidRules([NOVEL]), []);
  assert.deepEqual(brokenForbidRules(null), []);
  // The ledger check itself still reports the rule as an error.
  assert.equal(checkRegexItem(BROKEN, [{ index: 0, text: "x" }]).status, "error");
});

const R = { what: "w", why: "y", addresses: "a" };
function deps(ledger) {
  const d = {
    calls: [],
    edits: [],
    blocked: [],
    getLedger: async () => ledger,
    call: async (name, args) => {
      d.calls.push([name, args]);
      return { ok: true, word_comment_id: 1, paragraph_index: 0, before: "b", after: "a" };
    },
    onEdit: async ({ ref, change }) => (d.edits.push(change), { id: `${ref}-1` }),
    onBlocked: async (b) => (d.blocked.push(b), { id: `${b.ref}-b1` }),
  };
  return d;
}
const EDIT = { comment_ref: "R1.1", find: "material", replace: "substrate", rationale: R };

test("tracked edit: a stored forbid rule that does not compile refuses the edit and names the rule", async () => {
  const d = deps([NOVEL, BROKEN]);
  const r = await applyTrackedEdit(d, EDIT, { ref: "R1.1" });
  assert.deepEqual([r.ok, r.code, r.broken_rules], [false, "ForbidRuleBroken", ["L2"]]);
  assert.match(r.message, /L2/);
  assert.match(r.message, /not applied/i);
  assert.deepEqual([d.calls.length, d.edits.length, d.blocked.length], [0, 0, 0]);
  // An approved exception does not get past it either.
  const o = await applyTrackedEdit(d, EDIT, { ref: "R1.1" }, { override: { blocked_by: ["L2"] } });
  assert.deepEqual([o.ok, o.code, d.calls.length], [false, "ForbidRuleBroken", 0]);
});

test("tracked edit: a broken rule that is no pre-check rule (must-keep, instruction, other scope) does not stop the edit", async () => {
  const d = deps([
    NOVEL,
    forbid("L3", "(", { check: regex("(", "", "present") }),
    forbid("L4", "(", { kind: "instruction" }),
    forbid("L5", "(", { scope: "response" }),
  ]);
  const r = await applyTrackedEdit(d, EDIT, { ref: "R1.1" });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual([d.calls.length, d.edits.length], [1, 1]);
});

test("service: ledger_add and ledger_update refuse a broken regex; a stored one blocks the tracked edit and ledger_check keeps the other results", async () => {
  const dir = await mkdtemp(join(tmpdir(), "broken-rule-svc-"));
  const doc = join(dir, "paper.docx");
  const calls = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: null,
    pushTurn: () => {},
    log: () => {},
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) => {
      const reply = (r) => {
        if (!r.accepted) resolve(r);
      };
      svc.handlers[type]({ type, request_id: "q", ...payload }, reply, "k", "word");
    });
  try {
    const item = { kind: "forbid", text: "x", check: regex("(") };
    const bad = await ask("ledger_add", { item });
    assert.equal(bad.ok, false);
    assert.match(bad.error, /regex/i);
    const good = await ask("ledger_add", { item: { ...item, check: regex("novel", "i") } });
    assert.equal(good.ok, true, good.error);
    const upd = await ask("ledger_update", { id: "L1", patch: { check: regex("a", "ii") } });
    assert.equal(upd.ok, false);

    // A file written before the schema checked the pattern.
    const path = join(dir, "paper.revision.json");
    const rev = replaceComments(await loadRevision(path, doc), [
      { ref: "R1.1", reviewer: 1, text: "c" },
    ]);
    const stored = { ...rev, ledger: [...rev.ledger, BROKEN], ledger_seq: 2 };
    await writeFile(path, JSON.stringify(stored), "utf8");
    const loaded = await ask("review_load");
    assert.equal(loaded.ok, true, loaded.error);
    assert.equal(loaded.revision.ledger.length, 2);

    const r = await applyTrackedEdit(
      { ...svc.trackedEditDeps("k"), call: async (...a) => (calls.push(a), { ok: true }) },
      EDIT,
      { ref: "R1.1" },
    );
    assert.deepEqual(
      [r.ok, r.code, r.broken_rules, calls.length],
      [false, "ForbidRuleBroken", ["L2"], 0],
    );

    // The check of both rules: the broken one is an error, the other is saved too.
    const done = new Promise((resolve) => {
      svc.handlers.ledger_check(
        {
          type: "ledger_check",
          request_id: "c",
          snapshot: { paragraphs: [{ index: 0, text: "a novel one" }] },
        },
        () => {},
        "k",
        "word",
      );
      const poll = async () => {
        const now = await loadRevision(path, doc);
        if (now.ledger.every((i) => i.status !== "unchecked")) resolve(now);
        else setTimeout(poll, 10);
      };
      poll();
    });
    const after = await done;
    assert.deepEqual(
      after.ledger.map((i) => [i.id, i.status]),
      [
        ["L1", "violations"],
        ["L2", "error"],
      ],
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

// Ledger import (ledger_import, also the way ledger_suggest candidates come
// in) goes through the input schema: a file with one rule that does not
// compile adds nothing, in append and in replace mode.
test("ledger import: a file with a regex that does not compile is refused as a whole", async () => {
  const rule = (text, pattern, flags = "i") => ({
    kind: "forbid",
    text,
    check: regex(pattern, flags),
  });
  const good = [rule("no novel", "novel"), rule("no robust", "robust")];
  const base = importLedgerItems(emptyRevision("/m/paper.docx"), [rule("kept", "kept")]).revision;
  for (const [name, bad, why] of [
    ["pattern", rule("broken", "("), /item 2: check\.pattern: invalid regex/],
    ["flags", rule("broken", "a", "ii"), /item 2: check\.pattern: invalid regex/],
  ]) {
    for (const mode of ["append", "replace"]) {
      assert.throws(
        () => importLedgerItems(base, [good[0], bad, good[1]], { mode }),
        why,
        `${name} ${mode}`,
      );
    }
  }
  assert.deepEqual([base.ledger.map((i) => i.id), base.ledger_seq], [["L1"], 1]);
  assert.equal(importLedgerItems(base, good).added.length, 2);

  // Through the service: the file on disk is as it was, the counter too.
  const dir = await mkdtemp(join(tmpdir(), "broken-rule-import-"));
  const doc = join(dir, "paper.docx");
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: null,
    pushTurn: () => {},
    log: () => {},
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
    );
  try {
    assert.equal((await ask("ledger_add", { item: rule("kept", "kept") })).ok, true);
    const path = join(dir, "paper.revision.json");
    const before = await readFile(path, "utf8");
    for (const mode of ["append", "replace"]) {
      const r = await ask("ledger_import", {
        doc: { items: [good[0], rule("broken", "("), good[1]] },
        mode,
      });
      assert.equal(r.ok, false, mode);
      assert.match(r.error, /item 2: check\.pattern: invalid regex/, mode);
      assert.equal(await readFile(path, "utf8"), before, `${mode}: nothing written`);
    }
    const ok = await ask("ledger_import", { doc: { items: good } });
    assert.deepEqual([ok.ok, ok.added_ids], [true, ["L2", "L3"]]);

    // A rule a file from before the check still holds is exported as it is,
    // and that export is refused on import: the rule is fixed first.
    const rev = await loadRevision(path, doc);
    await writeFile(path, JSON.stringify({ ...rev, ledger: [...rev.ledger, forbid("L4", "(")] }));
    const out = ledgerExportDoc(await loadRevision(path, doc));
    assert.equal(out.items.at(-1).check.pattern, "(");
    const back = await ask("ledger_import", { doc: out, mode: "replace" });
    assert.equal(back.ok, false);
    assert.match(back.error, /item 4: check\.pattern: invalid regex/);
    assert.equal((await loadRevision(path, doc)).ledger.length, 4, "replace emptied nothing");
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

// No edit can be applied while an active forbid rule does not compile, so no
// review turn is started either (one turn per comment in a Propose all run).
test("service: review_propose is refused while a stored forbid rule does not compile", async () => {
  const dir = await mkdtemp(join(tmpdir(), "broken-rule-propose-"));
  const doc = join(dir, "paper.docx");
  const pushed = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: null,
    pushTurn: (...a) => pushed.push(a),
    log: () => {},
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
    );
  try {
    const path = join(dir, "paper.revision.json");
    const rev = replaceComments(await loadRevision(path, doc), [
      { ref: "R1.1", reviewer: 1, text: "c" },
    ]);
    // A must-keep rule that does not compile stops no edit, so it stops no turn.
    const keep = forbid("L1", "(", { check: regex("(", "", "present") });
    await writeFile(path, JSON.stringify({ ...rev, ledger: [keep], ledger_seq: 1 }));
    const ok = await ask("review_propose", { ref: "R1.1", paragraphs: [] });
    assert.deepEqual([ok.ok, pushed.length], [true, 1], ok.error);

    const now = await loadRevision(path, doc);
    await writeFile(path, JSON.stringify({ ...now, ledger: [keep, BROKEN], ledger_seq: 2 }));
    const r = await ask("review_propose", { ref: "R1.1", paragraphs: [] });
    assert.equal(r.ok, false);
    assert.match(r.error, /L2/);
    assert.match(r.error, /invalid regex/i);
    assert.equal(pushed.length, 1, "no second turn");
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
