// The paragraph of a tracked edit is compared right before the edit (H2).
// The daemon judges an edit (forbid list, lost facts) on the paragraph as
// the last status refresh and the recorded changes give it. Word edits the
// paragraph as it is NOW. When the author changed it in between, the edit
// was judged on a text that is no longer there: forbid "novel substrate",
// snapshot "A useful material.", document "A novel material.", edit
// material -> substrate passed the check and wrote the forbidden phrase.
// The daemon now sends the hash of the paragraph it judged; the task pane
// hashes the paragraph it is about to edit, in the same Word.run, and
// changes nothing when they differ (StaleTarget).
import { plainPane } from "./helpers/plain-pane.mjs";
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";
import { paragraphHash } from "../daemon/doc-diff.mjs";
import { createReviewService } from "../daemon/review-service.mjs";

// A fake Word document. paras: strings, or { text, reviewed } for a
// paragraph with pending tracked changes (text: raw, with the deleted words
// still in it; reviewed: what getReviewedText(current) gives).
// reviewedApi false: a Word without getReviewedText (the task pane then
// reads the raw text, as its snapshot does).
function fakeDoc(paras, { mode = "uniqueLocalId", reviewedApi = true } = {}) {
  const state = { edits: [], comments: [], mode: "Off", modeSets: 0 };
  const coll = (items) => ({ items, load() {} });
  const hits = (text, q) => {
    const out = [];
    for (let i = text.indexOf(q); i !== -1; i = text.indexOf(q, i + q.length)) out.push(i);
    return out;
  };
  const range = (p, what) => ({
    owner: p,
    load() {},
    paragraphs: { getFirst: () => p },
    search: (q) => coll(hits(what ?? p.text, q).map(() => range(p, q))),
    insertText: (s, where) => (state.edits.push({ paragraph: p.k, text: s, where }), range(p, s)),
    insertComment: (c) => (state.comments.push(c), { id: `C${state.comments.length}`, load() {} }),
    compareLocationWith: (other) => ({ value: other.owner === p ? "Equal" : "Before" }),
    getReviewedText() {
      if (!reviewedApi) throw new Error("ApiNotFound: getReviewedText");
      return { value: p.reviewed ?? p.text };
    },
  });
  const items = paras.map((x, k) => {
    const p = {
      k,
      text: typeof x === "string" ? x : x.text,
      reviewed: typeof x === "string" ? null : x.reviewed,
      style: "Normal",
      ...(mode === "uniqueLocalId" ? { uniqueLocalId: x.id ?? `U${k}` } : {}),
      load() {},
      getRange: () => range(p),
      search: (q) => coll(hits(p.text, q).map(() => range(p, q))),
    };
    return p;
  });
  const document = {
    load() {},
    get changeTrackingMode() {
      return state.mode;
    },
    set changeTrackingMode(v) {
      state.mode = v;
      state.modeSets += 1;
    },
    getSelection: () => ({ select() {} }),
    body: {
      paragraphs: coll(items),
      search: (q) => coll(items.flatMap((p) => hits(p.text, q).map(() => range(p, q)))),
    },
  };
  globalThis.Word = {
    run: async (fn) => fn({ document, sync: async () => {} }),
    InsertLocation: { replace: "Replace", after: "After" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    ChangeTrackingVersion: { current: "Current" },
  };
  return state;
}

const { toolTrackedEdit } = await import("../taskpane/shared/tools-word.js");
const view = await import("../taskpane/shared/review-view.js");

const COMMENT = "[R1.1] w / y / a";
const EDIT = { find: "material", replace: "substrate", comment_text: COMMENT };
const hashed = (text, extra = {}) => ({
  ...EDIT,
  expected_before_hash: paragraphHash(text),
  ...extra,
});
const untouched = (s) => [s.edits.length, s.comments.length, s.mode];

test("the task pane and the daemon hash a paragraph identically", () => {
  const corpus = [
    "plain",
    "  lead and trail  ",
    "a\u0005b\u0005\u0005c",
    "tab\tand\nnewline\r",
    "30 °C and 5 µm",
    "한글 문단과 emoji 😀 and a surrogate pair 𝒳",
    "",
    null,
    " nbsp ",
    "A useful material.",
    "A novel material.",
  ];
  for (const s of corpus) assert.equal(view.paragraphHash(s), paragraphHash(s), JSON.stringify(s));
  assert.notEqual(
    view.paragraphHash("A useful material."),
    view.paragraphHash("A novel material."),
  );
  // Normalized first: spacing and comment anchors do not make a paragraph "changed".
  assert.equal(
    view.paragraphHash(" A  useful\u0005 material. "),
    paragraphHash("A useful material."),
  );
});

test("task pane: the hash matches, the edit goes in as before and says it was checked", async () => {
  const s = fakeDoc(["Title", "A useful material."]);
  const r = await toolTrackedEdit(hashed("A useful material."));
  assert.deepEqual(
    [r.ok, r.paragraph_index, r.before, r.after, r.target_checked],
    [true, 1, "A useful material.", "A useful substrate.", true],
  );
  assert.deepEqual(untouched(s), [1, 1, "Off"]);
});

test("task pane: a word next to the edit changed since the daemon read the paragraph: StaleTarget, nothing changed", async () => {
  const s = fakeDoc(["Title", "A novel material."]);
  const r = await toolTrackedEdit(
    hashed("A useful material.", { expected_paragraph: { index: 1, id: "U1" } }),
  );
  assert.deepEqual([r.ok, r.code, r.applied], [false, "StaleTarget", undefined]);
  assert.deepEqual([r.paragraph_index, r.paragraph_id], [1, "U1"]);
  assert.match(r.message, /changed|not as/i);
  assert.deepEqual(untouched(s), [0, 0, "Off"], "no edit, no comment, track mode as it was");
});

test("task pane: no hash to compare with (a daemon that could not tell the paragraph): refused, not applied", async () => {
  const s = fakeDoc(["Title", "A useful material."]);
  for (const bad of [undefined, null, "", 12345678]) {
    const r = await toolTrackedEdit({ ...EDIT, expected_before_hash: bad });
    assert.deepEqual([r.ok, r.code], [false, "StaleTarget"], String(bad));
  }
  assert.deepEqual(untouched(s), [0, 0, "Off"]);
});

test("task pane: the paragraph moved (its index changed, its text did not): applied where it is now", async () => {
  for (const mode of ["uniqueLocalId", "index"]) {
    const s = fakeDoc(["Title", "A new paragraph of the author.", "A useful material."], { mode });
    // The daemon saw it at index 1.
    const r = await toolTrackedEdit(
      hashed("A useful material.", { expected_paragraph: { index: 1, id: null } }),
    );
    assert.deepEqual([r.ok, r.paragraph_index], [true, 2], mode);
    assert.deepEqual(
      s.edits.map((e) => e.paragraph),
      [2],
      mode,
    );
  }
});

test("task pane: the same words in several paragraphs: the paragraph the hint picks is the one compared", async () => {
  const texts = ["A useful material.", "A plain material.", "A useful material."];
  for (const mode of ["uniqueLocalId", "index"]) {
    const hint = (k) => (mode === "uniqueLocalId" ? `U${k}` : `p${k}`);
    // The hinted paragraph is as the daemon read it.
    let s = fakeDoc(texts, { mode });
    let r = await toolTrackedEdit(hashed(texts[1], { paragraph_hint: hint(1) }));
    assert.deepEqual([r.ok, r.paragraph_index], [true, 1], mode);
    assert.deepEqual(
      s.edits.map((e) => e.paragraph),
      [1],
      mode,
    );
    // The hash of another paragraph does not pass for it.
    s = fakeDoc(texts, { mode });
    r = await toolTrackedEdit(hashed(texts[0], { paragraph_hint: hint(1) }));
    assert.deepEqual([r.ok, r.code, s.edits.length], [false, "StaleTarget", 0], mode);
    // A twin with the same text reads the same, so the check holds for it.
    s = fakeDoc(texts, { mode });
    r = await toolTrackedEdit(hashed(texts[0], { paragraph_hint: hint(2) }));
    assert.deepEqual([r.ok, r.paragraph_index], [true, 2], mode);
    // Without a hint the answer is still Ambiguous, before any comparison.
    s = fakeDoc(texts, { mode });
    r = await toolTrackedEdit(hashed(texts[0]));
    assert.deepEqual([r.ok, r.code, s.edits.length], [false, "Ambiguous", 0], mode);
  }
});

test("task pane: a paragraph with pending tracked changes is compared in its reviewed form, as the snapshot is", async () => {
  // The author replaced "useful" by "novel", tracked: the raw text holds both.
  const para = { text: "A usefulnovel material.", reviewed: "A novel material." };
  let s = fakeDoc(["Title", para]);
  let r = await toolTrackedEdit(hashed("A novel material."));
  assert.deepEqual([r.ok, s.edits.length], [true, 1]);
  s = fakeDoc(["Title", para]);
  r = await toolTrackedEdit(hashed("A usefulnovel material."));
  assert.deepEqual([r.ok, r.code, s.edits.length], [false, "StaleTarget", 0]);
  s = fakeDoc(["Title", para]);
  r = await toolTrackedEdit(hashed("A useful material."));
  assert.deepEqual([r.ok, r.code, s.edits.length], [false, "StaleTarget", 0]);
  // A Word without getReviewedText: the raw text, the fallback of the snapshot.
  s = fakeDoc(["Title", para], { reviewedApi: false });
  r = await toolTrackedEdit(hashed("A usefulnovel material."));
  assert.deepEqual([r.ok, s.edits.length], [true, 1]);
});

// ---- the daemon side -------------------------------------------------------------

const R = { what: "w", why: "y", addresses: "a" };
const ARGS = { comment_ref: "R1.1", find: "material", replace: "substrate", rationale: R };
function deps(ctx, result) {
  const d = {
    calls: [],
    edits: [],
    notApplied: [],
    stale: [],
    lines: [],
    log: (s) => d.lines.push(s),
    getLedger: async () => [],
    getReviewContext: async () => ctx,
    onAttempt: async () => ({ id: "R1.1-b1" }),
    call: async (name, args) => {
      d.calls.push(args);
      return typeof result === "function" ? result(args) : result;
    },
    onEdit: async ({ change }) => (d.edits.push(change), { id: "R1.1-1" }),
    onNotApplied: async (x) => d.notApplied.push(x.attempt_id),
    onStaleTarget: async (x) => d.stale.push(x.ref),
  };
  return d;
}
const OK = {
  ok: true,
  word_comment_id: 1,
  paragraph_index: 1,
  before: "A useful material.",
  after: "A useful substrate.",
  target_checked: true,
};

test("daemon: the call carries the hash of the paragraph it judged, with the changes recorded since the snapshot in it", async () => {
  const ctx = { paragraphs: ["Title", "A useful material."], paragraph_ids: ["U0", "U1"] };
  let d = deps(ctx, OK);
  await applyTrackedEdit(d, ARGS, { ref: "R1.1" });
  assert.equal(d.calls[0].expected_before_hash, paragraphHash("A useful material."));
  assert.deepEqual(d.calls[0].expected_paragraph, { index: 1, id: "U1" });

  // An earlier edit of this turn made it "A handy material.".
  d = deps(
    {
      ...ctx,
      recent_changes: [{ kind: "replace", find: "useful", text: "handy", paragraph_index: 1 }],
    },
    OK,
  );
  await applyTrackedEdit(d, ARGS, { ref: "R1.1" });
  assert.equal(d.calls[0].expected_before_hash, paragraphHash("A handy material."));

  // The paragraph cannot be told from here (two hits, no hint): no hash, and
  // the task pane refuses unless it cannot tell either.
  d = deps({ paragraphs: ["A useful material.", "A plain material."] }, OK);
  await applyTrackedEdit(d, ARGS, { ref: "R1.1" });
  assert.equal("expected_before_hash" in d.calls[0], false);
  assert.equal("expected_paragraph" in d.calls[0], false);
});

test("daemon: StaleTarget is not recorded, its attempt is taken back, and the model is told to read again", async () => {
  const ctx = { paragraphs: ["Title", "A useful material."], paragraph_ids: ["U0", "U1"] };
  const d = deps(ctx, { ok: false, code: "StaleTarget", message: "pane text", paragraph_index: 1 });
  const r = await applyTrackedEdit(d, ARGS, { ref: "R1.1" });
  assert.deepEqual([r.ok, r.code, r.applied], [false, "StaleTarget", undefined]);
  assert.match(r.message, /not applied/i);
  assert.match(r.message, /read the paragraph again/i);
  assert.deepEqual([d.edits.length, d.notApplied, d.stale], [0, ["R1.1-b1"], ["R1.1"]]);
  assert.ok(
    d.lines.some((l) => /stale target R1\.1/.test(l)),
    d.lines.join("\n"),
  );
});

test("daemon: an applied edit the task pane did not check (an old task pane) is recorded and flagged", async () => {
  const ctx = { paragraphs: ["Title", "A useful material."], paragraph_ids: ["U0", "U1"] };
  const d = deps(ctx, { ...OK, target_checked: undefined });
  const r = await applyTrackedEdit(d, ARGS, { ref: "R1.1" });
  assert.deepEqual([r.ok, d.edits.length], [true, 1], "it is in the document, so it is recorded");
  assert.match(r.warning, /did not compare|reload the task pane/i);
  const checked = deps(ctx, OK);
  const ok = await applyTrackedEdit(checked, ARGS, { ref: "R1.1" });
  assert.doesNotMatch(ok.warning ?? "", /reload the task pane/i);
});

// ---- daemon and task pane together -------------------------------------------------

const snap = (texts) => ({
  paragraphs: texts.map((text, index) => ({ index, id: `U${index}`, text, w_ins: 0, w_del: 0 })),
  comments: [],
});

// The review service of one document whose task pane is toolTrackedEdit on
// the fake Word of the moment.
async function world() {
  const dir = await mkdtemp(join(tmpdir(), "stale-target-"));
  const doc = join(dir, "paper.docx");
  const events = [];
  const calls = [];
  const callTool = async (name, args) => (calls.push(args), toolTrackedEdit(args));
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool,
    pushTurn: () => {},
    sendEvent: (key, obj) => events.push(obj),
    log: () => {},
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...plainPane(payload) }, resolve, "k", "word"),
    );
  const edit = (args = ARGS) =>
    applyTrackedEdit({ ...svc.trackedEditDeps("k"), call: callTool }, args, { ref: "R1.1" });
  await ask("review_save_comments", { comments: [{ ref: "R1.1", reviewer: 1, text: "Rename." }] });
  await ask("ledger_add", {
    item: {
      kind: "forbid",
      text: "no novel substrate",
      check: { type: "regex", pattern: "novel substrate", flags: "i" },
    },
  });
  return { ask, edit, events, calls, cleanup: () => rm(dir, { recursive: true, force: true }) };
}

test("service: an undo that took more than the edit asks the pane of that document for a refresh, on the comment of the turn", async () => {
  const dir = await mkdtemp(join(tmpdir(), "undo-more-"));
  const events = [];
  const lines = [];
  try {
    const svc = createReviewService({
      getContext: () => ({ activeDoc: join(dir, "paper.docx"), caps: ["target_hash"] }),
      callTool: async () => ({}),
      pushTurn: () => {},
      sendEvent: (key, obj) => events.push([key, obj]),
      log: (l) => lines.push(l),
      env: { CC_OFFICE_WSL_PATHS: "0" },
    });
    await svc.trackedEditDeps("k").onUndoTookMore({ ref: "R1.1", paragraph_index: 4 });
    assert.deepEqual(events, [
      ["k", { type: "review_event", event: "undo_took_more", ref: "R1.1", paragraph_index: 4 }],
    ]);
    assert.deepEqual(lines, ["[review] undo took more R1.1 paragraph=4: refresh asked"]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("together: the author changes a word next to the edit after the snapshot; the edit is refused before Word is touched", async () => {
  const w = await world();
  try {
    await w.ask("review_refresh", { snapshot: snap(["Title", "A useful material."]) });
    // The author types in Word while the model works.
    let word = fakeDoc(["Title", "A novel material."]);
    const r = await w.edit();
    assert.deepEqual([r.ok, r.code], [false, "StaleTarget"], JSON.stringify(r));
    assert.match(r.message, /read the paragraph again/i);
    assert.deepEqual(untouched(word), [0, 0, "Off"], "Word unchanged");
    let c = (await w.ask("review_load")).revision.comments[0];
    assert.deepEqual(
      [c.changes.length, c.blocked.length],
      [0, 0],
      "nothing recorded, no attempt left",
    );
    // The task pane is asked to refresh, so the next try is judged on the text as it is.
    assert.deepEqual(
      w.events.map((e) => [e.event, e.ref]),
      [["stale_target", "R1.1"]],
    );
    await w.ask("review_refresh", { snapshot: snap(["Title", "A novel material."]) });
    word = fakeDoc(["Title", "A novel material."]);
    const again = await w.edit();
    assert.deepEqual([again.ok, again.blocked_by], [false, ["L1"]], "now the forbid list sees it");
    assert.equal(word.edits.length, 0);
  } finally {
    await w.cleanup();
  }
});

test("together: the paragraph is as the daemon read it: applied and recorded, also after an earlier edit of the same turn", async () => {
  const w = await world();
  try {
    await w.ask("review_refresh", { snapshot: snap(["Title", "A useful material."]) });
    let word = fakeDoc(["Title", "A useful material."]);
    const r = await w.edit();
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(word.edits.length, 1);
    assert.doesNotMatch(r.warning ?? "", /reload the task pane/i);
    // The paragraph now holds a pending tracked replace; the next edit of it
    // is compared with the reviewed form, which the recorded change gives.
    word = fakeDoc([
      "Title",
      { text: "A useful materialsubstrate.", reviewed: "A useful substrate." },
    ]);
    const next = await w.edit({ ...ARGS, find: "useful", replace: "handy" });
    assert.equal(next.ok, true, JSON.stringify(next));
    const c = (await w.ask("review_load")).revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [2, 0]);
  } finally {
    await w.cleanup();
  }
});

test("together: an approved exception is compared too", async () => {
  const w = await world();
  try {
    await w.ask("review_refresh", { snapshot: snap(["Title", "A novel material."]) });
    fakeDoc(["Title", "A novel material."]);
    const blocked = await w.edit();
    assert.deepEqual(blocked.blocked_by, ["L1"]);
    // The author rewrites the paragraph, then allows the exception.
    let word = fakeDoc(["Title", "A novel and costly material."]);
    let r = await w.ask("review_override", { ref: "R1.1", blocked_id: blocked.blocked_id });
    assert.deepEqual([r.ok, r.applied, r.result.code], [true, false, "StaleTarget"]);
    assert.deepEqual(untouched(word), [0, 0, "Off"]);
    assert.deepEqual(
      r.revision.comments[0].blocked.map((b) => b.state),
      ["blocked"],
    );
    assert.equal(
      "expected_before_hash" in w.calls.at(-1),
      true,
      "the stored edit has no hash of its own",
    );
    // With the paragraph as the daemon has it, the exception applies.
    word = fakeDoc(["Title", "A novel material."]);
    r = await w.ask("review_override", { ref: "R1.1", blocked_id: blocked.blocked_id });
    assert.deepEqual([r.ok, r.applied, word.edits.length], [true, true, 1]);
  } finally {
    await w.cleanup();
  }
});
