// The cite stage of the acceptance runner (scripts/e2e/run-acceptance.mjs
// mainCite): do tracked edits leave reference-manager fields whole? Word is
// not here. The fixture and the checker are run as files; the edits of the
// run go through the review service with a Word that answers what it is
// told to; the judges get recorded inputs, written by hand in the form the
// pane driver, word-com.ps1 and marked-diff-word.ps1 return (none was
// recorded from Word before the first run of the stage).
import { test } from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { createReviewService } from "../daemon/review-service.mjs";
import { addBlocked, dropBlocked, loadRevision, saveRevision } from "../daemon/revision-store.mjs";
import { splitReviewRegex } from "../daemon/review-split.mjs";
import { paragraphHash } from "../daemon/doc-diff.mjs";
import { missing } from "./fixtures/py-packages.mjs";
import * as J from "../scripts/e2e/judge.mjs";

const here = (p) => fileURLToPath(new URL(p, import.meta.url));
const FIXTURE = here("./fixtures/citation_manuscript.docx");
const BUILDER = here("../scripts/e2e/make_citation_fixture.py");
const CHECKER = here("../tools/citation_fields_check.py");
const python = (args) => execFileSync("python3", args, { encoding: "utf8", maxBuffer: 64 << 20 });
const SPEC = JSON.parse(python([BUILDER, "--cases"]));
const TEXTS = JSON.parse(python([BUILDER, "--dump", FIXTURE]));
const REVIEW = readFileSync(here("./fixtures/cite_review.txt"), "utf8");

// The checker on two files: its exit code, its lines, its JSON line.
function check(before, after, extra = []) {
  let exit = 0;
  let out;
  try {
    // --detail, as the runner asks for it: the list of fields with their states.
    out = python([CHECKER, "--before", before, "--after", after, "--detail", ...extra]);
  } catch (e) {
    exit = e.status;
    out = String(e.stdout);
  }
  const lines = out.trim().split("\n");
  return { exit, result: JSON.parse(lines.pop()), lines };
}

test("the committed fixture holds the fields of both plugins, whole", () => {
  const { exit, result } = check(FIXTURE, FIXTURE);
  assert.equal(exit, 0);
  assert.deepEqual(result.before.by_kind, {
    "EN.CITE": 2,
    "EN.CITE.DATA": 1,
    "EN.REFLIST": 1,
    REF: 1,
    ZOTERO_BIBL: 1,
    ZOTERO_ITEM: 11,
    hyperlink: 1,
  });
  assert.deepEqual(result.before.by_state, { live: 18 });
  assert.deepEqual(
    [result.fld_data.before, result.properties.before, result.doc_vars.before],
    [2, 2, 3],
  );
  assert.equal(
    J.fieldsLine(result),
    "fields 18 -> 18 (live 18, in deletions 0, in insertions 0, split 0)",
  );
});

test("a fresh build of the fixture has the same fields", { skip: missing("docx") }, async () => {
  const dir = await mkdtemp(join(tmpdir(), "cite-fixture-"));
  try {
    const out = join(dir, "built.docx");
    python([BUILDER, out]);
    assert.deepEqual(check(FIXTURE, out).result.failures, []);
  } finally {
    await rm(dir, { recursive: true });
  }
});

test("every text the run edits is in the document once, and the citation texts are where the cases need them", () => {
  const whole = TEXTS.join("\n");
  for (const c of SPEC.cases) assert.equal(whole.split(c.find).length, 2, c.name);
  assert.deepEqual(
    SPEC.cases.map((c) => c.name),
    ["next-to", "spans", "straddles", "delete-sentence", "formatted"],
  );
  const by = Object.fromEntries(SPEC.cases.map((c) => [c.name, c]));
  // next-to: the citation is in the sentence, not in the find.
  assert.ok(TEXTS.some((t) => t.includes("(Arden, 2019) and are cheap to run.")));
  assert.ok(!/\(/.test(by["next-to"].find));
  // spans: the whole visible text of the citation is inside, with words on both sides.
  assert.match(by.spans.find, /\w \(Belloc, 2020\) \w/);
  assert.ok(by.spans.replace.includes("(Belloc, 2020)"));
  // straddles: it ends inside the visible text of the citation.
  assert.ok(by.straddles.find.endsWith("(Corvane,"));
  // delete-sentence: a whole sentence with an EndNote citation.
  assert.equal(by["delete-sentence"].replace, "");
  assert.ok(by["delete-sentence"].find.includes("(Garrowby, 2017)."));
});

test("the review of the stage splits into its four comments without a model", () => {
  const refs = splitReviewRegex(REVIEW).map((c) => c.ref);
  assert.deepEqual(
    refs.filter((r) => /^R/.test(r)),
    J.CITE_REFS,
  );
  assert.ok(J.CITE_TURN_REFS.every((r) => J.CITE_REFS.includes(r)));
  assert.ok(!J.CITE_TURN_REFS.includes(J.CITE_SEED_REF));
  const split = { message: "Split into 5 comments (found by the numbering).", draft_refs: refs };
  assert.equal(J.judge("cite-split", split).pass, true);
  assert.equal(
    J.judge("cite-split", { ...split, draft_refs: ["R1.1", "R1.2", "R1.3"] }).pass,
    false,
  );
  const cards = { comments: J.CITE_REFS.map((ref) => ({ ref })) };
  assert.equal(J.judge("cite-draft", cards).pass, true);
  assert.equal(J.judge("cite-draft", { comments: cards.comments.slice(1) }).pass, false);
});

// The review service with the document of the fixture and a Word that
// answers `word(name, args)`.
async function service(word) {
  const dir = await mkdtemp(join(tmpdir(), "cite-stage-"));
  const doc = join(dir, "paper.docx");
  const lines = [];
  const calls = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash", "tracked_rewrite"] }),
    callTool: async (name, args) => {
      calls.push([name, args]);
      return word(name, args);
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
  await ask("review_save_comments", {
    comments: splitReviewRegex(REVIEW)
      .filter((c) => /^R/.test(c.ref))
      .map((c) => ({ ref: c.ref, reviewer: 1, text: c.text })),
  });
  await ask("review_refresh", {
    snapshot: {
      paragraphs: TEXTS.map((text, index) => ({
        index,
        id: null,
        text,
        reviewed_text: text,
        w_ins: 0,
        w_del: 0,
      })),
      comments: [],
    },
  });
  return { doc, file, ask, lines, calls, cleanup: () => rm(dir, { recursive: true }) };
}
const WORD0 = { ok: true, revisions: [] };
const WORD1 = { ok: true, revisions: [{ type: 2, text: "x", start: 1, end: 2 }] };

test("each edit of the run passes the guards of the daemon and reaches Word with the hash of its paragraph", async () => {
  const t = await service((_name, args) => ({
    ok: true,
    word_comment_id: 9,
    paragraph_index: 3,
    before: "b",
    after: "a",
    narrowed: !!args.narrow,
    target_checked: true,
  }));
  try {
    for (const c of SPEC.cases) {
      const seeded = addBlocked(await loadRevision(t.file, t.doc), J.CITE_SEED_REF, {
        args: J.citeEdit(c),
        blocked_by: [J.FORBID_GUARD_SEED_RULE],
      });
      await saveRevision(t.file, seeded.revision);
      const at = t.lines.length;
      const sent = t.calls.length;
      const r = await t.ask("review_override", {
        ref: J.CITE_SEED_REF,
        blocked_id: seeded.blocked.id,
      });
      assert.equal(r.applied, true, `${c.name}: ${r.result?.message}`);
      // The answer brings the file without the item: the pane takes it
      // (review-pane.js ask) and the row leaves the card, which is what the
      // pane step apply-seeded waits for.
      assert.deepEqual(r.revision.comments.find((x) => x.ref === J.CITE_SEED_REF).blocked, []);
      assert.equal(t.calls.length, sent + 1, c.name);
      const [name, args] = t.calls[sent];
      assert.equal(name, "office_tracked_edit");
      assert.equal(args.find, c.find);
      const target = TEXTS.find((x) => x.includes(c.find));
      assert.equal(args.expected_before_hash, paragraphHash(target), c.name);
      assert.match(args.comment_text, /^\[R1\.4\] Acceptance run: /);
      // The pane shows nothing after an edit that went in (step: the view).
      const outcome = J.citeOutcome({
        step: { ok: true, message: "", message_error: false, row_gone: true },
        log: t.lines.slice(at).join("\n"),
        before: seeded.revision,
        revision: JSON.parse(await readFile(t.file, "utf8")),
        word: WORD1,
        wordBefore: WORD0,
        seedId: seeded.blocked.id,
        ref: J.CITE_SEED_REF,
      });
      assert.equal(outcome.answer, "applied", c.name);
      assert.equal(outcome.changes.length, 1);
      assert.equal(outcome.changes[0].kind, c.replace === "" ? "delete" : "replace");
      assert.deepEqual(outcome.word_revisions, [0, 1]);
    }
    // The edit that spans the citation is narrowed to the words that differ,
    // and those still hold the citation: Word gets a range over the field.
    const spans = t.calls.find(([, a]) => a.find === SPEC.cases[1].find)[1];
    assert.equal(spans.narrow.find, "slow growth (Belloc, 2020) in narrow");
    const straddles = t.calls.find(([, a]) => a.find === SPEC.cases[2].find)[1];
    assert.equal(straddles.narrow, undefined);
  } finally {
    await t.cleanup();
  }
});

test("an edit Word refuses is read as refused with its code, and the item of the run goes off the file", async () => {
  const c = SPEC.cases[1];
  const t = await service(() => ({
    ok: false,
    code: "NotFound",
    message: `"${c.find}" not found`,
  }));
  try {
    const seeded = addBlocked(await loadRevision(t.file, t.doc), J.CITE_SEED_REF, {
      args: J.citeEdit(c),
      blocked_by: [J.FORBID_GUARD_SEED_RULE],
    });
    await saveRevision(t.file, seeded.revision);
    const r = await t.ask("review_override", {
      ref: J.CITE_SEED_REF,
      blocked_id: seeded.blocked.id,
    });
    assert.equal(r.applied, false);
    const revision = JSON.parse(await readFile(t.file, "utf8"));
    const args = {
      // review-pane.js override: the message of the result, shown as an error.
      step: { ok: true, message: r.result.message, message_error: true, row_gone: false },
      log: t.lines.join("\n"),
      before: seeded.revision,
      revision,
      word: WORD0,
      wordBefore: WORD0,
      seedId: seeded.blocked.id,
      ref: J.CITE_SEED_REF,
    };
    const outcome = J.citeOutcome(args);
    assert.deepEqual(
      [outcome.answer, outcome.code, outcome.word_unchanged],
      ["refused", "NotFound", true],
    );
    // A refusal that left a tracked change behind is not "unchanged".
    assert.equal(J.citeOutcome({ ...args, word: WORD1 }).word_unchanged, false);
    // No answer from the pane step: unknown, never applied or refused.
    assert.equal(
      J.citeOutcome({ ...args, step: { ok: false, error: "Timeout 180000ms exceeded." } }).answer,
      "unknown",
    );
    const cleaned = dropBlocked(
      await loadRevision(t.file, t.doc),
      J.CITE_SEED_REF,
      seeded.blocked.id,
    );
    assert.ok(cleaned, "the refused item is back on file until the run drops it");
    assert.deepEqual(cleaned.comments.find((x) => x.ref === J.CITE_SEED_REF).blocked, []);
  } finally {
    await t.cleanup();
  }
});

test("refusalCode: the messages of the two tools", () => {
  const code = (m) => J.refusalCode(m);
  assert.equal(code('"a b" not found'), "NotFound");
  assert.equal(code('"a" occurs 2 times; pass paragraph_hint or a longer find'), "Ambiguous");
  assert.equal(code("find is not in the paragraph text"), "AmbiguousInParagraph");
  assert.equal(
    code("Not applied: the paragraph changed in Word after it was last read, so the edit ..."),
    "StaleTarget",
  );
  assert.equal(
    code('Not applied: the paragraph of "x" could not be told from the last status refresh, ...'),
    "StaleTarget",
  );
  assert.equal(
    code(
      "The paragraph holds a field (a citation or cross-reference, for one), so it is not rewritten as a whole. Nothing was changed.",
    ),
    "ParagraphNotPlain",
  );
  assert.equal(
    code("Not applied: the change puts a number into the manuscript (7) that ..."),
    "NumberNeedsAuthor",
  );
  assert.equal(code("something else"), "Refused");
});

// Checker results as the runner keeps them, made from real files: the
// fixture, and copies of it broken by hand the way a tool could break them.
const CHECKS = await (async () => {
  const dir = await mkdtemp(join(tmpdir(), "cite-files-"));
  // A copy of the fixture whose word/document.xml went through one edit.
  const copy = (name, body) => {
    const out = join(dir, name);
    python([
      "-c",
      [
        "import sys, zipfile",
        "src, out = sys.argv[1], sys.argv[2]",
        "END = '<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>'",
        "BEGIN = '<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>'",
        'DEL = \'<w:del w:id="900" w:author="Hand" w:date="2026-01-01T00:00:00Z">\'',
        "def gone(x): return DEL + x.replace('w:instrText', 'w:delInstrText').replace('<w:t ', '<w:delText ').replace('</w:t>', '</w:delText>') + '</w:del>'",
        "zin = zipfile.ZipFile(src)",
        "x = zin.read('word/document.xml').decode()",
        "a = x.index(BEGIN); c = x.index(END, a) + len(END); s = x.index('separate', a)",
        body,
        "z = zipfile.ZipFile(out, 'w')",
        "[z.writestr(i, y.encode() if i.filename == 'word/document.xml' else zin.read(i.filename)) for i in zin.infolist()]",
        "z.close()",
      ].join("\n"),
      FIXTURE,
      out,
    ]);
    return out;
  };
  const files = {
    // The first citation lies whole in a tracked deletion.
    tracked: copy("tracked.docx", "y = x[:a] + gone(x[a:c]) + x[c:]"),
    // ... and after accept-all it is gone.
    accepted: copy("accepted.docx", "y = x[:a] + x[c:]"),
    // The begin marker and the instruction in a deletion, the rest outside.
    half: copy("half.docx", "b = x.rindex('<w:r>', a, s); y = x[:a] + gone(x[a:b]) + x[b:]"),
    // After accept-all of that: a separate and an end without a begin.
    halfAccepted: copy("half-accepted.docx", "b = x.rindex('<w:r>', a, s); y = x[:a] + x[b:]"),
  };
  const out = {
    same: check(FIXTURE, FIXTURE),
    tracked: check(FIXTURE, files.tracked),
    accepted: check(FIXTURE, files.accepted, ["--expect-removed", "1"]),
    acceptedNoneExpected: check(FIXTURE, files.accepted, ["--expect-removed", "0"]),
    half: check(FIXTURE, files.half),
    halfAccepted: check(FIXTURE, files.halfAccepted, ["--expect-removed", "0"]),
  };
  await rm(dir, { recursive: true });
  return out;
})();

test("the checker results the judges are given: whole in a deletion passes, half in one fails with the markers", () => {
  assert.deepEqual([CHECKS.same.exit, CHECKS.tracked.exit, CHECKS.accepted.exit], [0, 0, 0]);
  assert.deepEqual(CHECKS.tracked.result.after.by_state, { deleted: 1, live: 17 });
  assert.deepEqual([CHECKS.half.exit, CHECKS.half.result.failures], [1, ["split-field"]]);
  assert.deepEqual(CHECKS.half.lines, [
    'FAIL split-field | word/document.xml p3 "Shaken flasks remain the usual first ves" | kind=ZOTERO_ITEM | before: begin@- instr@- separate@- result@- end@- | after: begin@del instr@del separate@- result@- end@-',
  ]);
  assert.equal(CHECKS.halfAccepted.exit, 1);
  assert.ok(CHECKS.halfAccepted.result.failures.includes("unbalanced-after"));
  assert.ok(
    CHECKS.halfAccepted.lines.some((l) => /separate without a begin .*p3 "Shaken flasks/.test(l)),
  );
});

const applied = {
  answer: "applied",
  changes: [{ kind: "replace", narrowed: true }],
  word_revisions: [0, 2],
};
const refused = { answer: "refused", code: "NotFound", word_unchanged: true, changes: [] };

test("the row of an edit: the answer of the tool is told, the fields decide", () => {
  const row = (outcome, c) =>
    J.citeCaseRow({ name: "spans", what: "a phrase", outcome, check: c?.result ?? c });
  const ok = row(applied, CHECKS.tracked);
  assert.equal(ok.result, "PASS");
  assert.equal(
    ok.got,
    "tool: applied (replace, narrowed) | fields 18 -> 18 (live 17, in deletions 1, in insertions 0, split 0)",
  );
  // A refusal is as good as an applied edit when the fields hold.
  assert.equal(row(refused, CHECKS.same).result, "PASS");
  assert.match(row(refused, CHECKS.same).got, /^tool: refused NotFound \| /);
  // Applied, and a field is left half in a deletion.
  const bad = row(applied, CHECKS.half);
  assert.equal(bad.result, "FAIL");
  assert.match(
    bad.got,
    /FAILED split-field; first: ZOTERO_ITEM in "Shaken flasks remain the usual first ves" begin@del instr@del separate@- result@- end@-/,
  );
  // Refused, but the document has another tracked change.
  assert.equal(row({ ...refused, word_unchanged: false }, CHECKS.same).result, "FAIL");
  // No readable answer, no saved file.
  assert.equal(
    row({ answer: "unknown", message: "timeout", changes: [] }, CHECKS.same).result,
    "FAIL",
  );
  assert.equal(row(applied, null).result, "FAIL");
  assert.match(row(applied, null).got, /no checker result/);
  // A citation deleted whole and typed again as words passes, and is told.
  const flat = {
    ...CHECKS.tracked.result,
    flattened: [{ kind: "ZOTERO_ITEM", shown: "(Belloc, 2020)", text: "Earlier invented work" }],
  };
  assert.equal(row(applied, flat).result, "PASS");
  assert.match(row(applied, flat).got, /; replaced by plain text: ZOTERO_ITEM \(Belloc, 2020\)$/);
});

// The pane's snapshot of the fixture, as the driver step returns it.
const SNAPSHOT = TEXTS.map((text, index) => ({ index, id: null, text, reviewed_text: text }));

test("the rewrite call is the first part the daemon would send, with plain_only and the hash", () => {
  const call = J.citeRewriteCall(SNAPSHOT, SPEC.rewrite);
  assert.equal(call.error, undefined);
  const at = TEXTS.findIndex((x) => x.includes(SPEC.rewrite.paragraph_has));
  assert.deepEqual(call.paragraph, { index: at, id: null });
  assert.equal(call.parts, 2);
  assert.equal(call.args.plain_only, true);
  assert.equal(call.args.paragraph_hint, `p${at}`);
  assert.equal(call.args.expected_before_hash, paragraphHash(TEXTS[at]));
  assert.match(call.args.comment_text, /^\[R1\.4\] /);
  // The last difference of the paragraph goes first.
  assert.deepEqual([call.args.find, call.args.replace], ["change", "alter"]);
  assert.equal(TEXTS[at].split(call.args.find).length, 2);
  assert.match(
    J.citeRewriteCall(
      SNAPSHOT.filter((p) => p.index !== at),
      SPEC.rewrite,
    ).error,
    /^0 paragraphs hold/,
  );
});

test("the rewrite row passes only on a ParagraphNotPlain refusal that changed nothing", () => {
  const call = J.citeRewriteCall(SNAPSHOT, SPEC.rewrite);
  const notPlain = {
    ok: true,
    result: {
      ok: false,
      code: "ParagraphNotPlain",
      paragraph_index: 8,
      message:
        "The paragraph holds a field (a citation or cross-reference, for one), so it is not rewritten as a whole. Nothing was changed.",
    },
  };
  const row = (more) =>
    J.citeRewriteRow({
      what: "w",
      call,
      step: notPlain,
      word: WORD0,
      wordBefore: WORD0,
      check: CHECKS.same.result,
      ...more,
    });
  assert.equal(row().result, "PASS");
  assert.match(
    row().got,
    /^tool: refused ParagraphNotPlain; tracked changes 0 -> 0 \| fields 18 -> 18/,
  );
  // It went in: the tool did not refuse a paragraph with fields.
  const went = { ok: true, result: { ok: true, word_comment_id: 4, narrowed: true } };
  assert.equal(row({ step: went, word: WORD1 }).result, "FAIL");
  assert.match(row({ step: went, word: WORD1 }).got, /^tool: applied /);
  // Refused for another reason: the paragraph was never judged as not plain.
  const stale = { ok: true, result: { ok: false, code: "StaleTarget", message: "x" } };
  assert.equal(row({ step: stale }).result, "FAIL");
  assert.match(row({ step: stale }).got, /refused StaleTarget/);
  assert.equal(row({ step: { ok: true, threw: "boom" } }).result, "FAIL");
  assert.equal(row({ call: { error: "0 paragraphs hold" }, step: null }).result, "FAIL");
  assert.equal(row({ check: CHECKS.half.result }).result, "FAIL");
});

// Word's answers (marked-diff-word.ps1): a paragraph's fields as
// "code U+001F result U+001E".
const wordJob = (id, fields, more = {}) => ({
  id,
  ok: true,
  error: null,
  dirty_on_open: false,
  revisions_left: 0,
  shape: { codes: ["", fields.map(([c, r]) => `${c}\u001f${r}\u001e`).join(""), ""] },
  ...more,
});
const F1 = [" ADDIN ZOTERO_ITEM CSL_CITATION {} ", "(Arden, 2019)"];
const F2 = [" ADDIN EN.CITE <EndNote/>", "(Garrowby, 2017)"];
const JOBS = {
  baseline: wordJob("baseline", [F1, F2]),
  tracked: wordJob("tracked", [F1, F2]),
  accept: wordJob("accept", [F2]),
  reject: wordJob("reject", [F1, F2]),
};

test("accept-all and reject-all: the original comes back, and only fields that lay in deletions are gone", () => {
  const good = {
    tracked: CHECKS.tracked.result,
    accepted: CHECKS.accepted.result,
    rejected: CHECKS.same.result,
    jobs: JOBS,
  };
  const e = J.citeAcceptEvidence(good);
  assert.equal(e.pass, true, JSON.stringify(e.checks));
  assert.deepEqual(e.deleted_fields, ["ZOTERO_ITEM"]);
  assert.deepEqual(e.removed_by_accept, [
    { kind: "ZOTERO_ITEM", text: "Shaken flasks remain the usual first ves" },
  ]);
  assert.deepEqual(e.word_field_counts, { original: 2, rejected: 2, accepted: 1 });
  const failed = (more) => {
    const x = J.citeAcceptEvidence({ ...good, ...more });
    assert.equal(x.pass, false);
    return Object.keys(x.checks).filter((k) => !x.checks[k]);
  };
  // Reject-all did not give the original back.
  assert.deepEqual(failed({ rejected: CHECKS.accepted.result }), ["rejected_equals_original"]);
  // Accept-all left half a field.
  assert.deepEqual(failed({ accepted: CHECKS.halfAccepted.result }), ["accepted_whole"]);
  // A field is gone after accept-all that lay in no deletion.
  assert.deepEqual(failed({ tracked: CHECKS.same.result }), ["accepted_lacks_only_the_deleted"]);
  // The checker was told to expect none gone, and one is.
  assert.deepEqual(failed({ accepted: CHECKS.acceptedNoneExpected.result }), ["accepted_whole"]);
  // Word could not open a file, left a revision, or lists other fields.
  assert.deepEqual(
    failed({ jobs: { ...JOBS, accept: { id: "accept", ok: false, error: "open: repair" } } }),
    ["word_opened_all", "no_revision_left"],
  );
  assert.deepEqual(
    failed({ jobs: { ...JOBS, reject: wordJob("reject", [F1, F2], { revisions_left: 1 }) } }),
    ["no_revision_left"],
  );
  assert.deepEqual(failed({ jobs: { ...JOBS, reject: wordJob("reject", [F2]) } }), [
    "word_lists_the_original_after_reject",
  ]);
  // No file was written at all.
  assert.equal(J.citeAcceptEvidence({ tracked: null, accepted: null, rejected: null }).pass, false);
  // Word changing a file while opening it is told, not judged.
  const dirty = J.citeAcceptEvidence({
    ...good,
    jobs: { ...JOBS, tracked: wordJob("tracked", [F1, F2], { dirty_on_open: true }) },
  });
  assert.deepEqual([dirty.pass, dirty.changed_on_open], [true, ["tracked"]]);
});

test("the plugins are found by name in what Word lists", () => {
  const e = J.pluginsEvidence({
    ok: true,
    com_addins: [
      {
        prog_id: "EndNote.CwywAddin",
        description: "EndNote Cite While You Write",
        connected: true,
      },
      { prog_id: "Other.AddIn", description: "Something", connected: false },
    ],
    templates: [{ name: "Zotero.dotm", installed: true, path: "C:\\x" }],
    errors: [],
  });
  assert.deepEqual(
    [e.read, e.zotero, e.endnote.length, e.com_addins, e.templates],
    [true, ["Zotero.dotm"], 1, 2, 1],
  );
  const none = J.pluginsEvidence({ ok: true, com_addins: [], templates: [], errors: [] });
  assert.deepEqual([none.zotero, none.endnote], [[], []]);
  assert.deepEqual(J.pluginsEvidence({ ok: false, error: "no Word" }).errors, ["no Word"]);
});

test("the cite flow uses only pane steps, judges and judge exports that exist", () => {
  const runner = readFileSync(here("../scripts/e2e/run-acceptance.mjs"), "utf8");
  const driver = readFileSync(here("../scripts/e2e/pane-driver.cjs"), "utf8");
  assert.match(runner, /cite: mainCite/);
  const flow = runner.slice(
    runner.indexOf("// ---- stage cite"),
    runner.indexOf("// Close the Word this run started"),
  );
  const steps = [...flow.matchAll(/\bpane\(\s*"([a-z-]+)"/g)].map((m) => m[1]);
  assert.ok(steps.includes("apply-seeded") && steps.includes("tracked-edit-direct"));
  for (const s of new Set(steps)) {
    assert.ok(new RegExp(`\\n  (async )?"?${s}"?\\(page`).test(driver), `no pane step ${s}`);
  }
  for (const [, name] of flow.matchAll(/judge: "([a-z-]+)"/g)) {
    assert.ok(name === "none" || J.JUDGES[name], `no judge ${name}`);
  }
  // The two markers of the stage, and the PowerShell the flow calls.
  assert.deepEqual(
    [...flow.matchAll(/\bmarker\(\s*"(cite-[a-z-]+)"/g)].map((m) => m[1]),
    ["cite-tracked", "cite-accept-reject"],
  );
  const com = readFileSync(here("../scripts/e2e/word-com.ps1"), "utf8");
  for (const [, action] of flow.matchAll(/\bcom\("([a-z-]+)"/g)) {
    assert.ok(com.includes(`\n  "${action}" {`), `no word-com.ps1 action ${action}`);
  }
  const word = readFileSync(here("../scripts/e2e/marked-diff-word.ps1"), "utf8");
  assert.match(word, /\$job\.save_as/);
  // A row that is only told is no failure of the run.
  const d = J.exitDecision({
    failed: false,
    rows: [
      { step: "a", result: "NOTE" },
      { step: "b", result: "PASS" },
    ],
    features: [],
  });
  assert.equal(d.code, 0);
});
