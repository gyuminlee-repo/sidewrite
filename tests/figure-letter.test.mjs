// Figure lane, step 2: an adopted figure change in the response letter. The
// checks count its requests as answered, the draft prompt reports it, the
// letter prints one figure line, and an answer written for another set of
// adopted figures is marked stale.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { checkResponses } from "../daemon/response-check.mjs";
import { buildResponsePrompt } from "../daemon/response-prompt.mjs";
import { buildLetterModel, figureLine } from "../daemon/response-model.mjs";
import { adoptedFigures, markStale, staleReasons } from "../daemon/response-quotes.mjs";
import { createResponseService, figureFileRows, refreshed } from "../daemon/response-service.mjs";
import { withDefaults, LETTER_DEFAULTS } from "../daemon/journal-profile.mjs";
import {
  RevisionSchema,
  emptyRevision,
  loadRevision,
  saveRevision,
  withRevision,
  addFigureChange,
  setFigureState,
  setResponse,
} from "../daemon/revision-store.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
const PROFILE = withDefaults({ colors: { response_body: "C00000", response_revised: "0070C0" } });
const SNAPSHOT = {
  paragraphs: [{ index: 0, text: "The growth of the strain was followed for four days." }],
};
const sha = (text) => createHash("sha256").update(text).digest("hex");
const AFTER = "the new figure file";
const fileRef = (path, text) => ({
  path,
  sha256: sha(text),
  bytes: text.length,
  mtime: "2026-01-01T00:00:00.000Z",
});
const COPY = {
  path: "figures/submission/fig2_after.pdf",
  sha256: sha(AFTER),
  at: "2026-02-01T00:00:00.000Z",
};

// One comment whose only request can be answered by a figure alone, with an
// answer text. state: of its figure change (null: none). early: the answer
// was written before the figure change was decided (else after it, with the
// Response tab open: a refresh, then the text).
function revisionWith(
  state,
  { kind = "revised", text = "We added the control curve.", early = false } = {},
) {
  let rev = RevisionSchema.parse({
    ...emptyRevision("paper.docx"),
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Figure 2 has no control.",
        requirements: [{ id: "q1", text: "add the control curve to Figure 2" }],
        requirements_outside: [
          { id: "q1", why: "A picture file.", at: "2026-01-01T00:00:00.000Z" },
        ],
      },
    ],
  });
  const write = (r) =>
    text ? setResponse(refreshed(r, SNAPSHOT, [], PROFILE), "R1.1", { text, source: "author" }) : r;
  if (early || !state) rev = write(rev);
  if (!state) return rev;
  rev = addFigureChange(rev, "R1.1", {
    figure: "Figure 2",
    kind,
    requirement_ids: ["q1"],
    summary: "A control curve was added and the time axis gives hours.",
    before: kind === "new" ? null : { source: fileRef("figures/fig2_before.pdf", "old") },
    after: { source: fileRef("figures/fig2_after.pdf", AFTER) },
  }).revision;
  if (state === "adopted")
    rev = setFigureState(rev, "R1.1", "R1.1-f1", "adopted", { adopted: COPY });
  if (state === "rejected") rev = setFigureState(rev, "R1.1", "R1.1-f1", "rejected");
  return early ? rev : write(rev);
}
const rows = (revision) => checkResponses({ revision, snapshot: SNAPSHOT, profile: PROFILE });
const ids = (list, id) => list.filter((r) => r.id === id);

// ---- checks ------------------------------------------------------------------

test("a request answered only by an adopted figure change is answered: no requirements row", () => {
  const out = rows(revisionWith("adopted"));
  assert.deepEqual(ids(out, "requirements"), []);
  assert.deepEqual(ids(out, "undecided_figures"), []);
  assert.deepEqual(
    out.filter((r) => r.level === "fail"),
    [],
  );
});

test("the same request with a proposed figure change still fails, and a warning says what to do", () => {
  const out = rows(revisionWith("proposed"));
  assert.deepEqual(ids(out, "requirements"), [
    {
      id: "requirements",
      level: "fail",
      ref: "R1.1",
      detail:
        "q1 has only a figure change that is not adopted yet: add the control curve to Figure 2",
    },
  ]);
  assert.deepEqual(ids(out, "undecided_figures"), [
    {
      id: "undecided_figures",
      level: "warn",
      ref: "R1.1",
      detail:
        "1 figure change(s) not decided: adopt or reject on the figures page; only adopted figure changes are cited",
    },
  ]);
});

test("without a figure change, or with a rejected one, the request is neither cited nor declined", () => {
  for (const state of [null, "rejected"]) {
    const out = rows(revisionWith(state));
    assert.equal(ids(out, "requirements").length, 1, String(state));
    assert.match(ids(out, "requirements")[0].detail, /^q1 is neither cited nor declined/);
    assert.deepEqual(ids(out, "undecided_figures"), []);
  }
});

test("an adopted figure change does not answer the comment by itself: the answer text is still asked for", () => {
  const out = rows(revisionWith("adopted", { text: "" }));
  assert.deepEqual(
    ids(out, "answered").map((r) => r.detail),
    ["no response and not declined"],
  );
});

// ---- the adopted copy on disk --------------------------------------------------

test("figure_file: a fail row when the adopted copy is missing or is not the file that was adopted", async () => {
  const dir = await mkdtemp(join(tmpdir(), "figure-letter-"));
  try {
    const doc = join(dir, "paper.docx");
    const rev = revisionWith("adopted");
    const missing = await figureFileRows(rev, doc);
    assert.equal(missing.length, 1);
    assert.deepEqual(
      [missing[0].id, missing[0].level, missing[0].ref],
      ["figure_file", "fail", "R1.1"],
    );
    assert.match(
      missing[0].detail,
      /the adopted copy of Figure 2 is missing: figures\/submission\/fig2_after\.pdf/,
    );

    await mkdir(join(dir, "figures", "submission"), { recursive: true });
    await writeFile(join(dir, COPY.path), "saved again in the submission folder");
    const changed = await figureFileRows(rev, doc);
    assert.equal(changed.length, 1);
    assert.match(changed[0].detail, /is not the file that was adopted/);

    await writeFile(join(dir, COPY.path), AFTER);
    assert.deepEqual(await figureFileRows(rev, doc), []);
    // Only adopted ones are looked at.
    assert.deepEqual(await figureFileRows(revisionWith("proposed"), join(dir, "x", "p.docx")), []);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("response_check and response_export see the row: a missing copy holds the letter back", async () => {
  const dir = await mkdtemp(join(tmpdir(), "figure-letter-"));
  try {
    const doc = join(dir, "paper.docx");
    const revPath = join(dir, "paper.revision.json");
    await saveRevision(revPath, { ...revisionWith("adopted"), manuscript: doc });
    assert.deepEqual(revisionWith("adopted").comments[0].response.figure_ids, ["R1.1-f1"]);
    const waiters = new Map();
    const svc = createResponseService({
      getContext: () => ({ activeDoc: doc }),
      sendEvent: (key, obj) => obj.event === "job_done" && waiters.get(obj.job_id)?.(obj),
      runPy: async () => {
        throw new Error("no Python tool is run while the letter is held");
      },
      log: () => {},
      env: ENV,
    });
    const ask = (type, payload = {}) =>
      new Promise((resolve) => {
        const reply = (r) =>
          r.accepted ? waiters.set(r.job_id, (e) => resolve(e.result)) : resolve(r);
        svc.handlers[type]({ type, request_id: "x", ...payload }, reply, "k", "word");
      });
    const withDoc = { snapshot: SNAPSHOT, headings: [] };

    let r = await ask("response_check", withDoc);
    assert.deepEqual(
      r.checks.filter((x) => x.level === "fail").map((x) => x.id),
      ["figure_file"],
    );
    assert.equal(r.summary.fail, 1);
    r = await ask("response_export", { ...withDoc, format: "docx" });
    assert.equal(r.exported, false);

    // With the copy in place nothing fails any more.
    await mkdir(join(dir, "figures", "submission"), { recursive: true });
    await writeFile(join(dir, COPY.path), AFTER);
    r = await ask("response_check", withDoc);
    assert.equal(r.summary.fail, 0, JSON.stringify(r.checks));
    // The refresh saved the adopted figure with the answer.
    const saved = await loadRevision(revPath, doc);
    assert.deepEqual(saved.comments[0].response.figure_ids, ["R1.1-f1"]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("a figure change adopted while the model writes the draft: the draft is kept and marked", async () => {
  const dir = await mkdtemp(join(tmpdir(), "figure-letter-"));
  try {
    const doc = join(dir, "paper.docx");
    const revPath = join(dir, "paper.revision.json");
    await saveRevision(revPath, { ...revisionWith("proposed", { text: "" }), manuscript: doc });
    const waiters = new Map();
    const prompts = [];
    let release = null;
    const svc = createResponseService({
      getContext: () => ({ activeDoc: doc }),
      oneshotJob: async () => ({
        model: "fake:model",
        run: (req) => {
          prompts.push(req.prompt);
          return new Promise((resolve) => (release = resolve));
        },
      }),
      sendEvent: (key, obj) => obj.event === "job_done" && waiters.get(obj.job_id)?.(obj),
      log: () => {},
      env: ENV,
    });
    const draft = () =>
      new Promise((resolve) => {
        const reply = (r) =>
          r.accepted ? waiters.set(r.job_id, (e) => resolve(e.result)) : resolve(r);
        svc.handlers.response_draft(
          {
            type: "response_draft",
            request_id: "x",
            ref: "R1.1",
            snapshot: SNAPSHOT,
            headings: [],
          },
          reply,
          "k",
          "word",
        );
      });
    const writing = async () => {
      while (!release) await new Promise((r) => setTimeout(r, 5));
      const go = release;
      release = null;
      return go;
    };

    // The prompt says the figure change is not adopted; it is adopted before
    // the model answers.
    let pending = draft();
    let go = await writing();
    assert.match(prompts[0], /Figure changes that are not adopted/);
    await withRevision(revPath, doc, (rev) =>
      setFigureState(rev, "R1.1", "R1.1-f1", "adopted", { adopted: COPY }),
    );
    go({ ok: true, json: { response: "No figure was changed." } });
    let r = await pending;
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.match(r.warnings.join(" "), /changed while this answer was drafted/);
    assert.deepEqual(r.revision.comments[0].response.stale, ["change_state"]);
    assert.equal(r.revision.comments[0].response.text, "No figure was changed.");

    // Drafted again with the figure adopted: the prompt reports it, no mark.
    pending = draft();
    go = await writing();
    assert.match(prompts[1], /Figure changes adopted for this comment/);
    go({ ok: true, json: { response: "We added the control curve to Figure 2." } });
    r = await pending;
    assert.deepEqual(r.warnings, []);
    assert.deepEqual(r.revision.comments[0].response.stale, []);
    assert.deepEqual(r.revision.comments[0].response.figure_ids, ["R1.1-f1"]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

// ---- the letter ----------------------------------------------------------------

test("the letter holds exactly one figure paragraph for an adopted figure change, in the revised colour", () => {
  const model = buildLetterModel(revisionWith("adopted"), PROFILE);
  const figures = model.paragraphs.filter((p) => p.kind === "figure");
  assert.deepEqual(figures, [
    {
      kind: "figure",
      text: "[Revised Figure 2, file fig2_after.pdf]",
      color: "0070C0",
      ref: "R1.1",
      reviewer: 1,
    },
  ]);
  // After the answer, and the summary of the author is not printed by the tool.
  assert.deepEqual(
    model.paragraphs.map((p) => p.kind),
    ["reviewer_head", "comment_head", "comment", "response", "figure"],
  );
  assert.equal(
    model.paragraphs.some((p) => p.text.includes("time axis gives hours")),
    false,
  );
});

test("no figure paragraph for a proposed or rejected figure change; a new figure has its own frame", () => {
  for (const state of ["proposed", "rejected", null]) {
    const model = buildLetterModel(revisionWith(state), PROFILE);
    assert.equal(model.paragraphs.filter((p) => p.kind === "figure").length, 0, String(state));
  }
  const fresh = buildLetterModel(revisionWith("adopted", { kind: "new" }), PROFILE);
  assert.deepEqual(
    fresh.paragraphs.filter((p) => p.kind === "figure").map((p) => p.text),
    ["[New Figure 2, file fig2_after.pdf]"],
  );
});

test("the figure frames come from the profile, with defaults", () => {
  assert.equal(LETTER_DEFAULTS.figure_line, "[Revised {figure}, file {file}]");
  assert.equal(LETTER_DEFAULTS.figure_new_line, "[New {figure}, file {file}]");
  const f = adoptedFigures(revisionWith("adopted").comments[0])[0];
  const L = withDefaults({
    response_letter: { figure_line: "[{figure} was replaced: {file}]" },
  }).response_letter;
  assert.equal(figureLine(f, L), "[Figure 2 was replaced: fig2_after.pdf]");
  // A label with "$" or a frame key in it is printed as it is.
  assert.equal(
    figureLine({ ...f, figure: "Figure $& {file}" }, LETTER_DEFAULTS),
    "[Revised Figure $& {file}, file fig2_after.pdf]",
  );
  // A declined comment with an adopted figure prints the figure line, not "[No change made]".
  const rev = revisionWith("adopted");
  rev.comments[0].declined = { reason: "The text needs no change." };
  const kinds = buildLetterModel(rev, PROFILE).paragraphs.map((p) => p.kind);
  assert.equal(kinds.includes("declined"), false);
  assert.equal(kinds.filter((k) => k === "figure").length, 1);
});

// ---- the draft prompt ----------------------------------------------------------

const prompt = (rev) =>
  buildResponsePrompt({ comment: rev.comments[0], profile: PROFILE, ledger: [] });

test("the draft prompt lists an adopted figure change with its label, summary and request ids, and no file", () => {
  const built = prompt(revisionWith("adopted"));
  assert.match(
    built.prompt,
    /Figure changes adopted for this comment \(figure files replaced or added by the author; you may report them\):\n- \[R1\.1-f1\] Figure 2 \(revised figure\) \(answers q1\)\n {2}What changed, in the words of the author: <<<A control curve was added and the time axis gives hours\.>>>/,
  );
  assert.deepEqual(built.figureIds, ["R1.1-f1"]);
  assert.equal(/fig2_after|submission|[0-9a-f]{64}/.test(built.prompt), false);
  assert.equal(built.prompt.includes("Figure changes that are not adopted"), false);
  // The two rules about claims allow the figure; the tool adds the figure line.
  assert.match(
    built.prompt,
    /3\. Never claim a change that is not listed as in the manuscript or as an adopted figure change\./,
  );
  assert.match(
    built.prompt,
    /4\. If no change is listed as in the manuscript and no figure change as adopted, do not say that anything was revised\./,
  );
  assert.match(built.prompt, /Do not write file names\. The tool adds the figure line/);
});

test("a figure change that is not adopted is listed under do not report", () => {
  for (const [state, why] of [
    ["proposed", "not decided by the author yet"],
    ["rejected", "rejected by the author"],
  ]) {
    const built = prompt(revisionWith(state));
    assert.equal(built.prompt.includes("Figure changes adopted for this comment"), false);
    assert.match(
      built.prompt,
      new RegExp(
        `Figure changes that are not adopted \\(do not report any of these\\):\\n- \\[R1\\.1-f1\\] Figure 2 \\(revised figure\\) \\(answers q1\\): ${why}`,
      ),
    );
    assert.deepEqual(built.figureIds, []);
    // What the author said changed is not given for a change that must not be reported.
    assert.equal(built.prompt.includes("time axis gives hours"), false);
  }
});

test("a comment without figure changes gets the prompt it got before", () => {
  const plain = prompt(revisionWith(null));
  assert.equal(/[Ff]igure change/.test(plain.prompt), false);
  assert.match(plain.prompt, /3\. Never claim a change that is not listed as in the manuscript\./);
  assert.match(
    plain.prompt,
    /4\. If no change is listed as in the manuscript, do not say that anything was revised\./,
  );
  assert.deepEqual(plain.figureIds, []);
  // A comment of a file from before the figure lane (no figure_changes key).
  const old = structuredClone(revisionWith(null).comments[0]);
  delete old.figure_changes;
  assert.equal(
    buildResponsePrompt({ comment: old, profile: PROFILE, ledger: [] }).prompt,
    plain.prompt,
  );
});

// ---- staleness -----------------------------------------------------------------

test("an answer written before a figure change was adopted is marked change_state at the next refresh", () => {
  // The answer stands, the figure change waits: nothing is stale.
  let rev = refreshed(revisionWith("proposed", { early: true }), SNAPSHOT, [], PROFILE);
  assert.deepEqual(rev.comments[0].response.stale, []);
  assert.deepEqual(rev.comments[0].response.figure_ids, []);

  // It is adopted on the figures page; the next refresh marks the answer and
  // takes the new set as the saved one.
  rev = setFigureState(rev, "R1.1", "R1.1-f1", "adopted", { adopted: COPY });
  assert.deepEqual(staleReasons(rev.comments[0], SNAPSHOT, null), ["change_state"]);
  rev = refreshed(rev, SNAPSHOT, [], PROFILE);
  assert.deepEqual(rev.comments[0].response.stale, ["change_state"]);
  assert.deepEqual(rev.comments[0].response.figure_ids, ["R1.1-f1"]);
  const stale = ids(rows(rev), "stale");
  assert.equal(stale.length, 1);
  assert.equal(stale[0].level, "fail");
  assert.match(stale[0].detail, /^change_state: the accepted changes or adopted figure changes/);

  // The mark stays over further refreshes until the answer is written again.
  rev = refreshed(rev, SNAPSHOT, [], PROFILE);
  assert.deepEqual(rev.comments[0].response.stale, ["change_state"]);

  // Taking the adoption back marks an answer that reported it.
  let back = structuredClone(rev);
  back.comments[0].response.stale = [];
  back = setFigureState(back, "R1.1", "R1.1-f1", "proposed");
  assert.deepEqual(markStale(back, SNAPSHOT).comments[0].response.stale, ["change_state"]);
});

test("no mark without an answer text, for a file from before the figure lane, or for the citations a pane showed", () => {
  // No answer yet: the adopted set is saved with the (empty) response, so the
  // answer that is written next is not marked.
  let rev = refreshed(revisionWith("adopted", { text: "" }), SNAPSHOT, [], PROFILE);
  assert.deepEqual(rev.comments[0].response.figure_ids, ["R1.1-f1"]);
  rev = setResponse(rev, "R1.1", { text: "We added the control curve.", source: "author" });
  rev = refreshed(rev, SNAPSHOT, [], PROFILE);
  assert.deepEqual(rev.comments[0].response.stale, []);

  // An answer of an old file: no figure changes, no figure_ids, nothing stale.
  const old = revisionWith(null);
  assert.deepEqual(markStale(old, SNAPSHOT).comments[0].response.stale, []);

  // response_save judges the citations the card showed with a response that
  // has no figure_ids: figures are not judged there.
  const adopted = revisionWith("adopted").comments[0];
  assert.deepEqual(
    staleReasons(
      { changes: [], figure_changes: adopted.figure_changes, response: { citations: [] } },
      SNAPSHOT,
      null,
    ),
    [],
  );
});
