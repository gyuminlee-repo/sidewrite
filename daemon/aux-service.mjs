// Tools tab: task-pane requests beside the review and the response letter.
//
// Requests (bridge.mjs extraHandlers): { type, request_id, ...payload } and the
// reply { type: "<type>_result", request_id, ok, ... }.
//   tools_state, profile_rules_apply, format_check
// Jobs (AUX_JOBS; they call a model or a Python tool): accepted at once, the
// result follows as { type: "tools_event", event: "job_done", job_id, result }.
//   selection_translate, selection_ask, profile_rules_extract, marked_build,
//   anonymize, submission_check
//
// The selection requests get the selected text from the pane and answer with
// text for the pane. They read no file and no revision, and nothing of the
// answer is stored: the document stays as it is.
//
// Journal guideline. profile_rules_extract turns the pasted guideline into
// rule candidates, each with the sentence it was taken from, and stores
// nothing. profile_rules_apply takes the candidates the author picked: the
// rules go into the ledger with their sentence, a format setting goes into
// the overrides of the paper's profile. A picked rule must cite a sentence
// of the guideline text that is sent with it.
//
// Marked copy. marked_build compares the saved file with the earlier version
// the pane sends and writes "<name>_marked.docx" next to the manuscript, with
// the tool and the options the script uses (tools/build_marked_diff.py), so
// the button and the script give the same file. A refusal of the tool is the
// error of the job, word for word.
//
// Anonymized copy. anonymize writes "<name>_anonymized.docx" from the saved
// file and answers the verifier's report (anonymize.mjs). A copy that fails
// the verifier is set aside under an INCOMPLETE name and the result says so.
//
// Checks. format_check reports and changes nothing: the format rules the
// profile turns on (format-check.mjs) over the formatting the pane read, and
// the count rules of the ledger (count-check.mjs) over the snapshot and the
// headings. Nothing of the result is stored.
//
// Submission checklist. submission_check answers the rows of "Before
// submission" (submission-check.mjs) for the document as the pane read it
// just now. The record is brought up to date with that document in memory
// and is not saved; nothing of a row is stored.
//
// Nothing here writes to the open document. A copy (marked, anonymized) is a
// new file next to the manuscript, made from the file as it was last saved.
// A copy that was made and passed its check is put on record with its hash
// and the hash of the manuscript it was made from (revision.artifacts); the
// submission checklist compares both with the disk.
//
// Handler rules (as in response-service.mjs): the document of a pane is read
// once per request (at(key)); a failure is an error reply, never a silent
// default.

import { existsSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, dirname, isAbsolute, join, relative, sep } from "node:path";
import {
  setArtifact,
  localDocPath,
  revisionPathFor,
  decisionsPathFor,
  loadRevision,
  withRevision,
  importLedgerItems,
  setProfile,
} from "./revision-store.mjs";
import { wrapHandlers } from "./handler-wrap.mjs";
import { refuseInPart } from "./revision-parts.mjs";
import { effectiveProfile, mergeProfile, regexProblems } from "./journal-profile.mjs";
import { extractGuidelineRules } from "./guideline-extract.mjs";
import { splitSentences } from "./review-split.mjs";
import { appendDecisions, decisionRecord } from "./decision-log.mjs";
import { SAVE_FIRST } from "./doc-bytes.mjs";
import { runPy as defaultRunPy } from "./py-tool.mjs";
import { runMarkedDiff } from "./marked-diff.mjs";
import { runAnonymize } from "./anonymize.mjs";
import { findFormatIssues, formatRulesOn } from "./format-check.mjs";
import { checkCountItem } from "./count-check.mjs";
import { refreshed, usableHeadings, writingJob } from "./response-service.mjs";
import { buildGlossPrompt } from "./response-prompt.mjs";
import { MANUSCRIPT, submissionPaths, submissionRows } from "./submission-check.mjs";
import { sha256File } from "./figure-service.mjs";

// The Tools tab names the format rules that are on (format-check.mjs).
export { formatRulesOn };

export const AUX_JOBS = new Set([
  "selection_translate",
  "selection_ask",
  "profile_rules_extract",
  "marked_build",
  "anonymize",
  "submission_check",
]);

export const MARKED_MODES = ["track_changes", "color"];

// A file arrives as base64 in one WebSocket frame of at most 100 MiB.
const MAX_BASE64 = 100 * 1024 * 1024;

// The longest selection that is sent to a model, in characters.
export const MAX_SELECTION = 20000;

export const ASK_SYSTEM =
  "Answer the question about the text of a scientific manuscript. Use only the text and general knowledge; say so when the text does not answer the question. Reply with the answer only.";

export const buildAskPrompt = (text, question) => ({
  system: ASK_SYSTEM,
  prompt: `Text:\n${text}\n\nQuestion: ${question}`,
  expectJson: false,
});

const isObj = (x) => !!x && typeof x === "object" && !Array.isArray(x);

function selectedText(msg) {
  const text = String(msg.text ?? "").trim();
  if (!text) throw new Error("Select some text in the document first.");
  if (text.length > MAX_SELECTION)
    throw new Error(
      `The selection has ${text.length} characters; at most ${MAX_SELECTION} are sent. Select less.`,
    );
  return text;
}

export function createAuxService({
  getContext,
  // The one-shot of ONE JOB of a pane: async (key, { role, claudeModel }?) =>
  // { run, model } (providers/oneshot-router.mjs pin).
  oneshotFor = async () => ({ run: null, model: null }),
  // Default language of the settings file: () => string | null.
  langDefault = null,
  sendEvent = () => {},
  runPy = defaultRunPy,
  exists = existsSync,
  log = console.log,
  env = process.env,
}) {
  // The saved file of a pane, read from its live context once.
  function at(key) {
    const doc = localDocPath(getContext(key)?.activeDoc, env);
    const need = () => {
      if (!doc) throw new Error(SAVE_FIRST);
      return doc;
    };
    return {
      key,
      get doc() {
        return need();
      },
      get path() {
        return revisionPathFor(need());
      },
    };
  }
  const load = (d) => loadRevision(d.path, d.doc);
  const update = (d, fn) => withRevision(d.path, d.doc, fn);
  const profileOf = (d, rev) =>
    effectiveProfile(rev, {
      docDir: dirname(d.doc),
      env,
      settingsDefault: langDefault?.() ?? null,
    });

  // The saved manuscript as it is on the disk, before a copy is made from it
  // (null: it cannot be read; the tool will say so itself).
  const sourceHash = (d) => sha256File(d.doc).catch(() => null);
  // Puts a copy that was made and passed its check on record
  // (revision.artifacts), for the submission checklist: the copy as it is
  // now and the manuscript it was made from. Recorded only when the
  // manuscript on the disk is still the file it was before the tool ran;
  // otherwise the record of that copy is taken away, and the checklist asks
  // for the copy again. A record that cannot be written does not fail the
  // copy: the file is there and the pane names it. A manuscript without a
  // review file gets none for this: a copy made from the Tools tab alone
  // leaves nothing else next to the manuscript.
  async function recordCopy(d, kind, path, sourceBefore, extra = {}) {
    if (!existsSync(d.path)) return;
    try {
      const sha256 = await sha256File(path);
      const source = await sourceHash(d);
      const same = !!source && source === sourceBefore;
      const docDir = dirname(d.doc);
      const rel = relative(docDir, path);
      const inside = rel && !rel.startsWith("..") && !isAbsolute(rel);
      const artifact = same
        ? {
            path: inside ? rel.split(sep).join("/") : path,
            sha256,
            at: new Date().toISOString(),
            source_sha256: source,
            ...extra,
          }
        : null;
      await update(d, (rev) => setArtifact(rev, kind, artifact));
      log(`[tools] ${kind} copy recorded=${same}`);
    } catch (e) {
      log(`[tools] ${kind} copy not recorded: ${e?.message ?? e}`);
    }
  }

  // The model of one job, or an error the pane shows.
  async function modelFor(key, what, options) {
    const job = await oneshotFor(key, options);
    if (!job?.run) throw new Error(`No model is available for ${what}. Pick a provider in Setup.`);
    return job;
  }
  // A translation or an answer is prose: a job of the pane's writing choice,
  // as a response draft is.
  const proseModel = writingJob(env);
  async function prose(job, request, what) {
    const r = await job.run(request);
    const text = r.ok ? String(r.text ?? "").trim() : "";
    if (!text) throw new Error(`${what} failed: ${r.error ?? "empty answer"}`);
    return text;
  }
  // The receipt of a selection request, next to a saved document: the
  // model and sizes, no text. A document that is not saved has no place for
  // one; a receipt that cannot be written does not fail the answer.
  async function selectionReceipt(key, kind, input, model, output) {
    const doc = localDocPath(getContext(key)?.activeDoc, env);
    if (!doc) return;
    await appendDecisions(decisionsPathFor(doc), [
      decisionRecord({ kind, input, model, output }),
    ]).then(undefined, (e) => log(`[tools] decision log failed: ${e?.message ?? e}`));
  }

  const impl = {
    // The selected text in another language. Only a receipt is written.
    async selection_translate(msg, key) {
      const text = selectedText(msg);
      const lang = String(msg.lang ?? "").trim();
      if (!lang) throw new Error("Name the language to translate into.");
      const job = await modelFor(key, "translating", proseModel);
      const out = await prose(job, buildGlossPrompt(text, lang), "translation");
      log(`[tools] selection_translate chars=${text.length} lang=${lang} model=${job.model}`);
      await selectionReceipt(key, "selection_translate", text, job.model, {
        lang,
        chars: text.length,
      });
      return { text: out, lang, model: job.model };
    },

    // An answer to a question about the selected text. Only a receipt is written.
    async selection_ask(msg, key) {
      const text = selectedText(msg);
      const question = String(msg.question ?? "").trim();
      if (!question) throw new Error("Type a question about the selected text.");
      const job = await modelFor(key, "answering", proseModel);
      const out = await prose(job, buildAskPrompt(text, question), "the answer");
      log(`[tools] selection_ask chars=${text.length} model=${job.model}`);
      await selectionReceipt(key, "selection_ask", `${text}\n${question}`, job.model, {
        chars: text.length,
      });
      return { text: out, model: job.model };
    },

    // Rule candidates from the guideline of a journal. Nothing is stored.
    async profile_rules_extract(msg, key) {
      const text = String(msg.text ?? "").trim();
      if (!text) throw new Error("Paste the guideline text first.");
      // Before the model call: the receipt goes next to this document.
      const receipts = decisionsPathFor(at(key).doc);
      const job = await modelFor(key, "reading the guideline");
      const r = await extractGuidelineRules(text, { oneshot: job.run, splitSentences });
      const kinds = r.candidates.map((c) => (c.profile ? "format" : c.check.type));
      // Codes and counts only: no word of the guideline or of the reply.
      log(
        `[tools] guideline rules candidates=${r.candidates.length} warnings=${r.warnings.length} reason=${r.reason ?? "-"} problem=${r.problem ? JSON.stringify(r.problem) : "-"} asks=${r.asks} model=${job.model}`,
      );
      try {
        await appendDecisions(receipts, [
          decisionRecord({
            kind: "guideline_rules",
            input: text,
            model: job.model,
            output: {
              candidates: r.candidates.length,
              kinds,
              warnings: r.warnings.length,
              reason: r.reason,
            },
          }),
        ]);
      } catch (e) {
        log(`[tools] decision log failed: ${e?.message ?? e}`);
      }
      return {
        candidates: r.candidates,
        warnings: r.warnings,
        reason: r.reason,
        problem: r.problem,
        model: job.model,
      };
    },

    // The candidates the author picked: rules into the ledger, format
    // settings into the profile of this paper. A rule already in the ledger
    // is skipped.
    async profile_rules_apply(msg, key) {
      const d = at(key);
      const picks = Array.isArray(msg.candidates) ? msg.candidates : [];
      if (!picks.length) throw new Error("Pick at least one rule.");
      const sentences = new Set(splitSentences(String(msg.text ?? "")));
      for (const c of picks) {
        if (!sentences.has(c?.source?.sentence))
          throw new Error(
            "A picked rule cites a sentence that is not in the guideline text. Extract the rules again.",
          );
      }
      const overrides = picks.reduce(
        (all, c) => (c.profile ? mergeProfile(all, c.profile) : all),
        {},
      );
      let added = [];
      let skipped = [];
      await update(d, (rev) => {
        // Rules and the profile belong to the manuscript of a supplement.
        refuseInPart(rev);
        const r = importLedgerItems(rev, picks);
        ({ added, skipped } = r);
        if (!Object.keys(overrides).length) return r.revision;
        const next = setProfile(r.revision, { overrides });
        const problems = regexProblems(next.profile_overrides);
        if (problems.length) throw new Error(`Not saved: ${problems.join("; ")}`);
        return next;
      });
      log(
        `[tools] guideline rules added=${added.length} skipped=${skipped.length} profile=${Object.keys(overrides).join(",") || "-"}`,
      );
      // The Review tab shows the ledger: it loads the saved state again.
      sendEvent(key, { type: "review_event", event: "ledger_changed" });
      return {
        added: added.length,
        skipped: skipped.length,
        added_ids: added.map((i) => i.id),
        profile_keys: Object.keys(overrides),
      };
    },

    // A copy of the saved file that shows what changed since the earlier
    // version in the request. mode: "track_changes" or "color".
    async marked_build(msg, key) {
      const d = at(key);
      const clean = d.doc;
      if (!MARKED_MODES.includes(msg.mode))
        throw new Error("Choose how the changes are shown: tracked changes or colours.");
      const base64 = String(msg.baseline_base64 ?? "");
      if (!base64) throw new Error("Choose the earlier version (.docx) first.");
      if (base64.length > MAX_BASE64) throw new Error("The earlier version is too large to read.");
      const { profile } = await profileOf(d, await load(d));
      const source = await sourceHash(d);
      const dir = await mkdtemp(join(tmpdir(), "marked-"));
      try {
        // A name of our own: the file name of the pane is not used as a path.
        const baseline = join(dir, "baseline.docx");
        await writeFile(baseline, Buffer.from(base64, "base64"));
        const r = await runMarkedDiff({
          baseline,
          clean,
          docPath: clean,
          profile: { ...profile, marked_mode: msg.mode },
          source: "disk",
          runPy,
          exists,
        });
        log(`[tools] marked_build mode=${msg.mode} ok=${r.ok} path=${r.path ?? "-"}`);
        if (!r.ok) throw new Error(r.error);
        await recordCopy(d, "marked", r.path, source, { mode: msg.mode });
        return {
          path: r.path,
          name: basename(r.path),
          mode: msg.mode,
          report: r.report,
          notice: r.notice,
        };
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    },

    // An anonymized copy of the saved file, with what the verifier found in
    // it. A copy that does not pass is not a result to use: ok is false, the
    // report and the name it was set aside under come with the error.
    async anonymize(_msg, key) {
      const d = at(key);
      const source = await sourceHash(d);
      const r = await runAnonymize({ src: d.doc, runPy, exists });
      log(`[tools] anonymize ok=${r.ok} total=${r.report?.total ?? "-"} path=${r.path ?? "-"}`);
      if (!r.ok && !r.report) throw new Error(r.error);
      // Only a copy the verifier passed is put on record: one that was set
      // aside as INCOMPLETE is not a copy to submit.
      if (r.ok) await recordCopy(d, "anonymized", r.path, source);
      return { ...r, name: r.path ? basename(r.path) : null };
    },

    // Format findings and count results for the document as the pane read
    // it. Report only: nothing is written to the document or the revision.
    async format_check(msg, key) {
      const d = at(key);
      const snapshot = msg.snapshot;
      if (!isObj(snapshot) || !Array.isArray(snapshot.paragraphs))
        throw new Error("The request holds no document snapshot.");
      const { headings } = usableHeadings(msg.headings, snapshot);
      const rev = await load(d);
      const { profile } = await profileOf(d, rev);
      const read = msg.format?.paragraphs;
      const format = { rules: formatRulesOn(profile), findings: [], error: null };
      if (!Array.isArray(read) || read.length !== snapshot.paragraphs.length) {
        format.error = `The formatting was read from ${Array.isArray(read) ? read.length : 0} paragraphs and the text from ${snapshot.paragraphs.length}; the format rules were not checked. Run the checks again.`;
      } else {
        format.findings = findFormatIssues(read, profile, { headings });
        if (msg.format.tables_read === false)
          format.error =
            "The tables could not be matched to their paragraphs; empty tables were not checked.";
      }
      const counts = rev.ledger
        .filter((item) => item.check.type === "count")
        .map((item) => ({
          text: item.text,
          ...checkCountItem(item, snapshot.paragraphs, { headings, profile }),
        }));
      log(
        `[tools] format_check paragraphs=${snapshot.paragraphs.length} headings=${headings.length} rules=${format.rules.length} findings=${format.findings.length} counts=${counts.length}`,
      );
      return { format, counts };
    },

    // The rows of the submission checklist for the open file. A job: it
    // hashes files, and a plain request is waited for 10 s only. It reads
    // only: the record is refreshed in memory, as the Response tab's check
    // refreshes it, and is not saved.
    async submission_check(msg, key) {
      const d = at(key);
      const snapshot = msg.snapshot;
      if (!isObj(snapshot) || !Array.isArray(snapshot.paragraphs))
        throw new Error("The request holds no document snapshot.");
      const { headings } = usableHeadings(msg.headings, snapshot);
      const stored = await load(d);
      const { profile } = await profileOf(d, stored);
      const revision = refreshed(stored, snapshot, headings, profile);
      // The files the record names, as they are on the disk now: a path of
      // the record is absolute, or relative to the manuscript folder with
      // "/" (as the figure files are stored). null: not there or not readable.
      const files = { [MANUSCRIPT]: await sourceHash(d) };
      for (const path of submissionPaths(revision)) {
        const abs = isAbsolute(path) ? path : join(dirname(d.doc), ...path.split("/"));
        files[path] = await sha256File(abs).catch(() => null);
      }
      const rows = submissionRows({
        revision,
        profile,
        live: { snapshot, headings, format: isObj(msg.format) ? msg.format : null },
        files,
        file: basename(d.doc),
      });
      const count = (state) => rows.filter((r) => r.state === state).length;
      log(
        `[tools] submission_check rows=${rows.length} ready=${count("ready")} not_ready=${count("not_ready")} not_applicable=${count("not_applicable")}`,
      );
      return { rows, doc_name: basename(d.doc) };
    },

    // Which saved file the copies are made from, and the profile in force.
    async tools_state(_msg, key) {
      const d = at(key);
      const { name, profile } = await profileOf(d, await load(d));
      return { doc_name: basename(d.doc), profile_name: name, policy_packs: profile.policy_packs };
    },
  };

  return {
    handlers: wrapHandlers(impl, {
      jobTypes: AUX_JOBS,
      sendEvent,
      eventType: "tools_event",
      log,
    }),
  };
}
