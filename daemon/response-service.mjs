// Response tab (C2): task-pane requests about the response letter.
//
// Requests (bridge.mjs extraHandlers): { type, request_id, ...payload } and the
// reply { type: "<type>_result", request_id, ok, ... }.
//   profile_get, profile_set, response_state, response_save,
//   response_import_previous
// Jobs (RESPONSE_JOBS; they call a model or a Python tool): accepted at once,
// the result follows as { type: "response_event", event: "job_done", job_id,
// result }.
//   profile_propose, response_draft, response_translate, response_locate,
//   response_check, response_export
//
// Who writes what. The model writes the prose of an answer and nothing else.
// States of the changes, citations, section names, line ranges, stale marks
// and every check come from code and from what the pane read in Word
// (snapshot, headings, the rendered PDF); a request cannot send them in.
//
// Handler rules (as in review-service.mjs): the document of a pane is read
// once per request (at(key)) and never again after an await; a failure is an
// error reply, never a silent default.

import { mkdtemp, rm, unlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  localDocPath,
  revisionPathFor,
  decisionsPathFor,
  sidecarPath,
  loadRevision,
  withRevision,
  writeJsonAtomic,
  findComment,
  applyStates,
  setResponse,
  setDeclined,
  setCitations,
  setProfile,
  setRender,
  setExport,
} from "./revision-store.mjs";
import { classifyRevision } from "./revision-classify.mjs";
import { appendDecisions, decisionRecord } from "./decision-log.mjs";
import { wrapHandlers } from "./handler-wrap.mjs";
import { checkLlmItem, inScope } from "./ledger-check.mjs";
import {
  effectiveProfile,
  listProfiles,
  loadProfile,
  profileDirs,
  regexProblems,
} from "./journal-profile.mjs";
import {
  STALE_ORDER,
  citationsFor,
  keepLines,
  markStale,
  snapshotHash,
  staleReasons,
  viewText,
} from "./response-quotes.mjs";
import { applyLineResults, lineRequests, locate, savePdf } from "./response-lines.mjs";
import { answerProse, buildLetterModel, letterPlainText } from "./response-model.mjs";
import { checkResponses, summarize, wordCount } from "./response-check.mjs";
import { buildGlossPrompt, buildResponsePrompt, parseDraft } from "./response-prompt.mjs";
import { runPy as defaultRunPy } from "./py-tool.mjs";
import { freeOutputIndex, outputName, MAX_OUTPUT_NAMES } from "./output-name.mjs";

export const RESPONSE_JOBS = new Set([
  "profile_propose",
  "response_draft",
  "response_translate",
  "response_locate",
  "response_check",
  "response_export",
]);

// An answer is prose, so a draft asks Claude for a stronger model than a
// judgment does. Only a default: the provider and model chosen in the pane
// are what is used (providers/oneshot-router.mjs pin).
export const DRAFT_CLAUDE_MODEL = "sonnet";

// A PDF arrives as base64 in one WebSocket frame of at most 100 MiB.
const MAX_BASE64 = 100 * 1024 * 1024;

const TOOLS = join(dirname(fileURLToPath(import.meta.url)), "..", "tools");
const tool = (name) => join(TOOLS, name);
const isObj = (x) => !!x && typeof x === "object" && !Array.isArray(x);

// The sidecars written with a letter, named after the docx.
const letterGroup = (docx) => ({
  docx,
  model: docx.replace(/\.docx$/, ".model.json"),
  manuscript: docx.replace(/\.docx$/, ".manuscript.json"),
});

// The files of an export are written one request at a time per manuscript
// (two panes may have the same document open): the second request looks for
// a free name only after the first has its files. A chain of its own, not
// withRevision's: an export saves the revision after its files are written.
const exporting = new Map();
function oneExport(doc, fn) {
  const run = (exporting.get(doc) ?? Promise.resolve()).catch(() => {}).then(fn);
  exporting.set(doc, run);
  return run;
}

// The headings the pane lists, kept only where the snapshot agrees: the
// paragraph is there and reads the same. The section name is then the
// snapshot's own text.
export function usableHeadings(headings, snapshot) {
  const byIndex = new Map(snapshot.paragraphs.map((p) => [p.index, viewText(p)]));
  const out = [];
  let dropped = 0;
  for (const h of Array.isArray(headings) ? headings : []) {
    const text = Number.isInteger(h?.index) ? byIndex.get(h.index) : undefined;
    const sameText = text && text === viewText({ text: h.text });
    if (sameText && Number.isInteger(h.level) && h.level >= 1)
      out.push({ index: h.index, level: h.level, text });
    else dropped += 1;
  }
  return { headings: out, dropped };
}

// What a draft is written from, as one value per stale mark: the sentences
// the letter quotes for the comment, and its changes with their states (the
// prompt lists each change as in the manuscript, revised by hand, or not in
// it). Compared when the draft is saved (response_draft).
const draftInput = (c) => ({
  quote_changed: JSON.stringify(
    (c.response?.citations ?? []).map((x) => [x.kind, x.quote, x.problem]),
  ),
  change_state: JSON.stringify((c.changes ?? []).map((ch) => [ch.id, ch.state])),
});

// States, then citations, then stale marks, from one snapshot. Pure.
//   - a change is cited only while it is in the manuscript (accepted or
//     revised by the author); applyStates runs first so citationsFor reads
//     the current state and paragraph of every change
//   - the stale marks are judged against the citations as saved, before the
//     fresh ones take their place. quote_changed and change_state then stay
//     until the answer is written again; lines_old follows the PDF.
export function refreshed(rev, snapshot, headings, profile) {
  let next = applyStates(rev, classifyRevision(rev, snapshot));
  const level = profile.response_letter.section_level;
  next = markStale(next, snapshot);
  for (const c of next.comments) {
    const fresh = keepLines(
      citationsFor(c, snapshot, headings, { level }),
      c.response?.citations ?? [],
    );
    if (!c.response && !fresh.length) continue;
    next = setCitations(next, c.ref, fresh);
  }
  return next;
}

export function createResponseService({
  getContext,
  // The one-shot of ONE JOB of a pane: async (key, { claudeModel }?) =>
  // { run, model } (providers/oneshot-router.mjs pin). run is null when the
  // daemon has no model to ask.
  oneshotJob = async () => ({ run: null, model: null }),
  // Default language of the settings file: () => string | null.
  langDefault = null,
  sendEvent = () => {},
  runPy = defaultRunPy,
  log = console.log,
  env = process.env,
}) {
  // The document of a pane, read from its live context once (see
  // review-service.mjs at()). A pane without a usable document gives a
  // handle that throws on first use.
  function at(key) {
    const doc = localDocPath(getContext(key)?.activeDoc, env);
    const need = () => {
      if (!doc)
        throw new Error(
          "Save the document to a local or WSL-reachable folder first. The response letter is written next to it.",
        );
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
  const dirsOf = (d) => profileDirs({ docDir: dirname(d.doc), env });
  const profileOf = (d, rev) =>
    effectiveProfile(rev, {
      docDir: dirname(d.doc),
      env,
      settingsDefault: langDefault?.() ?? null,
    });

  async function receipt(d, records) {
    try {
      await appendDecisions(decisionsPathFor(d.doc), records);
    } catch (e) {
      log(`[response] decision log failed: ${e?.message ?? e}`);
    }
  }

  // What the pane read in Word, checked before anything is computed from it.
  function docOf(msg) {
    const snapshot = msg.snapshot;
    if (!isObj(snapshot) || !Array.isArray(snapshot.paragraphs))
      throw new Error("The request holds no document snapshot.");
    const { headings, dropped } = usableHeadings(msg.headings, snapshot);
    return { snapshot, headings, dropped };
  }

  // The revision brought up to date with the document, saved, and its profile.
  // peek(rev, snapshot): called with the stored revision right before it is
  // brought up to date.
  async function current(d, msg, peek = null) {
    const { snapshot, headings, dropped } = docOf(msg);
    if (dropped) log(`[response] headings dropped=${dropped}`);
    const { profile } = await profileOf(d, await load(d));
    const revision = await update(d, (rev) => {
      peek?.(rev, snapshot);
      return refreshed(rev, snapshot, headings, profile);
    });
    return { snapshot, headings, dropped, profile, revision };
  }
  // What the pane found as headings, for a run that cites no section: how
  // many, at which levels, how many the snapshot did not confirm. No text.
  function logHeadings({ headings, dropped }) {
    const levels = new Map();
    for (const h of headings) levels.set(h.level, (levels.get(h.level) ?? 0) + 1);
    const text = [...levels].sort((a, b) => a[0] - b[0]).map(([l, n]) => `${l}:${n}`);
    log(
      `[response] headings n=${headings.length} levels=${text.join(",") || "-"} dropped=${dropped}`,
    );
  }

  // A tool that could not run (no interpreter, a crash) is an error with its
  // own words; a tool that ran and said no is handled by the caller.
  async function py(name, args) {
    const r = await runPy(tool(name), args);
    if (!r.json) throw new Error(r.error ?? `${name} gave no answer`);
    return r;
  }

  // checkResponses plus the LLM-checked rules of the response scope, asked
  // once per rule for all answers. job: taken when the first rule needs it.
  async function allChecks(d, key, { revision, snapshot, headings, profile }) {
    const rows = checkResponses({ revision, snapshot, headings, profile });
    const rules = (revision.ledger ?? []).filter(
      (i) => inScope(i, "response") && i.check?.type === "llm",
    );
    if (!rules.length) return rows;
    const lines = [];
    // Every paragraph the letter prints as an answer, reasons for declining
    // included (the list the letter itself is written from).
    for (const c of revision.comments)
      for (const text of answerProse(c)) lines.push({ index: lines.length, text, ref: c.ref });
    if (!lines.length) return rows;
    const job = await oneshotJob(key);
    const records = [];
    for (const item of rules) {
      const add = (level, ref, detail) =>
        rows.push({ id: "ledger", level, ref, detail: `${item.id}: ${detail}` });
      if (!job?.run) {
        add("fail", null, "no model is available to check this rule");
        continue;
      }
      const out = await checkLlmItem(item, lines, {
        oneshot: job.run,
        onDecision: (x) =>
          records.push(
            decisionRecord({
              kind: "response_ledger_check",
              input: { instruction: x.instruction, paragraph: x.paragraph },
              model: job.model,
              output: {
                item_id: x.item_id,
                ref: lines[x.paragraph_index]?.ref,
                verdict: x.verdict,
              },
            }),
          ),
      });
      if (out.status === "error") add("fail", null, out.error);
      for (const v of out.violations ?? [])
        add("fail", lines[v.paragraph_index]?.ref ?? null, `${v.match} (${v.reason})`);
      for (const u of out.unsure ?? [])
        add("warn", lines[u.paragraph_index]?.ref ?? null, `not judged: ${u.reason}`);
    }
    if (records.length) await receipt(d, records);
    return rows;
  }

  // ---- export -----------------------------------------------------------------

  // Writes the letter group under the first name that is free (or the plain
  // name with overwrite) and verifies the file. Returns the paths and what
  // the verifier said.
  async function writeLetter(d, model, manuscript, overwrite) {
    const wanted = sidecarPath(d.doc, ".response.docx");
    let replace = overwrite;
    for (let n = 1; ; n++) {
      if (n > MAX_OUTPUT_NAMES) throw new Error(`No free name for ${wanted}.`);
      if (!replace)
        n = await freeOutputIndex(wanted, { from: n, group: (p) => Object.values(letterGroup(p)) });
      const paths = letterGroup(outputName(wanted, n));
      // Without overwrite a sidecar is created, never replaced: a file that
      // took the name since the look is someone else's and stays. `mine` are
      // the files this request made; only they are ever removed.
      const mine = [];
      const dropMine = () => Promise.all(mine.map((p) => unlink(p).catch(() => {})));
      if (replace) {
        await writeJsonAtomic(paths.model, model);
        await writeJsonAtomic(paths.manuscript, manuscript);
      } else {
        let taken = false;
        for (const [path, data] of [
          [paths.model, model],
          [paths.manuscript, manuscript],
        ]) {
          try {
            await writeFile(path, JSON.stringify(data, null, 2) + "\n", { flag: "wx" });
            mine.push(path);
          } catch (e) {
            await dropMine();
            if (e.code !== "EEXIST") throw e;
            taken = true;
            break;
          }
        }
        if (taken) continue;
      }
      let w;
      try {
        w = await py("response_docx.py", [
          "--model",
          paths.model,
          "--out",
          paths.docx,
          ...(replace ? ["--overwrite"] : []),
        ]);
      } catch (e) {
        await dropMine();
        throw e;
      }
      // Taken since we looked, or open in Word: the next free name. The file
      // that is there stays as it is.
      if (w.json.code === "exists" || w.json.code === "locked") {
        await dropMine();
        replace = false;
        continue;
      }
      if (!w.ok) {
        await dropMine();
        throw new Error(w.json.error ?? w.error ?? "the letter could not be written");
      }
      if (!replace) mine.push(paths.docx);
      const v = await py("response_verify.py", [
        "--docx",
        paths.docx,
        "--model",
        paths.model,
        "--manuscript",
        paths.manuscript,
      ]);
      if (!v.json.ok) {
        // A new file of ours that is not what the model says does not stay.
        await dropMine();
        const axes = [...new Set((v.json.failures ?? []).map((f) => f.axis))].join(", ");
        throw new Error(
          `The written letter failed verification (${axes || v.json.error || "unknown"}).` +
            (replace ? ` The file is at ${paths.docx}.` : ""),
        );
      }
      return {
        paths,
        wanted,
        paragraphs: w.json.paragraphs ?? model.paragraphs.length,
        verify: v.json,
      };
    }
  }

  // One plain-text file per reviewer (the editor: E).
  async function writeText(d, model, overwrite) {
    const reviewers = [...new Set(model.paragraphs.map((p) => p.reviewer))].sort((a, b) => a - b);
    const out = [];
    const wantedAll = [];
    for (const reviewer of reviewers) {
      const wanted = sidecarPath(d.doc, `.response.${reviewer === 0 ? "E" : `R${reviewer}`}.txt`);
      wantedAll.push(wanted);
      const text = `${letterPlainText(model, reviewer)}\n`;
      for (let n = 1; ; n++) {
        if (!overwrite) n = await freeOutputIndex(wanted, { from: n });
        const path = outputName(wanted, n);
        try {
          await writeFile(path, text, { encoding: "utf8", flag: overwrite ? "w" : "wx" });
          out.push(path);
          break;
        } catch (e) {
          if (overwrite || e.code !== "EEXIST") throw e;
        }
      }
    }
    return { paths: out, wanted: wantedAll };
  }

  const impl = {
    async profile_get(msg, key) {
      const d = at(key);
      const rev = await load(d);
      const available = await listProfiles({ dirs: dirsOf(d) });
      try {
        const p = await profileOf(d, rev);
        return {
          name: p.name,
          path: p.path,
          profile: p.profile,
          overrides: rev.profile_overrides,
          available,
          problems: p.problems,
        };
      } catch (e) {
        // The pane must still be able to pick another profile.
        return {
          name: rev.journal_profile,
          path: null,
          profile: null,
          overrides: rev.profile_overrides,
          available,
          problems: [e?.message ?? String(e)],
        };
      }
    },

    async profile_set(msg, key) {
      const d = at(key);
      const name = msg.journal_profile;
      if (name !== undefined && name !== null) await loadProfile(name, { dirs: dirsOf(d) });
      if (msg.overrides !== undefined && !isObj(msg.overrides))
        throw new Error("overrides must be an object.");
      const revision = await update(d, (rev) => {
        const next = setProfile(rev, { journal_profile: name, overrides: msg.overrides });
        const problems = regexProblems(next.profile_overrides);
        if (problems.length) throw new Error(`Not saved: ${problems.join("; ")}`);
        return next;
      });
      const p = await profileOf(d, revision);
      log(
        `[response] profile name=${p.name} overrides=${Object.keys(revision.profile_overrides).join(",") || "-"}`,
      );
      return { revision, name: p.name, path: p.path, profile: p.profile, problems: p.problems };
    },

    // Reads an existing letter and proposes overrides and earlier answers.
    // Nothing is saved: the pane shows the proposal and the author applies it.
    async profile_propose(msg, key) {
      const d = at(key);
      const base64 = String(msg.docx_base64 ?? "");
      if (!base64) throw new Error("No letter in the request.");
      if (base64.length > MAX_BASE64) throw new Error("The letter is too large to read.");
      let heads = {};
      try {
        const { profile } = await profileOf(d, await load(d));
        heads = {
          reviewer_head: profile.reviewer_head ?? null,
          comment_head: profile.comment_head ?? null,
          response_letter: { editor_label: profile.response_letter.editor_label },
        };
      } catch {
        /* an unreadable profile: the reader uses its own head patterns */
      }
      const dir = await mkdtemp(join(tmpdir(), "resp-"));
      try {
        // Names of our own: the file name of the pane is not used as a path.
        const docx = join(dir, "letter.docx");
        const profile = join(dir, "profile.json");
        await writeFile(docx, Buffer.from(base64, "base64"));
        await writeFile(profile, JSON.stringify(heads));
        const r = await py("response_read.py", ["--docx", docx, "--profile", profile]);
        if (!r.json.ok) throw new Error(r.json.error ?? "the letter could not be read");
        const { overrides = {}, evidence = {}, previous = {} } = r.json;
        return { overrides, evidence, previous };
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    },

    // previous: { "<ref>": "<answer of the earlier round>" }.
    async response_import_previous(msg, key) {
      const d = at(key);
      if (!isObj(msg.previous)) throw new Error("previous must be an object of ref: text.");
      const imported = [];
      const unknown = [];
      const revision = await update(d, (rev) => {
        let next = rev;
        for (const [ref, text] of Object.entries(msg.previous)) {
          if (!findComment(next, ref) || typeof text !== "string" || !text.trim()) {
            unknown.push(ref);
            continue;
          }
          next = setResponse(next, ref, { previous: text });
          imported.push(ref);
        }
        return next;
      });
      return { revision, imported, unknown };
    },

    async response_state(msg, key) {
      const d = at(key);
      const state = await current(d, msg);
      logHeadings(state);
      return { revision: state.revision };
    },

    // The author's own text and decisions. text: the answer; declined:
    // { reason } or null; declined_requirements: the whole list for the comment.
    // snapshot, headings (optional): the document as it is now. With them the
    // citations are brought up to date first, so the text is saved against
    // the changes it answers and a later change shows as stale.
    async response_save(msg, key) {
      const d = at(key);
      const ref = msg.ref;
      // The text was typed against the citations the card showed, and those
      // are the stored ones. Where the document of this request no longer
      // agrees with them (a shown sentence edited, a cited change taken
      // back), the text is saved with that mark: the author reads it against
      // the citations as they are now and saves again. A card that showed no
      // citation gives nothing to compare.
      let moved = [];
      if (msg.snapshot !== undefined)
        await current(d, msg, (rev, snapshot) => {
          const shown = findComment(applyStates(rev, classifyRevision(rev, snapshot)), ref);
          if (shown?.response?.citations?.length) moved = staleReasons(shown, snapshot, null);
        });
      const marked = "text" in msg && moved.length > 0;
      const revision = await update(d, (rev) => {
        const comment = findComment(rev, ref);
        if (!comment) throw new Error(`No reviewer comment ${ref}`);
        let next = rev;
        if ("text" in msg) {
          const text = String(msg.text ?? "");
          // The translation was made of the text it stood under: under another
          // text it is dropped (the check then asks for a new one), never kept.
          const gloss = text === (comment.response?.text ?? "") ? {} : { gloss: "" };
          next = setResponse(next, ref, { text, source: "author", ...gloss });
          // Written again: what was stale about the old text is not any more.
          const r = findComment(next, ref).response;
          next = setCitations(
            next,
            ref,
            r.citations,
            STALE_ORDER.filter((s) =>
              s === "lines_old" ? r.stale.includes(s) : moved.includes(s),
            ),
          );
        }
        if ("declined" in msg) next = setDeclined(next, ref, msg.declined);
        if ("declined_requirements" in msg) {
          const list = msg.declined_requirements;
          if (!Array.isArray(list)) throw new Error("declined_requirements must be a list.");
          const known = new Set((comment.requirements ?? []).map((q) => q.id));
          const seen = new Set();
          for (const x of list) {
            if (!known.has(x?.id)) throw new Error(`No request ${x?.id} in ${ref}`);
            if (seen.has(x.id)) throw new Error(`Request ${x.id} is listed twice.`);
            seen.add(x.id);
            if (!String(x.reason ?? "").trim())
              throw new Error(`A declined request needs a reason: ${x.id}`);
          }
          next = setResponse(next, ref, {
            declined_requirements: list.map((x) => ({ id: x.id, reason: String(x.reason).trim() })),
          });
        }
        return next;
      });
      if (marked) log(`[response] save ${ref} input changed: ${moved.join(",")}`);
      return {
        revision,
        warnings: marked
          ? [
              "The manuscript changed under this answer. Read it against the citations shown now, then save it again.",
            ]
          : [],
      };
    },

    async response_draft(msg, key) {
      const d = at(key);
      const ref = msg.ref;
      // Provider and model of this job, fixed now (the pane's choice).
      const job = await oneshotJob(key, {
        claudeModel: env.CC_OFFICE_RESPONSE_MODEL || DRAFT_CLAUDE_MODEL,
      });
      if (!job?.run)
        throw new Error("No model is available for drafting. Pick a provider in Setup.");
      const { snapshot, headings, profile, revision: before } = await current(d, msg);
      const comment = findComment(before, ref);
      if (!comment) throw new Error(`No reviewer comment ${ref}`);
      const L = profile.response_letter;
      const written = draftInput(comment);
      const marksThen = comment.response?.stale ?? [];
      const built = buildResponsePrompt({
        comment,
        profile,
        ledger: before.ledger,
        previous: comment.response?.text || comment.response?.previous || null,
      });
      const drafted = parseDraft(await job.run({ system: built.system, prompt: built.prompt }));
      if (drafted.error) throw new Error(drafted.error);
      const warnings = [];
      let gloss = "";
      if (L.gloss_enabled && L.gloss_lang) {
        const g = await job.run(buildGlossPrompt(drafted.text, L.gloss_lang));
        gloss = g.ok ? String(g.text ?? "").trim() : "";
        if (!gloss) warnings.push(`translation failed: ${g.error ?? "empty answer"}`);
      }
      // The model took its time: another request may have read the document
      // again since (a sentence edited, a change decided). The draft is kept,
      // and what it was not written for is marked, as a refresh marks it.
      let moved = [];
      const revision = await update(d, (rev) => {
        const now = findComment(rev, ref);
        if (!now) throw new Error(`No reviewer comment ${ref}`);
        const input = draftInput(now);
        // The Review tab's refresh marks and leaves the citations as they
        // are: a mark that was not there when the draft was asked for tells
        // the same as a changed quote.
        const marks = now.response?.stale ?? [];
        moved = Object.keys(written).filter(
          (k) => input[k] !== written[k] || (marks.includes(k) && !marksThen.includes(k)),
        );
        const next = setResponse(rev, ref, { text: drafted.text, gloss, source: "model" });
        const r = findComment(next, ref).response;
        const stale = moved.length
          ? STALE_ORDER.filter((s) => moved.includes(s) || r.stale.includes(s))
          : r.stale.filter((s) => s === "lines_old");
        return setCitations(next, ref, r.citations, stale);
      });
      if (moved.length) {
        log(`[response] draft ${ref} input changed: ${moved.join(",")}`);
        warnings.push(
          "The manuscript changed while this answer was drafted. Read it, then save or draft it again.",
        );
      }
      const words = wordCount(drafted.text);
      await receipt(d, [
        decisionRecord({
          kind: "response_draft",
          input: built.prompt,
          model: job.model,
          output: {
            ref,
            lang: L.lang,
            words,
            change_ids: built.changeIds,
            instruction_ids: built.instructionIds,
            forbid_ids: built.forbidIds,
          },
        }),
      ]);
      const citations = findComment(revision, ref).response.citations.length;
      log(
        `[response] draft ${ref} lang=${L.lang} model=${job.model} words=${words} citations=${citations}`,
      );
      const checks = checkResponses({ revision, snapshot, headings, profile }).filter(
        (row) => row.ref === ref,
      );
      return { revision, checks, warnings };
    },

    // The translation paragraph of the answer as it is saved (after the author
    // wrote or changed it by hand; a draft translates its own text). Only the
    // translation is written: the answer, its citations and marks stay.
    async response_translate(msg, key) {
      const d = at(key);
      const ref = msg.ref;
      const job = await oneshotJob(key, {
        claudeModel: env.CC_OFFICE_RESPONSE_MODEL || DRAFT_CLAUDE_MODEL,
      });
      if (!job?.run)
        throw new Error("No model is available for translating. Pick a provider in Setup.");
      const before = await load(d);
      const comment = findComment(before, ref);
      if (!comment) throw new Error(`No reviewer comment ${ref}`);
      const L = (await profileOf(d, before)).profile.response_letter;
      if (!L.gloss_enabled || !L.gloss_lang)
        throw new Error("The second language is off in the letter settings.");
      const text = comment.response?.text ?? "";
      if (!text.trim()) throw new Error(`There is no answer to translate for ${ref}.`);
      const g = await job.run(buildGlossPrompt(text, L.gloss_lang));
      const gloss = g.ok ? String(g.text ?? "").trim() : "";
      if (!gloss) throw new Error(`translation failed: ${g.error ?? "empty answer"}`);
      const revision = await update(d, (rev) => {
        if ((findComment(rev, ref)?.response?.text ?? "") !== text)
          throw new Error(`The answer to ${ref} changed while it was translated. Translate again.`);
        return setResponse(rev, ref, { gloss });
      });
      await receipt(d, [
        decisionRecord({
          kind: "response_translate",
          input: text,
          model: job.model,
          output: { ref, lang: L.gloss_lang, words: wordCount(gloss) },
        }),
      ]);
      log(`[response] translate ${ref} lang=${L.gloss_lang} model=${job.model}`);
      return { revision };
    },

    // pdf_base64: the document as Word rendered it, margin line numbers on.
    async response_locate(msg, key) {
      const d = at(key);
      const state = await current(d, msg);
      logHeadings(state);
      const { snapshot, revision: before } = state;
      const base64 = String(msg.pdf_base64 ?? "");
      if (base64.length > MAX_BASE64) throw new Error("The PDF is too large to read.");
      // <stem>.render.pdf is a file of the daemon, like revision.json: each
      // render replaces the last one, and response_render names it by hash.
      const pdf = await savePdf(d.doc, base64);
      const requests = lineRequests(before.comments, snapshot);
      const dir = await mkdtemp(join(tmpdir(), "resp-"));
      let out;
      try {
        out = await locate(pdf.path, requests, { run: runPy, tmpJson: join(dir, "quotes.json") });
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
      const located = applyLineResults(before.comments, out.results, pdf.sha);
      const same = (a, b) =>
        a.length === b.length && a.every((x, i) => x.kind === b[i].kind && x.quote === b[i].quote);
      const revision = await update(d, (rev) => {
        let next = setRender(rev, {
          pdf_sha: pdf.sha,
          snapshot_hash: snapshotHash(snapshot),
          at: new Date().toISOString(),
          lines_indexed: out.lines_indexed,
        });
        for (const { ref, citations } of located) {
          const r = findComment(next, ref)?.response;
          // Answers are matched by position: skip a comment whose citations
          // changed while the PDF was read.
          if (!r || !citations.length || !same(r.citations, citations)) continue;
          next = setCitations(
            next,
            ref,
            citations,
            r.stale.filter((s) => s !== "lines_old"),
          );
        }
        return next;
      });
      const found = out.results.filter((x) => x.match).length;
      log(
        `[response] locate pdf=${pdf.bytes} lines_indexed=${out.lines_indexed} located=${found}/${requests.length}`,
      );
      return {
        revision,
        lines_indexed: out.lines_indexed,
        located: `${found}/${requests.length}`,
        pdf_path: pdf.path,
      };
    },

    async response_check(msg, key) {
      const d = at(key);
      const state = await current(d, msg);
      const checks = await allChecks(d, key, state);
      const summary = summarize(checks);
      log(`[response] check fail=${summary.fail} warn=${summary.warn}`);
      return { revision: state.revision, checks, summary };
    },

    // format: "docx" (the letter, with its model and the manuscript paragraph
    // list next to it) or "text" (one file per reviewer). force: write although
    // a check fails. overwrite: replace the files of the plain names; without
    // it a file that is there is never touched and the next free name is used.
    async response_export(msg, key) {
      const d = at(key);
      const format = msg.format ?? "docx";
      if (format !== "docx" && format !== "text")
        throw new Error(`Unknown export format ${JSON.stringify(format)} (docx or text).`);
      const state = await current(d, msg);
      const checks = await allChecks(d, key, state);
      const summary = summarize(checks);
      if (summary.fail && !msg.force) {
        log(`[response] export held fail=${summary.fail}`);
        return { exported: false, checks, summary };
      }
      const { profile, snapshot } = state;
      const model = buildLetterModel(state.revision, profile, {
        gloss: !!profile.response_letter.gloss_enabled,
      });
      const overwrite = msg.overwrite === true;
      if (format === "text") {
        const { paths, wanted } = await oneExport(d.doc, () => writeText(d, model, overwrite));
        log(`[response] export text files=${paths.length}`);
        return { exported: true, paths, wanted, checks, summary };
      }
      const { paths, wanted, paragraphs, verify } = await oneExport(d.doc, () =>
        writeLetter(d, model, snapshot.paragraphs.map(viewText), overwrite),
      );
      const revision = await update(d, (rev) =>
        setExport(rev, {
          path: paths.docx,
          exported_at: new Date().toISOString(),
          fail: summary.fail,
          warn: summary.warn,
        }),
      );
      log(`[response] export path=${paths.docx} paragraphs=${paragraphs} verify=ok`);
      // wanted: the usual name. A different path means a file of that name was
      // there, or open in Word, and was kept.
      return {
        exported: true,
        path: paths.docx,
        wanted,
        paths,
        checks,
        summary,
        verify,
        revision,
      };
    },
  };

  return {
    handlers: wrapHandlers(impl, {
      jobTypes: RESPONSE_JOBS,
      sendEvent,
      eventType: "response_event",
      log,
    }),
  };
}
