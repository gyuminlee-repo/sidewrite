// Tools tab: task-pane requests beside the review and the response letter.
//
// Requests (bridge.mjs extraHandlers): { type, request_id, ...payload } and the
// reply { type: "<type>_result", request_id, ok, ... }.
//   tools_state
// Jobs (AUX_JOBS; they call a model or a Python tool): accepted at once, the
// result follows as { type: "tools_event", event: "job_done", job_id, result }.
//   selection_translate, selection_ask
//
// The selection requests get the selected text from the pane and answer with
// text for the pane. They read no file and no revision, and nothing of the
// answer is stored: the document stays as it is.
//
// Nothing here writes to the open document. A copy (marked, anonymized) is a
// new file next to the manuscript, made from the file as it was last saved.
//
// Handler rules (as in response-service.mjs): the document of a pane is read
// once per request (at(key)); a failure is an error reply, never a silent
// default.

import { basename, dirname } from "node:path";
import { localDocPath, revisionPathFor, loadRevision } from "./revision-store.mjs";
import { wrapHandlers } from "./handler-wrap.mjs";
import { effectiveProfile } from "./journal-profile.mjs";
import { SAVE_FIRST } from "./doc-bytes.mjs";
import { buildGlossPrompt } from "./response-prompt.mjs";

export const AUX_JOBS = new Set(["selection_translate", "selection_ask"]);

// A translation or an answer is prose: the model of a response draft
// (response-service.mjs DRAFT_CLAUDE_MODEL), unless the pane chose another.
const PROSE_CLAUDE_MODEL = "sonnet";

// The longest selection that is sent to a model, in characters.
export const MAX_SELECTION = 20000;

export const ASK_SYSTEM =
  "Answer the question about the text of a scientific manuscript. Use only the text and general knowledge; say so when the text does not answer the question. Reply with the answer only.";

export const buildAskPrompt = (text, question) => ({
  system: ASK_SYSTEM,
  prompt: `Text:\n${text}\n\nQuestion: ${question}`,
  expectJson: false,
});

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
  // The one-shot of ONE JOB of a pane: async (key, { claudeModel }?) =>
  // { run, model } (providers/oneshot-router.mjs pin).
  oneshotFor = async () => ({ run: null, model: null }),
  // Default language of the settings file: () => string | null.
  langDefault = null,
  sendEvent = () => {},
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
  const profileOf = (d, rev) =>
    effectiveProfile(rev, {
      docDir: dirname(d.doc),
      env,
      settingsDefault: langDefault?.() ?? null,
    });

  // The model of one job, or an error the pane shows.
  async function modelFor(key, what, options) {
    const job = await oneshotFor(key, options);
    if (!job?.run) throw new Error(`No model is available for ${what}. Pick a provider in Setup.`);
    return job;
  }
  const proseModel = { claudeModel: env.CC_OFFICE_RESPONSE_MODEL || PROSE_CLAUDE_MODEL };
  async function prose(job, request, what) {
    const r = await job.run(request);
    const text = r.ok ? String(r.text ?? "").trim() : "";
    if (!text) throw new Error(`${what} failed: ${r.error ?? "empty answer"}`);
    return text;
  }

  const impl = {
    // The selected text in another language. Nothing is written anywhere.
    async selection_translate(msg, key) {
      const text = selectedText(msg);
      const lang = String(msg.lang ?? "").trim();
      if (!lang) throw new Error("Name the language to translate into.");
      const job = await modelFor(key, "translating", proseModel);
      const out = await prose(job, buildGlossPrompt(text, lang), "translation");
      log(`[tools] selection_translate chars=${text.length} lang=${lang} model=${job.model}`);
      return { text: out, lang, model: job.model };
    },

    // An answer to a question about the selected text. Nothing is written.
    async selection_ask(msg, key) {
      const text = selectedText(msg);
      const question = String(msg.question ?? "").trim();
      if (!question) throw new Error("Type a question about the selected text.");
      const job = await modelFor(key, "answering", proseModel);
      const out = await prose(job, buildAskPrompt(text, question), "the answer");
      log(`[tools] selection_ask chars=${text.length} model=${job.model}`);
      return { text: out, model: job.model };
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
