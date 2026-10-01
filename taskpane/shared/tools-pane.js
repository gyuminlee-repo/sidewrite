// Tools tab (Word only). DOM glue: every result is computed by the daemon
// (daemon/aux-service.mjs); the wording of the screen is in pure helpers
// (tools-view.js). The pane reads Word (the selection, the snapshot, the
// headings, the formatting) and never writes to it: a copy is a file of its
// own, made by the daemon from the saved file.
// Text from the daemon or the document is inserted with textContent only.

import {
  docLine,
  selectionHead,
  candidateKind,
  candidateCite,
  rulesAdded,
  extractMessage,
  fileTooLarge,
  markedLines,
  checkRows,
  anonymizeLines,
} from "./tools-view.js";
import { bytesToBase64 } from "./response-view.js";

const CONNECTION_LOST =
  "The connection to the daemon was lost before the result came. Run it again.";

export function initToolsPane({
  sendRequest,
  word,
  // The pane's stored settings (taskpane.js): toolsLang is kept there.
  settings = {},
  saveSettings = () => {},
  // () => the request that confirms the stored providers while it is under
  // way (providers-pane.js whenRestored), else null.
  providersReady = () => null,
  // (role) => what the last request of that kind of work really ran on
  // ("Claude · Opus 5.5"), or null when the pane does not know (taskpane.js).
  usedLabel = () => null,
}) {
  const root = document.getElementById("tools-root");
  if (!root) return { onEvent() {}, reload: async () => {} }; // Excel pane: no Tools tab
  const $ = (id) => document.getElementById(id);
  const state = {
    doc: null, // the tools_state answer, or null when the document is not saved
    rules: null, // { text, candidates, warnings, boxes } of the last guideline read
    busy: false,
    asking: 0, // job requests sent whose { accepted, job_id } has not come yet
    jobs: new Map(), // job_id -> { resolve, reject } of an askJob waiting for job_done
    jobsDone: new Map(), // job_id -> result of a job_done that came before its answer
  };

  function el(tag, props = {}, ...kids) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(props)) {
      if (k === "text") n.textContent = v;
      else n.setAttribute(k, v);
    }
    for (const k of kids) if (k) n.append(k);
    return n;
  }
  function say(text, bad = false) {
    const m = $("tools-message");
    m.textContent = text;
    m.classList.toggle("form-error", bad);
    m.hidden = !text;
  }
  async function ask(type, payload = {}) {
    const r = await sendRequest(type, payload);
    if (!r.ok) throw new Error(r.error || `${type} failed`);
    return r;
  }
  // A job answers { accepted, job_id } at once; its result comes in a
  // job_done tools_event, possibly before the answer.
  async function askJob(type, payload) {
    state.asking += 1;
    let job_id;
    try {
      ({ job_id } = await ask(type, payload));
    } finally {
      state.asking -= 1;
    }
    let result = state.jobsDone.get(job_id);
    state.jobsDone.delete(job_id);
    result ??= await new Promise((resolve, reject) => state.jobs.set(job_id, { resolve, reject }));
    return result;
  }
  function jobDone(msg) {
    const w = state.jobs.get(msg.job_id);
    if (w) {
      state.jobs.delete(msg.job_id);
      return w.resolve(msg.result);
    }
    // A job of the connection before this one is not waited for: its result
    // is a file next to the manuscript or nothing at all.
    if (state.asking) state.jobsDone.set(msg.job_id, msg.result);
  }
  // A new connection never brings the answers of the old one: every wait
  // ends now, with an error.
  function dropJobs() {
    for (const w of state.jobs.values()) w.reject(new Error(CONNECTION_LOST));
    state.jobs.clear();
    state.jobsDone.clear();
  }

  // Every action of the tab goes through guard(): one at a time, in the order
  // pressed, and only once the daemon has confirmed the stored provider.
  let chain = Promise.resolve();
  const buttons = [];
  function setBusy(busy) {
    state.busy = busy;
    for (const b of buttons) b.disabled = busy;
  }
  async function run(fn) {
    try {
      const confirming = providersReady();
      if (confirming) {
        say("Waiting for the daemon to confirm the stored provider choice...");
        await confirming;
      }
      say("");
      setBusy(true);
      await fn();
    } catch (e) {
      say(e.message, true);
    } finally {
      setBusy(false);
    }
  }
  const guard = (fn) => () => (chain = chain.then(() => run(fn)));
  function onClick(id, fn) {
    const b = $(id);
    b.addEventListener("click", guard(fn));
    buttons.push(b);
  }

  function renderDoc() {
    $("tools-doc").textContent = docLine(state.doc);
    $("tools-policy-scope").checked = state.doc?.policy_packs?.review_scope !== false;
  }

  // ---- sections -------------------------------------------------------------

  // Selection: the selected text goes to the model, the answer is shown here.
  // No call writes to Word.
  async function selection(kind) {
    const text = String(await word.selectionText()).trim();
    if (!text) throw new Error("Select some text in the document first.");
    let r;
    if (kind === "translate") {
      const lang = $("tools-lang").value.trim();
      if (!lang) throw new Error("Name the language to translate into.");
      settings.toolsLang = lang;
      saveSettings(settings);
      say("Translating the selection...");
      r = await askJob("selection_translate", { text, lang });
    } else {
      const question = $("tools-question").value.trim();
      if (!question) throw new Error("Type a question about the selected text.");
      say("Asking about the selection...");
      r = await askJob("selection_ask", { text, question });
    }
    if (!r.ok) throw new Error(r.error || "The request failed.");
    say("");
    $("tools-selection").replaceChildren(
      el("div", {
        class: "review-reqs-line",
        text: selectionHead(kind, text, r, usedLabel("writing")),
      }),
      el("div", { class: "review-preview", text: r.text }),
    );
  }
  $("tools-lang").value = settings.toolsLang ?? "";
  onClick("tools-translate", () => selection("translate"));
  onClick("tools-ask", () => selection("ask"));

  // Journal guideline: candidates with the sentence they cite; the picked
  // ones are added to the ledger and the profile by the daemon.
  function renderRules() {
    const box = $("tools-rules");
    box.replaceChildren();
    const r = state.rules;
    $("tools-rules-add").hidden = !r?.candidates.length;
    if (!r) return;
    r.boxes = r.candidates.map((c) => {
      const pick = el("input", { type: "checkbox" });
      pick.checked = true;
      box.append(
        el(
          "label",
          { class: "review-ledger-item" },
          pick,
          el("div", { class: "review-reqs-line", text: candidateKind(c) }),
          el("div", { class: "review-preview", text: candidateCite(c) }),
        ),
      );
      return pick;
    });
    for (const w of r.warnings) box.append(el("div", { class: "response-check-warn", text: w }));
  }
  async function extractRules() {
    const text = $("tools-guideline").value.trim();
    if (!text) throw new Error("Paste the guideline text first.");
    say("Reading the guideline...");
    const r = await askJob("profile_rules_extract", { text });
    if (!r.ok) throw new Error(r.error || "The guideline could not be read.");
    state.rules = { text, candidates: r.candidates, warnings: r.warnings, boxes: [] };
    renderRules();
    const m = extractMessage(r);
    say(m.text, m.bad);
  }
  async function addRules() {
    const r = state.rules;
    const candidates = r.candidates.filter((_, i) => r.boxes[i].checked);
    if (!candidates.length) throw new Error("Pick at least one rule.");
    const done = await ask("profile_rules_apply", { text: r.text, candidates });
    state.rules = null;
    renderRules();
    say(rulesAdded(done));
  }
  // The review_scope policy pack of this paper's profile (rules 2 to 4 of a
  // review turn), saved as an override of the profile.
  async function savePolicy() {
    const on = $("tools-policy-scope").checked;
    const r = await ask("profile_set", { overrides: { policy_packs: { review_scope: on } } });
    state.doc = { ...state.doc, policy_packs: r.profile.policy_packs };
    renderDoc();
    say(
      on
        ? "Review turns keep to what the reviewer asked."
        : "Review turns no longer carry rules 2 to 4 for this paper.",
    );
  }
  onClick("tools-rules-extract", extractRules);
  onClick("tools-rules-add", addRules);
  $("tools-policy-scope").addEventListener("change", guard(savePolicy));
  renderRules();

  const lines = (rows) =>
    rows.map((x) =>
      el("div", { class: x.bad ? "response-check-fail" : "review-reqs-line", text: x.text }),
    );

  // Checks: the pane reads the text, the headings and the formatting; the
  // daemon answers findings. A finding with a paragraph is a button that
  // selects the paragraph. Nothing is changed.
  async function runChecks() {
    $("tools-checks").replaceChildren();
    say("Reading the document...");
    const snapshot = await word.collectReviewSnapshot();
    const headings = await word.collectHeadings();
    const format = await word.readFormatSnapshot();
    const r = await ask("format_check", { snapshot, headings, format });
    const rows = checkRows(r);
    $("tools-checks").replaceChildren(
      ...rows.map((x) => {
        if (x.index == null) return lines([x])[0];
        const go = el("button", { type: "button", class: "review-violation", text: x.text });
        go.addEventListener(
          "click",
          guard(() => word.selectParagraphByIndex(x.index)),
        );
        return go;
      }),
    );
    const found = rows.filter((x) => x.bad).length;
    say(found ? `${found} findings. Nothing was changed.` : "The checks found nothing.");
  }
  onClick("tools-check", runChecks);

  // Marked copy: the earlier version is read here and sent; the daemon
  // compares it with the saved file and writes the copy.
  async function buildMarked() {
    const f = $("tools-baseline").files?.[0];
    if (!f) throw new Error("Choose the earlier version (.docx) first.");
    if (!/\.docx$/i.test(f.name)) throw new Error(`${f.name} is not a .docx file.`);
    if (fileTooLarge(f.size)) throw new Error(`${f.name} is too large to send to the daemon.`);
    const mode = $("tools-marked-mode").value || "track_changes";
    $("tools-marked").replaceChildren();
    say(`Comparing the saved file with ${f.name}...`);
    const r = await askJob("marked_build", {
      mode,
      baseline_base64: bytesToBase64(new Uint8Array(await f.arrayBuffer())),
    });
    say("");
    $("tools-marked").replaceChildren(...lines(markedLines(r)));
  }
  onClick("tools-marked-build", buildMarked);

  // Anonymized copy: made by the daemon from the saved file, then verified.
  async function anonymize() {
    $("tools-anonymized").replaceChildren();
    say("Writing and checking the anonymized copy...");
    const r = await askJob("anonymize", {});
    if (!r.ok && !r.report) throw new Error(r.error || "The copy could not be written.");
    say("");
    $("tools-anonymized").replaceChildren(...lines(anonymizeLines(r)));
  }
  onClick("tools-anonymize", anonymize);

  renderDoc(); // loaded on "welcome" (taskpane.js calls reload once connected)

  return {
    onEvent(msg) {
      if (msg.event === "job_done") jobDone(msg);
    },
    // A new connection (welcome): the waits of the old one end with an error.
    async reload() {
      dropJobs();
      try {
        state.doc = await ask("tools_state");
      } catch (e) {
        state.doc = null;
        say(e.message, true);
      }
      renderDoc();
    },
  };
}
