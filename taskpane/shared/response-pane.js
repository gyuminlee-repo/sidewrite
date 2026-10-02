// Response tab (Word only). DOM glue: what a response cites, where it is in
// the manuscript, whether it is out of date and every check are decided by
// the daemon (daemon/response-service.mjs); the wording of the screen is in
// pure helpers (response-view.js). The pane reads Word (snapshot, headings,
// the rendered PDF) and never writes to it: the letter is a file of its own.
// Text from the daemon or the document is inserted with textContent only.

import {
  bytesToBase64,
  pdfTooLarge,
  staleText,
  translationMissing,
  shownCitations,
  citationText,
  checkLine,
  checksSummary,
  canExport,
  proposalRows,
  overridesFromForm,
  openRequirements,
  undecidedText,
  nextStep,
  exportMessage,
} from "./response-view.js";

const CONNECTION_LOST =
  "The connection to the daemon was lost before the result came. The saved state was loaded again; run it again if its result is missing.";
const EXPORT_LABEL = "Export letter (.docx)";
const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;

// " with Claude · Opus 5.5" for a result line, or "" when it is not known.
const withModel = (label) => (label ? ` with ${label}` : "");

export function initResponsePane({
  sendRequest,
  word,
  // () => the request that confirms the stored providers while it is under
  // way (providers-pane.js whenRestored), else null.
  providersReady = () => null,
  // (role) => what the last request of that kind of work really ran on
  // ("Claude · Opus 5.5"), or null when the pane does not know (taskpane.js).
  usedLabel = () => null,
}) {
  const root = document.getElementById("response-root");
  if (!root) return { onEvent() {}, reload: async () => {} }; // Excel pane: no Response tab
  const $ = (id) => document.getElementById(id);
  const state = {
    revision: null,
    profile: null, // the profile_get answer: { name, path, profile, overrides, available, problems }
    summary: null, // { fail, warn } of the last check; null once something changed
    checks: [],
    proposal: null, // what a loaded letter proposes, until it is applied
    output: "",
    texts: new Map(), // ref -> answer typed and not saved yet
    reasons: new Map(), // "<ref>" or "<ref>:<request id>" -> reason typed and not saved yet
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
    const m = $("response-message");
    m.textContent = text;
    m.classList.toggle("form-error", bad);
    m.hidden = !text;
  }
  async function ask(type, payload = {}) {
    const r = await sendRequest(type, payload);
    if (!r.ok) throw new Error(r.error || `${type} failed`);
    if (r.revision) state.revision = r.revision;
    return r;
  }
  // A job answers { accepted, job_id } at once; its result comes in a
  // job_done response_event, possibly before the answer.
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
    if (!result.ok) throw new Error(result.error || `${type} failed`);
    if (result.revision) state.revision = result.revision;
    return result;
  }
  function jobDone(msg) {
    const w = state.jobs.get(msg.job_id);
    if (w) {
      state.jobs.delete(msg.job_id);
      return w.resolve(msg.result);
    }
    if (state.asking) return state.jobsDone.set(msg.job_id, msg.result);
    // Nobody waits for it: a job of the connection before this one. The
    // daemon saved what it did; show it instead of dropping it.
    if (msg.result?.ok && msg.result.revision) {
      state.revision = msg.result.revision;
      changed();
      say("A job finished after the connection came back; its result is shown.");
      render();
    }
  }
  // A new connection never brings the answers of the old one: every wait
  // ends now, with an error. Returns how many were waiting.
  function dropJobs() {
    const n = state.jobs.size + state.asking;
    for (const w of state.jobs.values()) w.reject(new Error(CONNECTION_LOST));
    state.jobs.clear();
    state.jobsDone.clear();
    return n;
  }

  // Every action of the tab goes through guard(): one at a time, in the order
  // pressed, and only once the daemon has confirmed the stored provider (right
  // after a (re)connect it has not yet, and would have to guess).
  let chain = Promise.resolve();
  let actionButtons = [];
  function setBusy(busy) {
    state.busy = busy;
    for (const b of actionButtons) b.disabled = busy;
  }
  async function run(fn, args) {
    try {
      const confirming = providersReady();
      if (confirming) {
        say("Waiting for the daemon to confirm the stored provider choice...");
        await confirming;
      }
      say("");
      setBusy(true);
      await fn(...args);
    } catch (e) {
      say(e.message, true);
    } finally {
      state.busy = false;
    }
    render();
  }
  const guard =
    (fn) =>
    (...args) =>
      (chain = chain.then(() => run(fn, args)));
  // An action on a button: the button is disabled while any action runs.
  function action(label, fn, cls = "btn-secondary btn-small") {
    const b = el("button", { type: "button", class: cls, text: label });
    b.addEventListener("click", guard(fn));
    b.disabled = state.busy;
    actionButtons.push(b);
    return b;
  }

  // The manuscript as it is now. Sent with every request that computes from
  // it; the daemon checks the headings against the snapshot.
  const docNow = async () => ({
    snapshot: await word.collectReviewSnapshot(),
    headings: await word.collectHeadings(),
  });
  // What the last check said is about the state before this change.
  function changed() {
    state.summary = null;
    state.checks = [];
  }
  const comments = () => state.revision?.comments ?? [];
  const letter = () => state.profile?.profile?.response_letter ?? null;

  // ---- rendering ----------------------------------------------------------
  function renderProfile() {
    const p = state.profile;
    const select = $("response-profile");
    select.replaceChildren(
      ...(p?.available ?? []).map((name) => el("option", { value: name, text: name })),
    );
    select.value = p?.name ?? "";
    $("response-profile-path").textContent = p?.path ?? "";
    const over = p?.overrides ?? {};
    $("response-lang").value = over.response_letter?.lang ?? "";
    $("response-lang").setAttribute("placeholder", letter()?.lang ?? "English");
    $("response-rationale-lang").value = over.rationale_lang ?? "";
    $("response-rationale-lang").setAttribute(
      "placeholder",
      p?.profile?.rationale_lang ?? "English",
    );
    $("response-gloss").checked = !!letter()?.gloss_enabled;
    $("response-gloss-lang").value = over.response_letter?.gloss_lang ?? "";
    $("response-gloss-lang").setAttribute("placeholder", letter()?.gloss_lang ?? "");
  }

  function renderProposal() {
    const box = $("response-proposal");
    box.replaceChildren();
    const p = state.proposal;
    $("response-proposal-apply").hidden = !p || !(p.rows.length || p.previous.length);
    if (!p) return;
    for (const r of p.rows)
      box.append(
        el("div", {
          class: "response-proposal-row",
          text: `${r.key}: ${r.current} -> ${r.proposed}${r.evidence ? ` (${r.evidence})` : ""}`,
        }),
      );
    for (const note of p.notes)
      box.append(el("small", { class: "response-proposal-note", text: note }));
  }

  // A reason box whose text survives a render.
  function reasonInput(key, cls) {
    const input = el("input", {
      type: "text",
      class: cls,
      placeholder: "Reason (printed in the letter)",
    });
    input.value = state.reasons.get(key) ?? "";
    input.addEventListener("input", () => state.reasons.set(key, input.value));
    return input;
  }
  function typedReason(key) {
    const reason = String(state.reasons.get(key) ?? "").trim();
    if (!reason) throw new Error("Write the reason first: it is printed in the letter.");
    return reason;
  }

  function card(c) {
    const box = el(
      "div",
      { class: "response-card", "data-ref": c.ref },
      el("strong", { class: "response-ref", text: c.ref }),
      el("div", { class: "response-comment", text: c.text }),
    );
    const undecided = undecidedText(c);
    if (undecided) box.append(el("div", { class: "response-undecided", text: undecided }));

    if (c.declined) {
      box.append(
        el("div", { class: "response-declined", text: `Declined: ${c.declined.reason}` }),
        action("Undo decline", async () => {
          await ask("response_save", { ref: c.ref, declined: null });
          changed();
          say(`${c.ref} is open again.`);
        }),
      );
      return box;
    }

    const stale = staleText(c.response?.stale);
    if (stale) box.append(el("div", { class: "response-stale", text: stale }));
    const text = el("textarea", { rows: "4", class: "response-text" });
    // A render rebuilds this box: unsaved text lives in state and is dropped
    // only when it is saved or a draft takes its place.
    text.value = state.texts.has(c.ref) ? state.texts.get(c.ref) : (c.response?.text ?? "");
    text.addEventListener("input", () => state.texts.set(c.ref, text.value));
    box.append(text);
    if (c.response?.gloss)
      box.append(el("div", { class: "response-gloss", text: c.response.gloss }));
    // The daemon drops the translation of an answer saved with another text.
    const untranslated = translationMissing(c, letter());
    if (untranslated)
      box.append(el("div", { class: "response-gloss-missing", text: untranslated }));
    if (c.response?.previous && !c.response?.text)
      box.append(
        el("small", {
          class: "response-previous",
          text: "An answer from an earlier letter is kept and is given to the draft.",
        }),
      );
    box.append(
      el(
        "div",
        { class: "response-actions" },
        action(
          "Draft",
          async () => {
            await draft(c.ref);
          },
          "btn-primary btn-small",
        ),
        action("Save", async () => {
          const r = await ask("response_save", {
            ref: c.ref,
            text: text.value,
            // The citations this card shows: the text is judged by them.
            shown: shownCitations(c),
            ...(await docNow()),
          });
          state.texts.delete(c.ref);
          changed();
          // A warning: the manuscript changed under the answer (it is marked).
          const warnings = r.warnings ?? [];
          say([`Saved ${c.ref}.`, ...warnings].join(" "), warnings.length > 0);
        }),
        // Of the answer as saved; the answer itself is not rewritten.
        untranslated &&
          action("Translate", async () => {
            say(`Translating ${c.ref}...`);
            await askJob("response_translate", { ref: c.ref });
            changed();
            say(`Translated ${c.ref}${withModel(usedLabel("writing"))}.`);
          }),
      ),
    );

    // The requests no accepted change and no adopted figure change answers:
    // each is cited, or declined with a reason the letter prints.
    const declined = c.response?.declined_requirements ?? [];
    const saveDeclined = (list) =>
      ask("response_save", { ref: c.ref, declined_requirements: list });
    for (const q of openRequirements(c)) {
      const key = `${c.ref}:${q.id}`;
      box.append(
        el(
          "div",
          { class: "response-req" },
          el("span", {
            text: `${q.id} ${q.text}: no accepted change${
              q.pending
                ? " (a change is pending in Word)"
                : q.figure
                  ? " (a figure or file waits for your decision on the figures page)"
                  : ""
            }`,
          }),
          reasonInput(key, "response-req-reason"),
          action("Decline", async () => {
            await saveDeclined([...declined, { id: q.id, reason: typedReason(key) }]);
            state.reasons.delete(key);
            changed();
            say(`Declined ${q.id} of ${c.ref}.`);
          }),
        ),
      );
    }
    for (const d of declined) {
      box.append(
        el(
          "div",
          { class: "response-req response-req-declined" },
          el("span", { text: `${d.id} declined: ${d.reason}` }),
          action("Undo", async () => {
            await saveDeclined(declined.filter((x) => x.id !== d.id));
            changed();
            say(`${d.id} of ${c.ref} is open again.`);
          }),
        ),
      );
    }

    // What the letter will print under the answer. Computed by the daemon.
    for (const cit of c.response?.citations ?? []) {
      const shown = letter()
        ? citationText(cit, letter())
        : { line: cit.problem ? `Not located: ${cit.problem}` : "", quote: cit.quote || null };
      const row = el("div", { class: "response-citation" });
      if (shown.line) row.append(el("div", { class: "response-citation-line", text: shown.line }));
      if (shown.quote)
        row.append(el("div", { class: "response-citation-quote", text: shown.quote }));
      box.append(row);
    }

    const key = c.ref;
    box.append(
      el(
        "div",
        { class: "response-decline" },
        reasonInput(key, "response-decline-reason"),
        action("Decline this comment", async () => {
          await ask("response_save", { ref: c.ref, declined: { reason: typedReason(key) } });
          state.reasons.delete(key);
          changed();
          say(`Declined ${c.ref}.`);
        }),
      ),
    );
    return box;
  }

  function render() {
    actionButtons = STATIC_BUTTONS.map((id) => $(id));
    renderProfile();
    renderProposal();
    $("response-comments").replaceChildren(...comments().map(card));
    $("response-next").textContent = nextStep(state.revision, state.summary).text;
    $("response-summary").textContent = state.summary ? checksSummary(state.summary) : "";
    $("response-checks").replaceChildren(
      ...state.checks.map((row) =>
        el("div", {
          class: `response-check-row response-check-${row.level}`,
          text: checkLine(row),
        }),
      ),
    );
    // "Export anyway" only after a check said what fails.
    const exportLabel = state.summary ? canExport(state.summary).label : EXPORT_LABEL;
    $("response-export").textContent = exportLabel;
    const last = state.revision?.response_export?.path;
    $("response-output").textContent = state.output || (last ? `Last export: ${last}` : "");
    setBusy(state.busy);
  }

  // ---- actions ------------------------------------------------------------
  async function draft(ref) {
    say(`Drafting ${ref}...`);
    const r = await askJob("response_draft", { ref, ...(await docNow()) });
    state.texts.delete(ref);
    changed();
    const fails = (r.checks ?? []).filter((x) => x.level === "fail");
    say(
      [
        `Drafted ${ref}${withModel(usedLabel("writing"))}.`,
        fails.length ? `${plural(fails.length, "check")} fail for it: ${fails[0].detail}.` : "",
        ...(r.warnings ?? []),
      ]
        .filter(Boolean)
        .join(" "),
      fails.length > 0,
    );
  }

  async function draftAll() {
    const refs = comments()
      .filter((c) => !c.declined && !c.response?.text?.trim())
      .map((c) => c.ref);
    if (!refs.length) return say("Nothing to draft: every comment has a response or is declined.");
    let done = 0;
    for (const ref of refs) {
      try {
        await draft(ref);
      } catch (e) {
        throw new Error(`Stopped at ${ref} after ${plural(done, "draft")}: ${e.message}`);
      }
      done += 1;
    }
    say(`Drafted ${plural(done, "response")}${withModel(usedLabel("writing"))}.`);
  }

  async function refresh() {
    await ask("response_state", await docNow());
    changed();
    say(`Refreshed from the manuscript: ${plural(comments().length, "comment")}.`);
  }

  async function locate() {
    const doc = await docNow();
    say("Rendering the manuscript to PDF in Word...");
    const pdf = await word.exportRenderPdf();
    if (pdfTooLarge(pdf.pdf_base64.length))
      throw new Error(
        `The PDF of the manuscript is too large to send to the daemon (${Math.round(pdf.byte_length / 1048576)} MB).`,
      );
    say("Reading the line numbers...");
    const r = await askJob("response_locate", { ...doc, pdf_base64: pdf.pdf_base64 });
    changed();
    if (!r.lines_indexed)
      throw new Error(
        "No line numbers were found in the PDF. Turn on continuous line numbers in Word (Layout > Line Numbers > Continuous), then update again.",
      );
    say(`Line numbers updated: ${r.located} quotes located.`);
  }

  async function check() {
    say("Checking...");
    const r = await askJob("response_check", await docNow());
    state.summary = r.summary;
    state.checks = r.checks;
    say(`Checked: ${checksSummary(r.summary)}`, r.summary.fail > 0);
  }

  async function exportLetter(format) {
    // Asked for on purpose: the last check failed and the button said so.
    const force = (state.summary?.fail ?? 0) > 0;
    say("Exporting...");
    const r = await askJob("response_export", { ...(await docNow()), format, force });
    state.summary = r.summary;
    state.checks = r.checks;
    const text = exportMessage(r);
    if (r.exported) state.output = text;
    say(text, !r.exported);
  }

  async function saveProfile() {
    const r = await ask("profile_set", {
      journal_profile: $("response-profile").value,
      overrides: overridesFromForm({
        lang: $("response-lang").value,
        rationaleLang: $("response-rationale-lang").value,
        gloss: $("response-gloss").checked,
        glossLang: $("response-gloss-lang").value,
      }),
    });
    takeProfile(r);
    changed();
    say(`Saved for this paper: profile ${r.name}.`);
  }
  // profile_set answers with the profile as it is now.
  function takeProfile(r) {
    state.profile = {
      ...state.profile,
      name: r.name,
      path: r.path,
      profile: r.profile,
      problems: r.problems ?? [],
      overrides: r.revision?.profile_overrides ?? state.profile?.overrides ?? {},
    };
  }

  async function loadLetter() {
    const input = $("response-letter-file");
    const f = input.files?.[0];
    input.value = ""; // the same file can be picked again
    if (!f) return;
    if (!/\.docx$/i.test(f.name)) throw new Error(`${f.name} is not a .docx file.`);
    if (pdfTooLarge(Math.ceil(f.size / 3) * 4))
      throw new Error(`${f.name} is too large to send to the daemon.`);
    say(`Reading ${f.name}...`);
    const r = await askJob("profile_propose", {
      docx_base64: bytesToBase64(new Uint8Array(await f.arrayBuffer())),
    });
    const rows = proposalRows(r, state.profile?.profile);
    const previous = Object.keys(r.previous ?? {});
    state.proposal = {
      overrides: r.overrides ?? {},
      previousByRef: r.previous ?? {},
      rows,
      previous,
      notes: r.evidence?.notes ?? [],
    };
    say(
      `Read the letter: ${rows.length} ${rows.length === 1 ? "setting differs" : "settings differ"}, ${plural(previous.length, "earlier answer")} found. Nothing is saved until you press Apply to this paper.`,
    );
  }

  async function applyProposal() {
    const p = state.proposal;
    if (!p) return;
    if (p.rows.length) takeProfile(await ask("profile_set", { overrides: p.overrides }));
    let unknown = [];
    if (p.previous.length)
      ({ unknown = [] } = await ask("response_import_previous", { previous: p.previousByRef }));
    state.proposal = null;
    changed();
    say(
      `Applied to this paper: ${plural(p.rows.length, "setting")}, ${plural(p.previous.length - unknown.length, "earlier answer")} kept for drafting.${
        unknown.length ? ` Not matched to a comment: ${unknown.join(", ")}.` : ""
      }`,
    );
  }

  // The buttons of the page (index.html); those of the cards are made in card().
  const STATIC_BUTTONS = [
    "response-profile-save",
    "response-proposal-apply",
    "response-refresh",
    "response-draft-all",
    "response-locate",
    "response-check",
    "response-export",
    "response-export-text",
  ];
  $("response-profile-save").addEventListener("click", guard(saveProfile));
  $("response-letter-file").addEventListener("change", guard(loadLetter));
  $("response-proposal-apply").addEventListener("click", guard(applyProposal));
  $("response-refresh").addEventListener("click", guard(refresh));
  $("response-draft-all").addEventListener("click", guard(draftAll));
  $("response-locate").addEventListener("click", guard(locate));
  $("response-check").addEventListener("click", guard(check));
  $("response-export").addEventListener(
    "click",
    guard(() => exportLetter("docx")),
  );
  $("response-export-text").addEventListener(
    "click",
    guard(() => exportLetter("text")),
  );

  // The saved state: no call into Word.
  async function load() {
    state.profile = await ask("profile_get");
    await ask("review_load");
    if (state.profile.problems?.length)
      say(`Profile problem: ${state.profile.problems.join("; ")}`, true);
  }

  render(); // loaded on "welcome" (taskpane.js calls reload once connected)

  return {
    onEvent(msg) {
      if (msg.event === "job_done") jobDone(msg);
    },
    // A new connection (welcome). The waits of the old one end with an error
    // (a job_done of the old connection never comes), so no progress line
    // stays. The saved state is loaded; when a job was cut off, the state is
    // also asked for against the manuscript, because what that job saved is
    // not known here.
    async reload() {
      const lost = dropJobs();
      try {
        await load();
        if (lost && comments().length) await ask("response_state", await docNow());
      } catch (e) {
        say(e.message, true);
      }
      if (lost) {
        changed();
        say(CONNECTION_LOST, true);
      }
      render();
    },
  };
}
