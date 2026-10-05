// Review tab (Word only). DOM glue: every decision lives in the daemon
// (daemon/review-service.mjs) or in pure helpers (review-view.js). Office.js
// actions that need no agent (snapshot, jump to a comment, comment cleanup)
// run here directly. Text from the daemon or the document is inserted with
// textContent only.

import {
  STATE_LABELS,
  STATUS_LABELS,
  countComments,
  cleanupReport,
  buildCommentsBackup,
  compareCommentLists,
  deleteAllReport,
  invariantReport,
  ledgerItemFromForm,
  commentsToResolve,
  changeFlags,
  requirementsFromText,
  consistencyMessage,
  consistencyLines,
  openRequestsText,
  doneBlockers,
  commentNotices,
  ledgerDraftFromChange,
  ledgerFormFromItem,
  ledgerCheckFromText,
  ledgerStatusText,
  pendingDecisions,
  invariantText,
  inFlightEdit,
  reviewerOfRef,
  progressText,
  countsSentence,
  ledgerHowText,
  requestStateText,
  requirementStates,
  clockText,
  commentState,
  firstLine,
  summaryText,
  reviewLine,
  reviewTextFromFile,
  REVIEW_FILE_MAX_BYTES,
  ledgerImportFromJson,
  candidateRule,
  candidateLabel,
  queueRefs,
  figuresUrl,
  hasFigureWork,
  SUBMISSION_STATE_LABELS,
  submissionCount,
  submissionGoLabel,
  UNDO_TOOK_MORE_TEXT,
  joinLabel,
  linkedFilesText,
  partOfText,
  editedInText,
  changesElsewhere,
  KEEP_WORD_OPEN,
  MINIMIZED_IN_TURN,
} from "./review-view.js";
import { createProposeQueue } from "./propose-queue.js";
import { stillPresentItems } from "./residue.js";
import { firstWords, paragraphLabel } from "./paragraph-label.js";
import { createAutoRefresh } from "./auto-refresh.js";

const UNKNOWN_HINT =
  "Word did not answer; this exception may already be applied. Refresh status, then check the paragraph.";
// An override waits for the Word tool, which may take up to the daemon's 60 s
// tool timeout (bridge.mjs TOOL_TIMEOUT_MS). A split or a ledger check is a
// job instead (askJob): no request timeout covers its model calls.
const OVERRIDE_WAIT = 120_000;
const CONNECTION_LOST =
  "The connection to the daemon was lost before the result came. Run it again.";
const APPLYING_HINT = "Applying this exception in Word. Wait for it to finish.";
// "Propose all": the longest wait for the end of one review turn.
const TURN_WAIT_MS = 30 * 60_000;
// How long the pane waits for the answer of review_propose (taskpane.js
// sendRequest waits as long by default); the request says so (wait_ms).
const PROPOSE_WAIT_MS = 10_000;
const PROVIDER_WAIT = "Waiting for the daemon to confirm the stored provider choice...";
const FIGURES_LABEL = "Figures and other files";

// Opens an address in the browser of the computer from inside an Office task
// pane (Office.context.ui.openBrowserWindow, requirement set
// OpenBrowserWindowApi 1.1: Word on Windows, Mac and iPad, not on the web).
// true: Office took it. false: no Office here or no such API; the link is
// then followed as a normal link.
export function openInBrowser(url, office = globalThis.Office) {
  try {
    if (!office?.context?.requirements?.isSetSupported?.("OpenBrowserWindowApi", "1.1"))
      return false;
    office.context.ui.openBrowserWindow(url);
    return true;
  } catch {
    return false;
  }
}

export function initReviewPane({
  sendRequest,
  word,
  onTurnStarted,
  isTurnInFlight,
  // () => the request that confirms the stored providers while it is under
  // way (providers-pane.js whenRestored), else null.
  providersReady = () => null,
  // (role) => what the last request of that kind of work really ran on
  // ("Claude · Opus 5.5"), or null when the pane does not know (taskpane.js).
  usedLabel = () => null,
  // () => stops the running turn (the Stop button of the chat).
  stopTurn = () => {},
  // () => the text of the last answer of the model in the chat, or null. It is
  // the reason a turn gives when it proposes no change.
  lastReply = () => null,
  // () => the document as this pane named it to the daemon (hello
  // active_doc), or null. The figures page finds the manuscript by it.
  docUrl = () => null,
  // () => where the daemon serves this pane, and with it the figures page.
  daemonOrigin = () => globalThis.location?.origin ?? "",
  // (tab, elementId) => shows that tab of the pane and brings the element
  // into view (taskpane.js, where the active tab is set). A row of the
  // submission checklist is fixed there.
  goTo = () => {},
  // (type, payload) => the result of a job of the Tools tab's service
  // (tools-pane.js job): its job_done comes as a tools_event, which that
  // pane waits for. The submission checklist is such a job.
  toolsJob = null,
  // () => called after a refresh in a supplementary file wrote what the
  // letter cites from it (taskpane.js: the Response tab reads it again).
  onPartRefreshed = () => {},
}) {
  const root = document.getElementById("review-root");
  if (!root) return { onEvent() {}, setTurnBusy() {} }; // Excel pane: no Review tab
  const $ = (id) => document.getElementById(id);
  const state = {
    revision: null,
    path: null,
    draft: [],
    backupShown: null,
    editingLedger: null, // L6: id of the ledger item loaded into the form
    confirmDone: null, // L5: ref whose "Done" was asked once with something left open
    answerDrafts: new Map(), // ref/question id -> answer typed and not saved yet
    reqDrafts: new Map(), // ref -> requirement text typed and not saved yet (Codex review 9)
    reqEdit: new Set(), // refs whose list of requests is open for editing
    previewRef: null, // the comment whose prompt is shown under its Advanced
    armedBlocked: new Set(), // ids of blocked edits whose "Allow exception" was pressed once
    armedReapply: new Set(), // ids of outcome_unknown edits whose "Re-apply anyway" was pressed once
    applyingBlocked: new Set(), // ids of blocked edits whose override request is in flight
    turn: null, // L9: { ref, startedAt } of the running review turn
    turnsDone: 0, // turn_done events seen, to spot one that beats its review_propose answer
    timer: null,
    jobs: new Map(), // job_id -> { resolve, reject } of an askJob waiting for job_done
    jobsDone: new Map(), // job_id -> result of a job_done that came before its answer
    candidates: [], // P15: ledger rule candidates of the last suggestion, each with `picked`
    turnStart: new Map(), // ref -> { changes, blocked } on record when its turn was asked for
    turnWaiters: new Map(), // P15: ref -> { resolve, timer } of a "Propose all" turn
    queueNote: null, // P15: "Propose all: 2/5 R1.2." while the queue runs
    reviewTurnOpen: false, // P15: a review turn was asked for and its turn_done has not come
    hiddenInTurn: false, // the page was hidden (Word minimized?) while a review turn was open
    openRef: null, // the comment whose card is open (one at a time)
    pasteOpen: false, // the paste box is shown although a review is split
    results: new Map(), // ref -> { text, bad, noChange, reason } of its last turn in this session
    notices: new Map(), // ref -> a sentence for the author that outlives the result of the turn
    cards: new Map(), // ref -> the elements of its card, as of the last render
    submission: null, // { rows, doc, unsaved, at } of the last "Check now" of this connection
    linked: null, // { part_of, parts, note } of the last load (null: a daemon from before links)
    linkCandidates: [], // manuscripts in the folder this file could join
  };
  // The manuscript this file belongs to, when it is a supplementary file. What
  // the manuscript owns (comment list, requests, rules) is then shown and not
  // changed here; changes are proposed here for this file as in a manuscript.
  const partOf = () => state.linked?.part_of ?? null;
  // The changes of a comment in the other files of the revision, as the last
  // load gave them: where its requests stand there.
  const elsewhere = (ref) => changesElsewhere(state.linked, ref);

  function el(tag, props = {}, ...kids) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(props)) {
      if (k === "text") n.textContent = v;
      else if (k === "on") for (const [ev, fn] of Object.entries(v)) n.addEventListener(ev, fn);
      else n.setAttribute(k, v);
    }
    for (const k of kids) if (k) n.append(k);
    return n;
  }
  // The way to the figures page of this manuscript, at the comment `ref` when
  // one is given: a link that Office opens in the browser where it can and
  // that is a normal link elsewhere, with its address under it to copy when
  // nothing opens.
  function figuresLink(ref = null) {
    const url = figuresUrl(daemonOrigin(), docUrl(), ref);
    const a = el("a", {
      class: "btn-secondary btn-small review-figures-open",
      href: url,
      target: "_blank",
      rel: "noopener",
      text: FIGURES_LABEL,
    });
    a.addEventListener("click", (ev) => {
      if (openInBrowser(url)) ev?.preventDefault?.();
    });
    return el(
      "div",
      { class: "review-figures", ...(ref ? { "data-figures-ref": ref } : {}) },
      a,
      el("div", { class: "review-figures-address", text: url }),
    );
  }
  // The message line as it was last written, for a line put in front of it.
  let said = { text: "", bad: false };
  // details: the exact numbers behind the sentence, under "Details".
  function say(text, bad = false, details = "") {
    said = { text, bad };
    const m = $("review-message");
    m.textContent = text;
    m.classList.toggle("form-error", bad);
    m.hidden = !text;
    $("review-message-details").textContent = text ? details : "";
    $("review-message-more").hidden = !text || !details;
  }
  // A paragraph, named by its section and its first words. Pressing it selects
  // the paragraph in Word. The internal id of a paragraph is never shown.
  function paragraphLink(index) {
    const b = el("button", {
      type: "button",
      class: "btn-link review-paragraph-link",
      text: paragraphLabel(state.revision?.snapshot, index).label,
    });
    b.title = "Select this paragraph in Word";
    b.addEventListener(
      "click",
      guard(() => word.selectParagraphByIndex(index)),
    );
    return b;
  }
  async function ask(type, payload = {}, timeoutMs) {
    const r = await sendRequest(type, payload, timeoutMs ? { timeoutMs } : undefined);
    if (!r.ok) throw new Error(r.error || `${type} failed`);
    if (r.revision) state.revision = r.revision;
    return r;
  }
  // review_split and ledger_check answer { accepted, job_id } at once; the
  // result comes in a job_done review_event, possibly before the answer.
  async function askJob(type, payload) {
    const { job_id } = await ask(type, payload);
    let result = state.jobsDone.get(job_id);
    state.jobsDone.delete(job_id);
    result ??= await new Promise((resolve, reject) => state.jobs.set(job_id, { resolve, reject }));
    if (!result.ok) throw new Error(result.error || `${type} failed`);
    if (result.revision) state.revision = result.revision;
    return result;
  }
  function jobDone(msg) {
    const w = state.jobs.get(msg.job_id);
    if (!w) return state.jobsDone.set(msg.job_id, msg.result);
    state.jobs.delete(msg.job_id);
    w.resolve(msg.result);
  }
  // A new connection never brings the job_done of the old one.
  function dropJobs() {
    for (const w of state.jobs.values()) w.reject(new Error(CONNECTION_LOST));
    state.jobs.clear();
    state.jobsDone.clear();
  }
  const guard =
    (fn) =>
    async (...a) => {
      try {
        // Right after a (re)connect the daemon has not confirmed the stored
        // provider yet: the action waits for that instead of being sent to a
        // daemon that would have to guess.
        const confirming = providersReady();
        if (confirming) {
          say(PROVIDER_WAIT);
          await confirming;
        }
        say("");
        await fn(...a);
      } catch (e) {
        say(e.message, true);
      }
      render();
    };
  // An action that writes to Word runs once at a time: a press while it
  // runs does nothing (two runs at once made two rationale comments, one of
  // them recorded nowhere). button(): the element to show as disabled, as it
  // is in the page at that moment (a render replaces the buttons of a card).
  const writing = new Set();
  const once =
    (key, fn, button = () => null) =>
    async (...a) => {
      if (writing.has(key)) return;
      writing.add(key);
      const pressed = button();
      if (pressed) pressed.disabled = true;
      try {
        await fn(...a);
      } finally {
        writing.delete(key);
        const now = button();
        if (now) now.disabled = false;
      }
    };
  const rationaleButtons = new Map(); // change id -> its button in the page

  // ---- rendering ----------------------------------------------------------
  // "Propose changes" buttons of the last render; setTurnBusy flips them when
  // any turn of the pane starts or ends (turn-state.js).
  let proposeButtons = [];
  function renderDraft() {
    const box = $("review-draft");
    box.replaceChildren();
    state.draft.forEach((c, i) => {
      const ref = el("input", { type: "text", value: c.ref, class: "review-ref" });
      ref.addEventListener("input", () => (c.ref = ref.value.trim()));
      const body = el("textarea", { rows: "3" });
      body.value = c.text;
      body.addEventListener("input", () => (c.text = body.value));
      // The part of the review the comment is under ("Abstract", "Round 2").
      // The split takes that label off its line, so it is shown and kept here.
      const section = el("input", {
        type: "text",
        class: "review-section",
        placeholder: "Section (optional)",
      });
      section.value = c.section ?? "";
      section.addEventListener("input", () => (c.section = section.value));
      // L5: what the comment asks for, one per line (empty after a regex split).
      const reqs = el("textarea", {
        rows: "2",
        class: "review-reqs",
        placeholder: "What it asks for, one per line",
      });
      reqs.value = c.requirementsText ?? "";
      reqs.addEventListener("input", () => (c.requirementsText = reqs.value));
      const drop = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Remove",
      });
      drop.addEventListener("click", () => {
        state.draft.splice(i, 1);
        render();
      });
      box.append(el("div", { class: "review-draft-row" }, ref, section, body, reqs, drop));
    });
    $("review-draft-actions").hidden = state.draft.length === 0;
  }

  function fillLedgerForm(f) {
    $("ledger-kind").value = f.kind;
    $("ledger-text").value = f.text;
    $("ledger-scope").value = f.scope;
    $("ledger-check-type").value = f.checkType;
    $("ledger-pattern").value = f.pattern;
    $("ledger-flags").value = f.flags;
    $("ledger-expect").value = f.expect;
  }

  function changeRow(ch) {
    const row = el(
      "button",
      { type: "button", class: `review-change state-${ch.state}` },
      el("span", { class: "review-badge", text: STATE_LABELS[ch.state] ?? ch.state }),
      el("span", {
        text: ` ${ch.kind}: ${ch.rationale.what}`,
      }),
    );
    row.title = `${ch.rationale.why} / ${ch.rationale.addresses}\nBefore: ${ch.before}\nAfter: ${ch.after}`;
    row.addEventListener(
      "click",
      guard(async () => {
        if (!ch.word_comment_id) throw new Error("This change has no comment to jump to.");
        await word.selectCommentById(ch.word_comment_id);
      }),
    );
    // L1-L3 flags next to the row; L6 "Add to ledger" on a rejected change.
    const wrap = el("div", { class: "review-change-wrap" }, row);
    for (const f of changeFlags(ch)) {
      wrap.append(el("span", { class: `review-flag review-flag-${f.kind}`, text: f.text }));
    }
    // Applied without its rationale comment: the comment can be added now.
    if (ch.comment_missing && ch.word_comment_id == null) {
      const add = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Add rationale comment",
      });
      const key = `rationale ${ch.id}`;
      add.disabled = writing.has(key);
      rationaleButtons.set(ch.id, add);
      add.addEventListener(
        "click",
        once(
          key,
          guard(async () => {
            const ref = ch.id.slice(0, ch.id.lastIndexOf("-"));
            const t = await ask("review_comment_text", { ref, change_id: ch.id });
            const r = await word.addChangeComment({
              ...t.change,
              comment_text: t.comment_text,
              known_comment_ids: t.known_comment_ids,
            });
            if (!r.ok)
              throw new Error(`${ch.id}: ${r.message || r.code || "the comment was not added"}`);
            await ask("review_comment_set", {
              ref,
              change_id: ch.id,
              word_comment_id: r.word_comment_id,
            });
            say(`${ch.id}: rationale comment added.`);
          }),
          () => rationaleButtons.get(ch.id),
        ),
      );
      wrap.append(add);
    }
    if (ch.state === "rejected") {
      const draft = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Add to ledger...",
      });
      draft.addEventListener(
        "click",
        guard(() => {
          state.editingLedger = null;
          fillLedgerForm(ledgerDraftFromChange(ch));
          say(`Draft rule from ${ch.id} is in the ledger form. Edit it, then Add.`);
        }),
      );
      wrap.append(draft);
    }
    return wrap;
  }

  // One override request per blocked edit; its buttons stay disabled until
  // the daemon answers (the daemon refuses an overlapping one too).
  async function override(ref, b, extra = {}) {
    state.applyingBlocked.add(b.id);
    render();
    try {
      const r = await ask("review_override", { ref, blocked_id: b.id, ...extra }, OVERRIDE_WAIT);
      if (r.outcome_unknown) throw new Error(UNKNOWN_HINT);
      if (!r.applied) throw new Error(r.result?.message || "Not applied");
    } finally {
      state.applyingBlocked.delete(b.id);
    }
  }

  function blockedRow(ref, b) {
    const newText = b.args.replace ?? b.args.insert_after ?? "";
    // No forbid rule: a normal edit Word did not answer. It is no exception,
    // so it has no Allow button, only the re-apply below.
    const own = b.blocked_by.length === 0;
    const label = el("span", {
      text: own
        ? `No answer from Word: "${newText}" `
        : `Blocked by ${b.blocked_by.join(", ")}: "${newText}" `,
    });
    if (b.state === "applying" || state.applyingBlocked.has(b.id)) {
      const busy = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Applying...",
      });
      busy.disabled = true;
      return el(
        "div",
        { class: "review-blocked" },
        label,
        el("span", { class: "review-unknown", text: APPLYING_HINT + " " }),
        busy,
      );
    }
    // The armed state lives in `state`: guard() re-renders after every
    // click, which rebuilds this button.
    const armed = state.armedBlocked.has(b.id);
    const allow = el("button", {
      type: "button",
      class: "btn-secondary btn-small",
      text: armed ? `Apply despite ${b.blocked_by.join(", ")}?` : "Allow exception",
    });
    allow.addEventListener(
      "click",
      guard(async () => {
        if (!state.armedBlocked.has(b.id)) {
          state.armedBlocked.add(b.id);
          return;
        }
        state.armedBlocked.delete(b.id);
        await override(ref, b);
      }),
    );
    if (b.state !== "outcome_unknown") {
      return el("div", { class: "review-blocked" }, label, allow);
    }
    // Word did not answer the exception, so it may already be applied: a
    // plain second Allow could apply it twice. The daemon refuses it too.
    allow.disabled = true;
    const armedAgain = state.armedReapply.has(b.id);
    const reapply = el("button", {
      type: "button",
      class: "btn-secondary btn-small",
      text: armedAgain ? "Re-apply anyway? It may apply twice. Press again" : "Re-apply anyway",
    });
    reapply.addEventListener(
      "click",
      guard(async () => {
        if (!state.armedReapply.has(b.id)) {
          state.armedReapply.add(b.id);
          return;
        }
        state.armedReapply.delete(b.id);
        await override(ref, b, { force: true });
      }),
    );
    return el(
      "div",
      { class: "review-blocked" },
      label,
      el("span", {
        class: "review-unknown",
        text: (own ? UNKNOWN_HINT.replace("this exception", "this edit") : UNKNOWN_HINT) + " ",
      }),
      ...(own ? [] : [allow]),
      reapply,
    );
  }

  // The paragraphs that go into the prompt of a comment: the selection in
  // Word when "send selected paragraphs" is on, else none. Preview prompt
  // and Propose changes both use this, so the preview shows the prompt the
  // turn will get.
  async function promptParagraphs() {
    return $("review-use-selection").checked ? await word.selectionParagraphs() : [];
  }

  // ---- the list: a row for each comment, its card when it is open ----------------
  const stateOf = (c, now = Date.now()) =>
    commentState(c, {
      working: state.turn?.ref === c.ref ? state.turn : null,
      now,
      noChange: state.results.get(c.ref)?.noChange === true,
      elsewhere: elsewhere(c.ref),
    });
  // Open or close the card of a row; one card is open at a time. A card opened
  // by the author is scrolled to the top of the list area, under the bar that
  // stays at the top of the tab.
  function paintOpen(ref, els) {
    const open = state.openRef === ref;
    els.body.hidden = !open;
    els.head.setAttribute("aria-expanded", open ? "true" : "false");
    els.card.classList.toggle("open", open);
  }
  function openCard(ref, { toggle = true, scroll = true } = {}) {
    state.openRef = toggle && state.openRef === ref ? null : ref;
    for (const [r, els] of state.cards) paintOpen(r, els);
    const els = state.cards.get(state.openRef);
    if (scroll && els) {
      const top = $("review-top");
      els.card.style?.setProperty?.("scroll-margin-top", `${(top?.offsetHeight ?? 0) + 4}px`);
      els.card.scrollIntoView?.({ block: "start" });
    }
  }
  // The words of a row, and the work indicator inside its card.
  function paintCard(c, els, now = Date.now()) {
    const st = stateOf(c, now);
    els.stateEl.textContent = st.label;
    els.stateEl.setAttribute("class", `review-comment-state review-state-${st.kind}`);
    const working = state.turn?.ref === c.ref ? state.turn : null;
    els.progressEl.hidden = !working;
    if (working) {
      els.clockEl.textContent = `Working on this comment: ${clockText(working.startedAt, now)}`;
    }
    return st;
  }
  function paintCards() {
    const now = Date.now();
    for (const c of state.revision?.comments ?? []) {
      const els = state.cards.get(c.ref);
      if (els) paintCard(c, els, now);
    }
  }
  function renderSummary() {
    const comments = state.revision?.comments ?? [];
    const box = $("review-summary");
    box.textContent = summaryText(
      comments,
      comments.map((c) => stateOf(c)),
    );
    box.hidden = comments.length === 0;
  }

  function renderComments() {
    const box = $("review-comments");
    $("review-preview-holder").append($("review-preview")); // out of the cards that go
    box.replaceChildren();
    proposeButtons = [];
    state.cards.clear();
    for (const c of state.revision?.comments ?? []) {
      const status = el("select", { class: "setting-select" });
      for (const [v, label] of Object.entries(STATUS_LABELS)) {
        const o = el("option", { value: v, text: label });
        if (v === c.status) o.selected = true;
        status.append(o);
      }
      status.addEventListener(
        "change",
        guard(async () => {
          // L5: "Done" needs a second try while something is left open:
          // a requirement no change answers, a change not decided in Word,
          // or a ledger rule in violation.
          const left = doneBlockers(
            c,
            state.revision?.ledger,
            elsewhere(c.ref),
            state.revision?.consistency,
          );
          if (status.value === "done" && left.length && state.confirmDone !== c.ref) {
            state.confirmDone = c.ref;
            throw new Error(
              `${c.ref}: ${left.join("; ")}. Choose Done again to mark it done anyway.`,
            );
          }
          state.confirmDone = null;
          await ask("review_set_status", { ref: c.ref, status: status.value });
        }),
      );
      const propose = el("button", {
        type: "button",
        class: "btn-primary btn-small",
        text: "Propose changes",
      });
      // In a supplementary file the requests are shown and not edited; the
      // turn is proposed here as in a manuscript, for this file.
      const main = partOf();
      proposeButtons.push(propose);
      propose.disabled = isTurnInFlight() || queue.running;
      // The turn is taken on the click, before the selection and refresh()
      // waits: a second click meanwhile is ignored (Codex review 8) and the
      // daemon, which queues the agent message before it answers, may send
      // turn_complete and turn_done before the answer (Codex review 7).
      // onTurnStarted returns a release that ends only this turn. The check
      // runs before guard() so an ignored click keeps the message line.
      const run = guard((release) => proposeRef(c.ref, release, true));
      propose.addEventListener("click", () => {
        if (isTurnInFlight() || queue.running) return;
        return run(
          onTurnStarted(`Review ${c.ref}: proposing changes (tracked, one comment per change)...`),
        );
      });
      const preview = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Preview prompt",
      });
      preview.addEventListener(
        "click",
        guard(async () => {
          // The same paragraphs as Propose changes would send now.
          const r = await ask("review_preview", {
            ref: c.ref,
            paragraphs: await promptParagraphs(),
          });
          $("review-preview").textContent = r.prompt;
          $("review-preview").hidden = false;
          state.previewRef = c.ref;
        }),
      );
      const reqs = el("textarea", {
        rows: "2",
        class: "review-reqs",
        placeholder: "What it asks for, one per line",
      });
      // A render rebuilds this textarea: unsaved text lives in state and is
      // dropped only when the save succeeds.
      reqs.value = state.reqDrafts.has(c.ref)
        ? state.reqDrafts.get(c.ref)
        : (c.requirements ?? []).map((r) => r.text).join("\n");
      reqs.addEventListener("input", () => state.reqDrafts.set(c.ref, reqs.value));
      const saveReqs = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Save requirements",
      });
      saveReqs.addEventListener(
        "click",
        guard(async () => {
          await ask("review_set_requirements", {
            ref: c.ref,
            requirements: requirementsFromText(reqs.value),
          });
          state.reqDrafts.delete(c.ref);
          state.reqEdit.delete(c.ref);
        }),
      );
      // What the reviewer asks: one line per request with where it stands. The
      // box to change the list opens from Edit.
      const rows = requirementStates(c, elsewhere(c.ref));
      const editing = state.reqEdit.has(c.ref);
      const edit = el("button", {
        type: "button",
        class: "btn-link",
        text: editing ? "Close the editor" : "Edit",
      });
      edit.addEventListener("click", () => {
        if (state.reqEdit.has(c.ref)) state.reqEdit.delete(c.ref);
        else state.reqEdit.add(c.ref);
        render();
      });
      reqs.hidden = !editing;
      saveReqs.hidden = !editing;
      const requestList = el(
        "div",
        { class: "review-req-list" },
        el(
          "div",
          { class: "review-questions-head" },
          "What the reviewer asks ",
          main ? el("span", { class: "review-part-note", text: `(${editedInText(main)})` }) : edit,
        ),
        ...(rows.length
          ? rows.map((r) =>
              el(
                "div",
                {
                  class: "review-req",
                  "data-req-id": r.id,
                  "data-req-text": r.text,
                  "data-state": r.state,
                  ...(r.part ? { "data-req-part": r.part } : {}),
                },
                el("span", { class: "review-req-text", text: r.text }),
                el("span", {
                  class: `review-req-state review-req-${r.state}`,
                  text: requestStateText(r),
                }),
              ),
            )
          : [
              el("div", {
                class: "review-req-none",
                text: "No requests listed yet. Use Edit to add them, one per line; until then a turn lists them itself.",
              }),
            ]),
      );
      const stateEl = el("span", { class: "review-comment-state" });
      const head = el(
        "button",
        { type: "button", class: "review-comment-head", "aria-expanded": "false" },
        el("strong", { text: c.ref }),
        el("span", { class: "review-comment-snippet", text: firstLine(c.text) }),
        stateEl,
      );
      head.addEventListener("click", () => openCard(c.ref));
      // The work of this comment, inside its card: the clock and a way to stop.
      const clockEl = el("span", { class: "review-card-clock" });
      const stop = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Stop",
      });
      stop.addEventListener("click", () => stopTurn());
      const progressEl = el("div", { class: "review-card-progress", role: "timer" }, clockEl, stop);
      progressEl.hidden = true;
      // What the last turn of this comment came to.
      const res = state.results.get(c.ref);
      const result = res
        ? el(
            "div",
            { class: `review-card-result${res.bad ? " form-error" : ""}` },
            el("div", { text: res.text }),
            res.reason
              ? el(
                  "details",
                  { class: "review-card-reason" },
                  el("summary", { text: "Why no change" }),
                  el("div", { text: res.reason }),
                )
              : null,
          )
        : null;
      // What the author has to look at by hand (an undo that took more
      // than the edit). It stays until the next turn of this comment.
      const notice = state.notices.has(c.ref)
        ? el("div", {
            class: "review-card-result review-card-notice form-error",
            text: state.notices.get(c.ref),
          })
        : null;
      const body = el(
        "div",
        { class: "review-comment-body" },
        progressEl,
        result,
        notice,
        el("div", { class: "review-card-actions" }, status, propose),
        c.section ? el("div", { class: "review-comment-section", text: c.section }) : null,
        el("div", { class: "review-comment-text", text: c.text }),
        requestList,
        // A request that is answered in another file, or a figure change:
        // the comparison and the decision are on the figures page.
        !main && hasFigureWork(c) ? figuresLink(c.ref) : null,
        main ? null : reqs,
        main ? null : saveReqs,
      );
      const card = el("div", { class: "review-comment", "data-ref": c.ref }, head, body);
      const advanced = el(
        "details",
        { class: "review-card-advanced" },
        el("summary", { text: "Advanced" }),
        preview,
      );
      advanced.open = state.previewRef === c.ref;
      state.cards.set(c.ref, { card, head, body, stateEl, progressEl, clockEl, advanced });
      // What the model asked instead of guessing (office_ask_author). An
      // answer goes into the prompt of the next turn.
      if ((c.questions ?? []).length) {
        body.append(
          el("div", { class: "review-questions-head", text: "Questions for the author" }),
        );
        for (const q of c.questions) {
          const draftKey = `${c.ref}/${q.id}`;
          const answer = el("textarea", {
            rows: "2",
            class: "review-answer",
            placeholder: "Your answer (it goes into the next turn)",
          });
          answer.value = state.answerDrafts.has(draftKey)
            ? state.answerDrafts.get(draftKey)
            : (q.answer ?? "");
          answer.addEventListener("input", () => state.answerDrafts.set(draftKey, answer.value));
          const saveAnswer = el("button", {
            type: "button",
            class: "btn-secondary btn-small",
            text: "Save answer",
          });
          saveAnswer.addEventListener(
            "click",
            guard(async () => {
              await ask("review_question_answer", { ref: c.ref, id: q.id, answer: answer.value });
              state.answerDrafts.delete(draftKey);
            }),
          );
          body.append(
            el("div", {
              class: "review-question",
              text: `${q.id}${q.requirement_id ? ` (${q.requirement_id})` : ""}: ${q.question}`,
            }),
            el("div", { class: "review-reqs-line", text: `Why it is asked: ${q.why}` }),
            answer,
            saveAnswer,
          );
        }
      }
      for (const ch of c.changes) body.append(changeRow(ch));
      // Terms the changes took out that other paragraphs still hold (residue.js).
      for (const it of stillPresentItems(c, state.revision?.snapshot)) {
        const line = el("div", { class: "review-residue", text: it.text });
        for (const i of it.indexes) line.append(" ", paragraphLink(i));
        if (it.more) line.append(el("span", { text: ` and ${it.more} more` }));
        body.append(line);
      }
      // A paragraph check finding or a stopped rewrite in a paragraph this
      // comment changed (commentNotices).
      for (const n of commentNotices(c, state.revision?.consistency)) {
        const line = el("div", { class: "review-residue review-notice", text: n.text });
        if (n.paragraph_index != null) line.append(" ", paragraphLink(n.paragraph_index));
        body.append(line);
      }
      for (const b of c.blocked) if (!inFlightEdit(b)) body.append(blockedRow(c.ref, b));
      body.append(advanced);
      box.append(card);
    }
    // The prompt shown by Preview prompt sits under Advanced in its card.
    const pre = $("review-preview");
    const shown = state.cards.get(state.previewRef);
    if (shown) shown.advanced.append(pre);
    else {
      state.previewRef = null;
      $("review-preview-holder").append(pre);
    }
    if (state.openRef && !state.cards.has(state.openRef)) state.openRef = null;
    for (const [ref, els] of state.cards) paintOpen(ref, els);
    paintCards();
  }

  function renderLedger() {
    const box = $("review-ledger");
    box.replaceChildren();
    for (const item of state.revision?.ledger ?? []) {
      const check = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Check",
      });
      check.addEventListener(
        "click",
        guard(() => runLedgerCheck([item.id])),
      );
      const remove = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Remove",
      });
      remove.addEventListener(
        "click",
        guard(() => ask("ledger_remove", { id: item.id })),
      );
      // L6: load the item into the form; Add becomes "Save <id>" (ledger_update).
      const edit = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Edit",
      });
      edit.addEventListener(
        "click",
        guard(() => {
          state.editingLedger = item.id;
          fillLedgerForm(ledgerFormFromItem(item));
        }),
      );
      const statusText = ledgerStatusText(item);
      const how = ledgerHowText(item);
      box.append(
        el(
          "div",
          { class: "review-ledger-item", "data-id": item.id, "data-kind": item.kind },
          el("strong", { text: item.kind === "forbid" ? "Forbidden wording" : "Instruction" }),
          el("span", { text: ` ${item.text}${how ? ` (${how})` : ""}. ${statusText}. ` }),
          check,
          partOf() ? null : edit,
          partOf() ? null : remove,
        ),
      );
    }
    // In a supplementary file the rules are those of the manuscript: they
    // are shown and checked here, and added or changed there.
    const main = partOf();
    $("review-ledger-note").textContent = main ? `These rules are ${editedInText(main)}.` : "";
    $("review-ledger-note").hidden = !main;
    for (const id of [
      "review-ledger-form",
      "review-ledger-advanced",
      "ledger-import",
      "ledger-suggest-review",
      "ledger-candidates",
    ]) {
      $(id).hidden = !!main;
    }
    const rules = state.revision?.ledger ?? [];
    $("review-ledger-count").textContent = rules.length ? `(${rules.length})` : "";
    $("review-ledger-add").textContent = state.editingLedger
      ? `Save ${state.editingLedger}`
      : "Add";
  }

  // The link between a manuscript and its supplementary files. A manuscript:
  // one line that names them. A supplementary file: the line that says where
  // its list comes from, with the way out. A file with no review of its own:
  // the box to join the revision of a manuscript (nothing joins by itself).
  async function joinRevision(main) {
    // The name the response letter gives this file; none typed: its file name.
    const label = $("review-link-label")?.value.trim() ?? "";
    const r = await ask("review_link_join", { main, ...(label ? { label } : {}) });
    state.linked = r.linked ?? null;
    state.linkCandidates = [];
    $("review-link-path").value = "";
    if ($("review-link-label")) $("review-link-label").value = "";
    say(`Joined the revision of ${partOf()?.name ?? "the manuscript"}.`);
  }
  function renderLink() {
    const main = partOf();
    const parts = state.linked?.parts ?? [];
    const line = $("review-linked");
    line.replaceChildren();
    if (main) {
      const leave = el("button", {
        type: "button",
        class: "btn-link review-link-leave",
        text: "Leave this revision",
      });
      leave.addEventListener(
        "click",
        guard(async () => {
          await ask("review_link_leave");
          await load();
          say(`This file left the revision of ${main.name}.`);
        }),
      );
      line.append(el("div", { class: "review-linked-text", text: partOfText(main) }));
      if (state.linked.note)
        line.append(el("div", { class: "form-error", text: state.linked.note }));
      // What the response letter calls this file, set here at any time.
      const name = el("input", {
        type: "text",
        id: "review-part-label",
        maxlength: "80",
        "aria-label": "Name of this file in the response letter",
      });
      name.value = main.label ?? "";
      const saveName = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Save name",
      });
      saveName.addEventListener(
        "click",
        guard(async () => {
          const r = await ask("review_link_label", { label: name.value });
          state.linked = r.linked ?? state.linked;
          say(`The response letter calls this file ${partOf()?.label}.`);
        }),
      );
      line.append(
        el(
          "div",
          { class: "review-part-label" },
          el("label", { for: "review-part-label", text: "Name in the response letter" }),
          el("div", { class: "review-link-typed" }, name, saveName),
        ),
      );
      line.append(leave);
    } else if (parts.length) {
      line.append(el("div", { class: "review-linked-text", text: linkedFilesText(parts) }));
    }
    line.hidden = !main && parts.length === 0;
    // Offered to a file with no review of its own, and to one whose
    // comments are all in the list of a manuscript of its folder (a
    // supplementary file that left its revision and kept its changes).
    const own = (state.revision?.comments ?? []).length;
    const offer = own === 0 ? state.draft.length === 0 : state.linkCandidates.length > 0;
    $("review-link-box").hidden = !state.linked || !!main || parts.length > 0 || !offer;
    // With no manuscript to offer, the box is the last thing on the tab.
    $("review-link-box").setAttribute(
      "class",
      `review-fold review-link${state.linkCandidates.length ? "" : " review-link-last"}`,
    );
    $("review-link-candidates").replaceChildren(
      ...state.linkCandidates.map((c) => {
        const b = el("button", {
          type: "button",
          class: "btn-primary btn-small review-link-candidate",
          text: joinLabel(c),
        });
        b.addEventListener(
          "click",
          guard(() => joinRevision(c.doc)),
        );
        return b;
      }),
    );
  }

  // P15: candidate rules from the review text, one checkbox each.
  function renderCandidates() {
    const box = $("ledger-candidates");
    box.replaceChildren();
    for (const c of state.candidates) {
      const tick = el("input", { type: "checkbox" });
      tick.checked = c.picked;
      tick.addEventListener("change", () => (c.picked = tick.checked));
      box.append(
        el(
          "label",
          { class: "review-ledger-item" },
          tick,
          el("span", { text: ` ${candidateLabel(c)}` }),
        ),
      );
    }
    $("ledger-candidates-actions").hidden = state.candidates.length === 0;
  }

  // L7: everything that waits for the author in one box.
  function renderDecisions() {
    const d = pendingDecisions(state.revision, state.linked);
    const box = $("review-decisions");
    // A request waits for the author only once a turn has run for its comment;
    // before that it waits for the tool. Those items stay in the list (hidden)
    // and are not counted.
    const turnRan = (ref) => {
      const c = (state.revision?.comments ?? []).find((x) => x.ref === ref);
      return (
        !!c &&
        (c.changes.length > 0 ||
          (c.questions ?? []).length > 0 ||
          (c.blocked ?? []).length > 0 ||
          (c.requirements_outside ?? []).length > 0 ||
          state.results.has(ref))
      );
    };
    const items = [
      ...d.blocked.map((x) => ({ x, shown: true })),
      ...d.ledger.map((x) => ({ x, shown: true })),
      ...d.requirements.map((x) => ({ x, shown: turnRan(x.ref) })),
    ];
    const total = items.filter((i) => i.shown).length;
    const count = el("div", { text: `Waiting for you: ${total}` });
    count.hidden = true; // the summary of the box says it
    box.replaceChildren(count);
    for (const { x, shown } of items) {
      const row = el("div", {
        class: "review-decision",
        text: x.text,
        role: "button",
        tabindex: "0",
        "data-kind": x.kind,
        "data-ref": x.ref ?? "",
        "data-req": x.kind.startsWith("request") ? x.id : "",
      });
      row.hidden = !shown;
      // Opens what it is about: the comment, or the rules.
      const go = () => {
        if (x.ref) openCard(x.ref, { toggle: false });
        else {
          $("review-ledger-box").open = true;
          $("review-ledger-box").scrollIntoView?.({ block: "start" });
        }
      };
      row.addEventListener("click", go);
      row.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault?.();
          go();
        }
      });
      box.append(row);
    }
    $("review-waiting-count").textContent = total ? `(${total})` : "";
    $("review-waiting-box").hidden = total === 0;
  }

  // L4: changes made outside the tool since the last refresh.
  function renderExternal() {
    const box = $("review-external");
    box.replaceChildren();
    const outside = state.revision?.external ?? [];
    $("review-external-count").textContent = outside.length ? `(${outside.length})` : "";
    $("review-external-box").hidden = outside.length === 0;
    for (const x of outside) {
      // A deleted paragraph is not in the snapshot: its old first words name it.
      const what =
        x.paragraph_index == null
          ? `\u201c${firstWords(x.before)}\u201d`
          : paragraphLabel(state.revision?.snapshot, x.paragraph_index).label;
      const verb = { inserted: "added", deleted: "deleted" }[x.kind] ?? "changed";
      const label =
        `A paragraph was ${verb} outside the tool: ${what}.` +
        (x.lost_facts.length
          ? ` The old text had ${x.lost_facts.join("; ")}, and the new text does not. Check that this is meant.`
          : "");
      const go = el("button", { type: "button", class: "review-violation", text: label });
      if (x.paragraph_index != null)
        go.addEventListener(
          "click",
          guard(() => word.selectParagraphByIndex(x.paragraph_index)),
        );
      go.title = `Before: ${x.before}\nAfter: ${x.after}`;
      const dismiss = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Dismiss",
      });
      dismiss.addEventListener(
        "click",
        guard(() => ask("review_external_dismiss", { id: x.id })),
      );
      box.append(
        el(
          "div",
          { class: `review-external-item${x.lost_facts.length ? " review-flag-lost" : ""}` },
          go,
          dismiss,
        ),
      );
    }
  }

  // L9: elapsed time of the running turn and comments not done. P15: the
  // place in a "Propose all" run goes in front ("Propose all: 2/5 R1.2.").
  function showProgress() {
    const bar = $("review-progress");
    paintCards();
    bar.hidden = !state.turn && !state.queueNote;
    if (bar.hidden) return;
    bar.textContent = [
      state.queueNote,
      state.turn && progressText({ ...state.turn, comments: state.revision?.comments ?? [] }),
    ]
      .filter(Boolean)
      .join(" ");
  }
  function startProgress(ref, startedAt) {
    state.turn = { ref, startedAt: startedAt ?? Date.now() };
    clearInterval(state.timer);
    state.timer = setInterval(showProgress, 1000);
    showProgress();
  }
  function stopProgress() {
    clearInterval(state.timer);
    state.timer = null;
    state.turn = null;
    showProgress();
  }

  // The findings of the last "Check revised paragraphs"; a click selects
  // the paragraph in Word.
  function renderConsistency() {
    const box = $("review-consistency");
    box.replaceChildren();
    for (const line of consistencyLines(state.revision?.consistency)) {
      const b = el("button", {
        type: "button",
        class: "review-violation",
        text: `${paragraphLabel(state.revision?.snapshot, line.paragraph_index).label}: ${line.text}`,
      });
      b.addEventListener(
        "click",
        guard(() => word.selectParagraphByIndex(line.paragraph_index)),
      );
      box.append(b);
      // Three or more changes in the paragraph: a turn of the comment with
      // the most of them writes it anew (office_tracked_rewrite).
      if (line.rewrite_ref) {
        const rewrite = el("button", {
          type: "button",
          class: "btn-secondary btn-small",
          text: "Rewrite this paragraph",
        });
        proposeButtons.push(rewrite);
        rewrite.disabled = isTurnInFlight() || queue.running;
        const run = guard((release) =>
          proposeRef(line.rewrite_ref, release, false, { rewrite_paragraph: line.paragraph_index }),
        );
        rewrite.addEventListener("click", () => {
          if (isTurnInFlight() || queue.running) return;
          return run(
            onTurnStarted(
              `Review ${line.rewrite_ref}: rewriting ${paragraphLabel(state.revision?.snapshot, line.paragraph_index).label} (tracked)...`,
            ),
          );
        });
        box.append(rewrite);
      }
    }
  }

  // What has nothing to show is hidden; the paste box shrinks to one line once
  // the review is split.
  function renderSections() {
    const comments = state.revision?.comments ?? [];
    const hasReview = comments.length > 0 || state.draft.length > 0;
    // In a supplementary file the review is pasted, split and changed in
    // the manuscript.
    const part = !!partOf();
    $("review-start").hidden = hasReview || part;
    $("review-paste-section").hidden = part;
    const collapsed = hasReview && !state.pasteOpen;
    $("review-paste-box").hidden = collapsed;
    $("review-paste-summary").hidden = !hasReview;
    $("review-paste-summary-text").textContent = comments.length
      ? reviewLine(comments)
      : `Review split into ${state.draft.length} comment${state.draft.length === 1 ? "" : "s"}, not saved yet`;
    $("review-paste-toggle").textContent = state.pasteOpen
      ? "Hide review text"
      : "Show review text";
    $("review-paste-toggle").setAttribute("aria-expanded", state.pasteOpen ? "true" : "false");
    // Refresh status stays reachable while nothing is loaded (no daemon), so a
    // press says why; once a review is loaded and empty there is nothing to list.
    $("review-comments-section").hidden = comments.length === 0 && !!state.revision;
    $("review-propose-all").hidden = comments.length === 0;
    $("review-more").hidden = comments.length === 0;
    $("review-submit-box").hidden = comments.length === 0;
    // The submission check counts the answers of the letter: it runs in the
    // manuscript. The comment cleanup under it is each file's own.
    $("review-submit-check").hidden = part;
    $("review-submit-rows").hidden = part;
    $("review-submit-help").textContent = part
      ? `The submission check runs in ${partOf().name}.`
      : "Checks this file and the files made from it. Nothing is changed.";
    const notDone = comments.filter((c) => c.status !== "done").length;
    // Once a check has run the header counts its rows, not the comments.
    $("review-submit-count").textContent = part
      ? ""
      : state.submission
        ? submissionCount(state.submission.rows)
        : notDone
          ? `(${notDone} not done)`
          : "";
    renderSubmission();
  }

  // The rows of the last "Check now": the state, one sentence, and the way
  // to where the row is fixed. They are as of that press: the line under
  // them says when, and a new connection drops them.
  function renderSubmission() {
    const box = $("review-submit-rows");
    const check = state.submission;
    if (!check) return box.replaceChildren();
    const rows = check.rows.map((row) => {
      const label = submissionGoLabel(row.go);
      let go = null;
      // The figures page is outside the pane: the link the comment cards use.
      if (row.go?.figures) go = figuresLink(row.go.ref ?? null);
      else if (label) {
        go = el("button", { type: "button", class: "btn-link review-submit-go", text: label });
        go.addEventListener(
          "click",
          guard(() =>
            Number.isInteger(row.go.paragraph)
              ? word.selectParagraphByIndex(row.go.paragraph)
              : goTo(row.go.tab, row.go.target),
          ),
        );
      }
      return el(
        "div",
        { class: `review-submit-row review-submit-${row.state}`, "data-row": row.id },
        el("span", {
          class: "review-submit-state",
          text: SUBMISSION_STATE_LABELS[row.state] ?? row.state,
        }),
        el("span", { class: "review-submit-text", text: row.text }),
        go,
      );
    });
    box.replaceChildren(
      ...rows,
      el("div", {
        class: "review-submit-asof",
        text: `${check.doc ? `${check.doc}, c` : "C"}hecked at ${check.at}. ${
          check.unsaved
            ? "This document has changes that are not saved: save it, then press Check now again."
            : "Press Check now again after a change."
        }`,
      }),
    );
  }

  function render() {
    // The cards are built anew: the author stays where they were.
    const keep = root.scrollTop;
    renderSections();
    renderLink();
    renderDraft();
    renderComments();
    renderSummary();
    renderConsistency();
    renderLedger();
    renderCandidates();
    renderDecisions();
    renderExternal();
    // The figures page records under the manuscript, not under one of its
    // supplementary files (daemon/figure-http.mjs).
    $("review-figures-all").replaceChildren(...(partOf() ? [] : [figuresLink()]));
    showProgress();
    $("review-file").textContent = state.path ? `Review file: ${state.path}` : "";
    if (root.scrollTop !== keep) root.scrollTop = keep;
  }

  // ---- actions --------------------------------------------------------------
  // Start the review turn of one comment. `release` (from onTurnStarted) ends
  // the turn taken for it when the request fails.
  // useSelection: send the paragraphs Preview prompt would show
  // (promptParagraphs); Propose all sends none, a selection belongs to one
  // comment.
  // What is on record for a comment: its changes and its blocked edits.
  function onRecord(ref) {
    const c = (state.revision?.comments ?? []).find((x) => x.ref === ref);
    return { changes: (c?.changes ?? []).length, blocked: (c?.blocked ?? []).length };
  }
  // What came of the turn of `ref`, for its card: how many changes it put on
  // record and what it ran on. null for a turn this pane did not start
  // (another connection).
  function proposedLine(ref) {
    const before = state.turnStart.get(ref);
    state.turnStart.delete(ref);
    if (!before) return null;
    const now = onRecord(ref);
    const n = now.changes - before.changes;
    const blocked = now.blocked - before.blocked;
    const ran = usedLabel("turn");
    const withModel = ran ? ` with ${ran}` : "";
    const blockedText = `${blocked} edit${blocked === 1 ? " was" : "s were"} blocked (see the blocked list of the comment)`;
    // Requests the turn left without a change, a question or a note.
    const open = openRequestsText(
      (state.revision?.comments ?? []).find((x) => x.ref === ref),
      elsewhere(ref),
    );
    const line =
      n > 0
        ? `Proposed ${n} change${n === 1 ? "" : "s"} for ${ref}${withModel}.${blocked > 0 ? ` ${blockedText}.` : ""}`
        : `No change was proposed for ${ref}${withModel}: ${blocked > 0 ? blockedText : "the reason is below"}.`;
    return { text: open ? `${line} ${open}` : line, noChange: n <= 0 };
  }
  // extra: more fields of the request (rewrite_paragraph for a rewrite turn).
  async function proposeRef(ref, release, useSelection, extra = {}) {
    state.reviewTurnOpen = true;
    state.hiddenInTurn = false;
    state.results.delete(ref);
    state.notices.delete(ref);
    // The card shows the work from the click on, not from the answer of the daemon.
    startProgress(ref);
    try {
      const paragraphs = useSelection ? await promptParagraphs() : [];
      // A fresh snapshot first: outside changes are caught before the turn
      // (L4) and the new words of the turn are checked against it (L2).
      await refresh();
      state.turnStart.set(ref, onRecord(ref));
      // Word minimized: its edits went in without their comments (stage 3
      // of the acceptance run of 2026-10-05). Said after the refresh, which
      // writes the message line; the turn is not held back.
      if (document.hidden === true) say(KEEP_WORD_OPEN, true);
      const done = state.turnsDone;
      // wait_ms: how long this pane waits for the answer. The daemon starts
      // no turn for a request that waited longer at its gate.
      const r = await ask(
        "review_propose",
        { ref, paragraphs, wait_ms: PROPOSE_WAIT_MS, ...extra },
        PROPOSE_WAIT_MS,
      );
      if (state.turnsDone === done && isTurnInFlight()) {
        if (r.started_at && state.turn?.ref === ref) state.turn.startedAt = r.started_at;
      } else if (state.turnsDone !== done)
        state.reviewTurnOpen = false; // its turn_done came first
      else if (state.turn?.ref === ref) stopProgress(); // no turn came of it
    } catch (e) {
      state.reviewTurnOpen = false;
      if (state.turn?.ref === ref) stopProgress();
      state.results.set(ref, { text: e.message, bad: true });
      release?.();
      throw e;
    }
  }

  // ---- P15: Propose all -------------------------------------------------------
  // A turn of the queue ends at its turn_done (after the reload and refresh
  // that event starts), at a new connection (reload), or after TURN_WAIT_MS
  // without either. turn_done says how the turn ended ({ ok, error }): a
  // usage limit, a refusal, a stop or a crash ends the run, and so does a
  // reload or refresh that fails after a turn that ran. The error event
  // of the chat is not used: the daemon may send it after turn_done.
  function waitTurn(ref) {
    return new Promise((resolve) => {
      const timer = setTimeout(
        () => endWait(ref, { ok: false, error: "No turn end after 30 minutes." }),
        TURN_WAIT_MS,
      );
      timer.unref?.();
      state.turnWaiters.set(ref, { resolve, timer });
    });
  }
  function endWait(ref, result) {
    const w = state.turnWaiters.get(ref);
    if (!w) return;
    state.turnWaiters.delete(ref);
    clearTimeout(w.timer);
    w.resolve(result);
  }
  function endAllWaits(error) {
    for (const ref of [...state.turnWaiters.keys()]) endWait(ref, { ok: false, error });
  }
  const queue = createProposeQueue({
    // The selection is not sent: it belongs to one comment, not to all.
    async proposeOne(ref) {
      if (isTurnInFlight()) return { ok: false, error: "Another turn is running." };
      const release = onTurnStarted(`Review ${ref}: proposing changes (Propose all)...`);
      const ended = waitTurn(ref); // before the request: turn_done may beat its answer
      try {
        await proposeRef(ref, release, false);
      } catch (e) {
        endWait(ref, { ok: false, error: e.message });
      }
      return ended;
    },
    onProgress({ index, total, ref }) {
      state.queueNote = `Propose all: ${index + 1}/${total} ${ref}.`;
      showProgress();
    },
  });
  function showQueueButtons() {
    $("review-propose-all").disabled = queue.running;
    $("review-propose-stop").hidden = !queue.running;
    for (const b of proposeButtons) b.disabled = isTurnInFlight() || queue.running;
  }
  // Not under guard(): the run outlasts many renders and keeps the message
  // line of each refresh until its own summary.
  $("review-propose-all").addEventListener("click", async () => {
    if (queue.running || isTurnInFlight()) return;
    // As guard() does for every other action: nothing is sent before the
    // daemon has confirmed the stored provider.
    const confirming = providersReady();
    if (confirming) {
      say(PROVIDER_WAIT);
      await confirming;
      if (queue.running || isTurnInFlight()) return; // a second press, meanwhile
      say("");
    }
    const refs = queueRefs(state.revision);
    if (!refs.length) return say("No open comment to propose changes for.", true);
    const running = queue.run(refs);
    showQueueButtons();
    const r = await running;
    state.queueNote = null;
    const ok = r.done.filter((d) => d.ok).length;
    const tail = `${ok} of ${refs.length} proposed.`;
    const notCounted = r.done.at(-1)?.unchecked ? `${r.failed} is not counted: ` : "";
    if (r.failed) say(`Propose all stopped at ${r.failed}: ${r.error} ${notCounted}${tail}`, true);
    else if (r.stopped) say(`Propose all stopped before ${r.next}. ${tail}`);
    else {
      // What the turns ran on, when the pane knows it.
      const ran = usedLabel("turn");
      say(`Propose all${ran ? ` with ${ran}` : ""}: ${tail}`);
    }
    render();
    showQueueButtons();
  });
  $("review-propose-stop").addEventListener("click", () => {
    queue.stop();
    say("Propose all stops after the comment that is running.");
  });

  async function load() {
    const r = await ask("review_load");
    state.path = r.path;
    state.linked = r.linked ?? null;
    // A file with no link: the manuscripts in its folder it could join.
    // One that holds comments is offered only a manuscript whose list has
    // them all (it was its supplementary file). Asked only of a daemon that
    // knows links.
    state.linkCandidates = [];
    const free = r.linked && !r.linked.part_of && r.linked.parts.length === 0;
    if (free) {
      try {
        const found = (await ask("review_link_candidates")).candidates ?? [];
        const own = (state.revision?.comments ?? []).length;
        state.linkCandidates = own === 0 ? found : found.filter((c) => c.covers === true);
      } catch {
        /* the folder could not be read: the path field is still there */
      }
    }
    $("review-link-box").open = state.linkCandidates.length > 0;
  }

  // Word.Document.saved as the pane sends it: true, false, or null when
  // Word did not say (tools-word.js documentSaved).
  const documentSaved = async () => (word.documentSaved ? await word.documentSaved() : null);
  const refresh = () => refreshWith(say);
  // report(text, bad): where the counts go. The message line for a refresh
  // the author or a turn asked for, the small line for an automatic one.
  // Returns the invariant report (review-view.js invariantReport).
  async function refreshWith(report) {
    // Has Word unsaved changes of the document (null: it did not say). The
    // daemon keeps it with the state of the document at this refresh.
    const saved = await documentSaved();
    const snapshot = await word.collectReviewSnapshot();
    const r = await ask("review_refresh", { snapshot, saved });
    // A supplementary file: the citations of its accepted changes (section
    // and quote) go into its own review file after every refresh, from the
    // same snapshot. The response letter of the manuscript reads them from
    // there; it never reads this document itself. A failure is told and
    // does not undo the refresh: the manuscript then says that the record
    // of this file is incomplete.
    let notCited = "";
    if (partOf()) {
      try {
        const headings = (await word.collectHeadings?.()) ?? [];
        await ask("response_state", { snapshot, headings });
        onPartRefreshed();
      } catch (e) {
        notCited = ` What the letter cites from this file was not brought up to date (${e.message}); press Refresh status again.`;
      }
    }
    // Word refuses Comment.resolved in a document in compatibility mode
    // (GeneralException on every comment). The refresh itself has succeeded,
    // so its counts are still shown and the refusal is added to them.
    const unresolved = [];
    for (const id of commentsToResolve(r.rows, state.revision, snapshot.comments)) {
      try {
        await word.resolveComment(id);
      } catch (e) {
        unresolved.push(e?.message ?? String(e));
      }
    }
    const inv = invariantReport(state.revision, snapshot.comments);
    const c = r.counts;
    const outside = r.external?.length
      ? ` Outside changes: ${r.external.length} (list under Changes made outside the tool).`
      : "";
    // One sentence on what to do next; the full counts are under Details.
    report(
      countsSentence(c, inv, { outside: r.external?.length ?? 0, unresolved: unresolved.length }) +
        notCited,
      !inv.ok || !!notCited,
      `Accepted ${c.accepted}, rejected ${c.rejected}, edited ${c.user_edited}, pending ${c.pending}, not found ${c.unlocated}. ` +
        `Changes ${inv.changes}, expected tool comments ${inv.expected_comments}, found ${inv.tool_comments}` +
        invariantText(inv) +
        outside +
        (unresolved.length
          ? ` ${unresolved.length} rationale comment(s) not marked resolved (Word: ${unresolved[0]}; one cause is a document in compatibility mode).`
          : ""),
    );
    return inv; // turn_done: a mismatch ends a Propose all run
  }

  // L10: changed paragraphs only unless full.
  async function runLedgerCheck(ids, full = false) {
    const snapshot = await word.collectReviewSnapshot();
    const r = await askJob("ledger_check", { ids, snapshot, full });
    const box = $("review-violations");
    box.replaceChildren();
    const rules = state.revision?.ledger ?? [];
    for (const res of r.outcomes) {
      const text = rules.find((i) => i.id === res.id)?.text ?? res.id;
      const found = res.violations.length;
      if (found) {
        box.append(
          el("div", {
            class: "review-found",
            text: `${text}: found in ${found} place${found === 1 ? "" : "s"}`,
          }),
        );
      }
      for (const v of res.violations) {
        const where =
          v.paragraph_index == null
            ? "Required text is gone"
            : `${paragraphLabel(state.revision?.snapshot, v.paragraph_index).label}${v.match ? ` has \u201c${v.match}\u201d` : " is waiting to be deleted"}`;
        const b = el("button", {
          type: "button",
          class: "review-violation",
          "data-rule": res.id,
          text: where,
        });
        if (v.paragraph_index != null)
          b.addEventListener(
            "click",
            guard(() => word.selectParagraphByIndex(v.paragraph_index)),
          );
        box.append(b);
      }
      if (res.status === "error")
        box.append(
          el("div", { class: "form-error", text: `The check of ${text} failed: ${res.error}` }),
        );
    }
  }

  // P15: the review text from a .txt or .md file instead of a paste. The
  // file is read in the pane (File.text()); the daemon is not involved.
  $("review-paste-toggle").addEventListener("click", () => {
    state.pasteOpen = !state.pasteOpen;
    render();
  });
  $("review-load-file").addEventListener("click", () => $("review-file-input").click());
  $("review-file-input").addEventListener(
    "change",
    guard(async () => {
      const input = $("review-file-input");
      const f = input.files?.[0];
      if (!f) return;
      const r = reviewTextFromFile({
        name: f.name,
        size: f.size,
        text: f.size > REVIEW_FILE_MAX_BYTES ? "" : await f.text(),
      });
      input.value = ""; // the same file can be picked again
      if (!r.ok) throw new Error(r.error);
      $("review-input").value = r.text;
      say(`Loaded ${f.name}. Split into comments next.`);
    }),
  );
  $("review-split").addEventListener(
    "click",
    guard(async () => {
      const r = await askJob("review_split", { text: $("review-input").value });
      state.draft = r.comments.map((c) => ({
        ...c,
        requirementsText: (c.requirements ?? []).map((q) => q.text).join("\n"),
      }));
      // regex: found by the numbering and layout of the text; llm: read by the model.
      const how = r.method === "llm" ? "read by Claude" : "found by the numbering in the text";
      if (r.warnings?.length) say(`Split into comments (${how}). ${r.warnings.join(" ")}`);
      else {
        const n = r.comments.length;
        say(
          `Split into ${n} comment${n === 1 ? "" : "s"} (${how}). Check the list, then press Save comment list.`,
        );
      }
    }),
  );
  $("review-save").addEventListener(
    "click",
    guard(async () => {
      const comments = state.draft.map((c) => {
        const reviewer = reviewerOfRef(c.ref);
        if (reviewer === null) {
          throw new Error(`Bad ref "${c.ref}": use R<reviewer>.<number> or E.<number>`);
        }
        return {
          ref: c.ref,
          reviewer,
          text: c.text.trim(),
          // Always sent: "" for an empty field clears a label stored by an
          // earlier save (a missing key would keep it).
          section: (c.section ?? "").trim(),
          requirements: requirementsFromText(c.requirementsText),
          // The terms the split's labelling found the comment to name.
          ...(Array.isArray(c.named_terms) ? { named_terms: c.named_terms } : {}),
        };
      });
      const r = await ask("review_save_comments", { comments });
      state.draft = [];
      // Comments a running turn or an open edit still uses are not removed.
      if (r.warnings?.length) say(r.warnings.join(" "));
    }),
  );
  $("review-link-join").addEventListener(
    "click",
    guard(() => joinRevision($("review-link-path").value)),
  );
  $("review-refresh").addEventListener(
    "click",
    once("refresh", guard(refresh), () => $("review-refresh")),
  );
  // The check reads the document as it is: a status refresh first.
  $("review-check-paragraphs").addEventListener(
    "click",
    once(
      "check-paragraphs",
      guard(async () => {
        await refresh();
        say("Checking the revised paragraphs...");
        const r = await askJob("review_check_paragraphs", {});
        say(consistencyMessage(r.consistency, usedLabel("turn")));
      }),
      () => $("review-check-paragraphs"),
    ),
  );
  $("review-check-all").addEventListener(
    "click",
    guard(() => runLedgerCheck([])),
  );
  $("review-check-full").addEventListener(
    "click",
    guard(() => runLedgerCheck([], true)),
  );
  $("ledger-suggest").addEventListener(
    "click",
    guard(() => {
      const d = ledgerCheckFromText($("ledger-text").value);
      $("ledger-check-type").value = d.checkType;
      $("ledger-pattern").value = d.pattern;
      $("ledger-flags").value = d.flags;
    }),
  );
  $("review-ledger-add").addEventListener(
    "click",
    guard(async () => {
      // Only the words typed: the check is drafted from them, as "Draft check
      // from text" does (a phrase in quotation marks is looked for as written,
      // anything else is read by Claude).
      let checkType = $("ledger-check-type").value;
      let pattern = $("ledger-pattern").value;
      let flags = $("ledger-flags").value;
      const drafted = checkType === "regex" && !pattern && !state.editingLedger;
      if (drafted) ({ checkType, pattern, flags } = ledgerCheckFromText($("ledger-text").value));
      const r = ledgerItemFromForm({
        kind: $("ledger-kind").value,
        text: $("ledger-text").value,
        scope: $("ledger-scope").value,
        checkType,
        pattern,
        flags,
        expect: $("ledger-expect").value,
      });
      if (r.error) throw new Error(r.error);
      if (state.editingLedger) {
        await ask("ledger_update", { id: state.editingLedger, patch: r.item });
        state.editingLedger = null;
      } else {
        await ask("ledger_add", { item: r.item });
      }
      $("ledger-text").value = "";
      $("ledger-pattern").value = "";
      if (drafted) {
        say(
          checkType === "regex"
            ? `Rule added. It looks for \u201c${pattern.replace(/\\(.)/g, "$1")}\u201d as written.`
            : "Rule added. Claude reads each paragraph to check it.",
        );
      }
    }),
  );
  // P15: the ledger as a JSON file. Export writes <stem>.ledger.json next to
  // the manuscript (the daemon answers its path); Import adds the rules of a
  // picked file that the ledger does not hold yet.
  async function importLedger(doc) {
    const r = await ask("ledger_import", { doc, mode: "append" });
    say(`Added ${r.added} ledger rule(s), skipped ${r.skipped} already in the ledger.`);
  }
  $("ledger-export").addEventListener(
    "click",
    guard(async () => {
      const r = await ask("ledger_export");
      say(`Exported ${r.count} ledger rule(s) to ${r.path}`);
    }),
  );
  $("ledger-import").addEventListener("click", () => $("ledger-file-input").click());
  $("ledger-file-input").addEventListener(
    "change",
    guard(async () => {
      const input = $("ledger-file-input");
      const f = input.files?.[0];
      if (!f) return;
      const text = f.size > REVIEW_FILE_MAX_BYTES ? "" : await f.text();
      input.value = "";
      if (f.size > REVIEW_FILE_MAX_BYTES) throw new Error(`${f.name} is larger than 2 MB.`);
      const r = ledgerImportFromJson(text);
      if (!r.ok) throw new Error(r.error);
      await importLedger(r.doc);
    }),
  );
  // P15: rule candidates from the review text in the paste box. The daemon
  // stores none of them; "Add selected" sends the ticked ones as an import.
  $("ledger-suggest-review").addEventListener(
    "click",
    guard(async () => {
      const r = await askJob("ledger_suggest", { text: $("review-input").value });
      // A candidate from a cue in the text starts ticked; one only the model
      // labelled waits for the author's tick.
      state.candidates = r.candidates.map((c) => ({ ...c, picked: c.why !== "model label" }));
      say(
        `${r.candidates.length} candidate rule(s) by ${r.method}. Tick the ones to keep, then Add selected.` +
          (r.warnings?.length ? ` ${r.warnings.join(" ")}` : ""),
      );
    }),
  );
  $("ledger-candidates-add").addEventListener(
    "click",
    guard(async () => {
      const items = state.candidates.filter((c) => c.picked).map(candidateRule);
      if (!items.length) throw new Error("Tick at least one candidate rule.");
      await importLedger({ items });
      state.candidates = [];
    }),
  );
  $("ledger-candidates-clear").addEventListener(
    "click",
    guard(() => {
      state.candidates = [];
    }),
  );
  // Tell the daemon which rationale comments the cleanup removed, so their
  // absence is not read as a reject or a missing comment (Codex review 7).
  async function markCleaned(ids) {
    if (ids?.length) await ask("review_comments_cleaned", { ids: ids.map(String) });
  }
  // Check now: the pane reads the document (text, comments, headings,
  // formatting) and the daemon answers the rows. Nothing is changed. The
  // formatting is read last and on its own: when Word cannot give it, the
  // row that needs it says so and the other rows are still computed.
  $("review-submit-check").addEventListener(
    "click",
    once(
      "submit-check",
      guard(async () => {
        if (!toolsJob) throw new Error("The checklist is not available in this pane.");
        say("Reading the document...");
        const snapshot = await word.collectReviewSnapshot();
        const headings = (await word.collectHeadings?.()) ?? [];
        const format = (await word.readFormatSnapshot?.().catch(() => null)) ?? null;
        // Has Word unsaved changes of the document (null: it did not say).
        // The rows about copies made from the saved file depend on it.
        const saved = await documentSaved();
        say("Checking...");
        const r = await toolsJob("submission_check", { snapshot, headings, format, saved });
        if (!r?.ok) throw new Error(r?.error || "The check failed.");
        state.submission = {
          rows: r.rows ?? [],
          doc: r.doc_name ?? null,
          // As the daemon took it (a daemon from before the flag: not said).
          unsaved: r.saved === false,
          at: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
        };
        say("");
      }),
      () => $("review-submit-check"),
    ),
  );
  $("review-clean-tool").addEventListener(
    "click",
    once(
      "clean-tool",
      guard(async () => {
        const r = await word.deleteToolComments();
        await markCleaned(r.deleted_ids);
        const rep = cleanupReport(r.before, r.after);
        say(rep.text, !rep.ok);
      }),
      () => $("review-clean-tool"),
    ),
  );
  // Delete ALL comments: first press shows the counts, second press reads
  // the comments again, backs up that fresh list and deletes the backed-up
  // threads that did not change meanwhile. If the list changed since it was
  // shown, nothing is deleted and the new count is shown for another press. window.confirm is unreliable in Office panes
  // (taskpane.js context-error note), so the confirmation is inline.
  $("review-delete-all").addEventListener(
    "click",
    once(
      "delete-all",
      guard(async () => {
        const btn = $("review-delete-all");
        const offer = (comments) => {
          const n = countComments(comments);
          state.backupShown = { comments, n };
          btn.textContent = `Back up and delete ${n.total} comments (tool ${n.tool}, yours ${n.user})`;
          return n;
        };
        if (!state.backupShown) {
          offer(await word.readAllComments());
          return;
        }
        const current = await word.readAllComments();
        const diff = compareCommentLists(state.backupShown.comments, current);
        if (!diff.same) {
          const n = offer(current);
          throw new Error(
            `Comments changed since the count was shown (${diff.text}). Nothing deleted. Press again to back up and delete ${n.total}.`,
          );
        }
        state.backupShown = null;
        btn.textContent = "Delete ALL comments...";
        const backup = buildCommentsBackup(current, { manuscript: state.path });
        const b = await ask("comments_backup", { backup });
        if (b.count !== current.length || b.reply_count !== backup.reply_count)
          throw new Error(
            `Backup holds ${b.count} comments and ${b.reply_count} replies, expected ${current.length} and ${backup.reply_count}. Nothing deleted.`,
          );
        // Only the backed-up threads, and only those unchanged since the
        // backup: a comment added or edited while the backup was written stays.
        const plan = await word.deleteBackedUpComments(backup.comments);
        const left = new Set((await word.listComments()).comments.map((c) => String(c.id)));
        await markCleaned(plan.delete_ids.filter((id) => !left.has(String(id))));
        const rep = deleteAllReport({
          shown: current.length,
          backedUp: b.count,
          deleted: plan.delete_ids.length,
          changed: plan.changed_ids.length,
          missing: plan.missing_ids.length,
          addedAfter: plan.new_ids.length,
          stillThere: plan.delete_ids.filter((id) => left.has(id)).length,
        });
        say(`${rep.text}. Backup: ${b.path}`, !rep.ok);
      }),
      () => $("review-delete-all"),
    ),
  );

  // ---- P15: automatic refresh ----------------------------------------------------
  // After any turn of the pane and when the pane gets the focus back
  // (auto-refresh.js). A review turn is left to its turn_done, which
  // refreshes already. The result goes to the small line under the file
  // name; the cards are rebuilt only when something changed, so a box the
  // author is typing in keeps the focus.
  const statusKey = () =>
    JSON.stringify([state.revision?.comments, state.revision?.external, state.revision?.ledger]);
  function autoNote(text, bad = false) {
    const n = $("review-auto");
    n.textContent = text;
    n.classList.toggle("form-error", bad);
  }
  const auto = createAutoRefresh({
    isBusy: () =>
      isTurnInFlight() ||
      queue.running ||
      state.reviewTurnOpen ||
      !state.revision?.comments?.length,
    async run() {
      const before = statusKey();
      const at = new Date().toLocaleTimeString();
      await refreshWith((text, bad) => autoNote(`Auto-refreshed ${at}: ${text}`, bad));
      if (statusKey() !== before) render();
    },
    onError: (e) => autoNote(`Auto-refresh failed: ${e?.message ?? e}`, true),
  });
  if (typeof window !== "undefined") {
    window.addEventListener?.("focus", () => auto.trigger("focus"));
  }
  document.addEventListener?.("visibilitychange", () => {
    if (document.hidden) {
      if (state.reviewTurnOpen) state.hiddenInTurn = true;
      return;
    }
    if (state.hiddenInTurn) {
      state.hiddenInTurn = false;
      say(MINIMIZED_IN_TURN, true);
    }
    auto.trigger("focus");
  });

  render(); // loaded on "welcome" (taskpane.js calls reload once connected)

  return {
    onEvent(msg) {
      if (msg.event === "job_done") {
        jobDone(msg);
      } else if (msg.event === "turn_done") {
        state.turnsDone += 1;
        state.reviewTurnOpen = false;
        stopProgress();
        // Propose all: the next comment may start only when this turn ran
        // AND the reload and refresh after it worked AND the refresh found
        // the recorded changes and the tool comments in Word to match. A
        // turn that ran but whose state is not confirmed is not counted
        // (unchecked) and ends the run; what it recorded is saved by the
        // daemon.
        const unchecked = (what) => {
          const error = `The turn for ${msg.ref} ran and its record is saved, but ${what}`;
          say(error, true);
          state.results.set(msg.ref, { text: error, bad: true });
          return { ok: false, unchecked: true, error };
        };
        const failed =
          msg.ok === false
            ? { ok: false, error: msg.error || "The turn ended with an error." }
            : null;
        load()
          .then(refresh)
          .then(
            (inv) => {
              if (failed) {
                state.turnStart.delete(msg.ref);
                state.results.set(msg.ref, { text: failed.error, bad: true });
                return failed;
              }
              // What came of the turn goes into its card.
              const line = proposedLine(msg.ref);
              if (line) {
                state.results.set(msg.ref, {
                  text: line.text,
                  noChange: line.noChange,
                  ...(line.noChange ? { reason: lastReply() } : {}),
                });
              }
              if (inv.ok) return { ok: true };
              // Only changes Word applied without their rationale comment:
              // said as that, not as a mismatch.
              if (
                inv.uncommented.length &&
                !inv.missing_comment.length &&
                !inv.unrecorded_comment_ids.length
              ) {
                return unchecked(
                  `${inv.uncommented.join(", ")} was applied in Word and its rationale comment could not be added; press Add rationale comment on that change, then Refresh status.`,
                );
              }
              return unchecked(
                `the changes on record and the tool comments in Word do not match (missing ${inv.missing_comment.join(",") || "-"}, unrecorded ${inv.unrecorded_comment_ids.join(",") || "-"}); check those in Word, then press Refresh status.`,
              );
            },
            (e) => {
              if (failed) {
                say(e.message, true);
                state.results.set(msg.ref, { text: failed.error, bad: true });
                return failed;
              }
              return unchecked(
                `the status check after it failed (${e.message}); press Refresh status to check again.`,
              );
            },
          )
          .then((result) => {
            render();
            endWait(msg.ref, result);
          });
      } else if (msg.event === "ledger_changed") {
        // The Tools tab added rules to the ledger (journal guideline).
        load()
          .catch((e) => say(e.message, true))
          .finally(render);
      } else if (msg.event === "late_result_recorded") {
        // A tracked edit that timed out was applied after all and is now recorded.
        say(`${msg.ref}: an edit that timed out was applied; recorded as ${msg.change_id}.`);
        load()
          .catch((e) => say(e.message, true))
          .finally(render);
      } else if (msg.event === "record_failed") {
        // The load turns the attempt into an unanswered edit in the list,
        // or, when the daemon could save the change this time, into the
        // change (its attempt is then gone from the list).
        load()
          .then(
            () => true,
            () => false,
          )
          .then((loaded) => {
            render();
            const left = (state.revision?.comments ?? [])
              .find((c) => c.ref === msg.ref)
              ?.blocked?.some((b) => b.id === msg.attempt_id);
            if (loaded && !left) {
              say(`${msg.ref}: the edit was applied in Word; its record is saved now.`);
            } else {
              say(
                `${msg.ref}: the edit was applied in Word but saving its record failed; press Refresh status to match them again.`,
                true,
              );
            }
          });
      } else if (msg.event === "undo_took_more") {
        // An edit of this comment was taken back, and Word took an earlier
        // pending change of the same paragraph with it. The refresh reads
        // that change from the document again (it is no longer pending
        // there); the card of the comment says what to look at.
        state.notices.set(msg.ref, UNDO_TOOK_MORE_TEXT);
        load()
          .then(refresh)
          .catch((e) => say(e.message, true))
          .finally(render);
      } else if (msg.event === "stale_target") {
        // A tracked edit was refused: its paragraph is no longer as the
        // daemon has it. A refresh gives the daemon the document as it is.
        load()
          .then(refresh)
          .catch((e) => say(e.message, true))
          .finally(render);
      } else if (
        msg.event === "change_recorded" ||
        msg.event === "edit_blocked" ||
        msg.event === "comment_noted" ||
        // A figure change was recorded or decided on the figures page.
        msg.event === "figures_changed" ||
        // The link to a manuscript or a supplementary file changed, or the
        // manuscript changed what this supplementary file copies.
        msg.event === "linked_changed"
      ) {
        load()
          .catch((e) => say(e.message, true))
          .finally(render);
      }
    },
    // Any turn of the pane (chat or review) started or ended.
    setTurnBusy(busy) {
      for (const b of proposeButtons) b.disabled = busy || queue.running;
    },
    // Any turn of the pane ended (taskpane.js, TURN_END_EVENTS).
    onAnyTurnEnded: () => auto.trigger("turn"),
    // A new connection (welcome) means no turn of this pane is running any
    // more: a daemon restart never sends turn_done, so clear the progress line.
    // Jobs of the old connection are ended with an error first, and so is a
    // "Propose all" run waiting for its turn.
    reload: () =>
      (dropJobs(),
      endAllWaits(CONNECTION_LOST),
      (state.submission = null),
      (state.reviewTurnOpen = false),
      stopProgress(),
      load())
        .catch((e) => say(e.message, true))
        .finally(render),
  };
}
