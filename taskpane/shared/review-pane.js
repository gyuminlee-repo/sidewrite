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
  unaddressedRequirements,
  ledgerDraftFromChange,
  ledgerFormFromItem,
  ledgerCheckFromText,
  ledgerStatusText,
  pendingDecisions,
  invariantText,
  inFlightEdit,
  reviewerOfRef,
  progressText,
  reviewTextFromFile,
  REVIEW_FILE_MAX_BYTES,
  ledgerImportFromJson,
  candidateRule,
  candidateLabel,
  queueRefs,
} from "./review-view.js";
import { createProposeQueue } from "./propose-queue.js";
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

export function initReviewPane({
  sendRequest,
  word,
  onTurnStarted,
  isTurnInFlight,
  // () => the request that confirms the stored providers while it is under
  // way (providers-pane.js whenRestored), else null.
  providersReady = () => null,
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
    confirmDone: null, // L5: ref whose "Done" was asked once with open requirements
    reqDrafts: new Map(), // ref -> requirement text typed and not saved yet (Codex review 9)
    armedBlocked: new Set(), // ids of blocked edits whose "Allow exception" was pressed once
    armedReapply: new Set(), // ids of outcome_unknown edits whose "Re-apply anyway" was pressed once
    applyingBlocked: new Set(), // ids of blocked edits whose override request is in flight
    turn: null, // L9: { ref, startedAt } of the running review turn
    turnsDone: 0, // turn_done events seen, to spot one that beats its review_propose answer
    timer: null,
    jobs: new Map(), // job_id -> { resolve, reject } of an askJob waiting for job_done
    jobsDone: new Map(), // job_id -> result of a job_done that came before its answer
    candidates: [], // P15: ledger rule candidates of the last suggestion, each with `picked`
    turnWaiters: new Map(), // P15: ref -> { resolve, timer } of a "Propose all" turn
    queueNote: null, // P15: "Propose all: 2/5 R1.2." while the queue runs
    reviewTurnOpen: false, // P15: a review turn was asked for and its turn_done has not come
  };

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
  function say(text, bad = false) {
    const m = $("review-message");
    m.textContent = text;
    m.classList.toggle("form-error", bad);
    m.hidden = !text;
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
          say("Waiting for the daemon to confirm the stored provider choice...");
          await confirming;
        }
        say("");
        await fn(...a);
      } catch (e) {
        say(e.message, true);
      }
      render();
    };

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
    const ids = ch.rationale.requirement_ids ?? [];
    const row = el(
      "button",
      { type: "button", class: `review-change state-${ch.state}` },
      el("span", { class: "review-badge", text: STATE_LABELS[ch.state] ?? ch.state }),
      el("span", {
        text: ` ${ch.kind}: ${ch.rationale.what}${ids.length ? ` [${ids.join(",")}]` : ""}`,
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
      add.addEventListener(
        "click",
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

  function renderComments() {
    const box = $("review-comments");
    box.replaceChildren();
    proposeButtons = [];
    for (const c of state.revision?.comments ?? []) {
      const status = el("select", { class: "setting-select" });
      for (const [v, label] of Object.entries(STATUS_LABELS)) {
        const o = el("option", { value: v, text: label });
        if (v === c.status) o.selected = true;
        status.append(o);
      }
      const open = unaddressedRequirements(c);
      status.addEventListener(
        "change",
        guard(async () => {
          // L5: "Done" with requirements no change answers needs a second try.
          if (status.value === "done" && open.length && state.confirmDone !== c.ref) {
            state.confirmDone = c.ref;
            throw new Error(
              `${c.ref}: ${open.map((r) => `${r.id} "${r.text}"`).join(", ")} has no change. Choose Done again to mark it done anyway.`,
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
        }),
      );
      const reqLine = (c.requirements ?? []).length
        ? `Requirements: ${c.requirements.map((r) => `${r.id} ${r.text}`).join("; ")}.` +
          (open.length ? ` No change yet: ${open.map((r) => r.id).join(", ")}.` : "")
        : "";
      const card = el(
        "div",
        { class: "review-comment" },
        el(
          "div",
          { class: "review-comment-head" },
          el("strong", { text: c.ref }),
          status,
          propose,
          preview,
        ),
        c.section ? el("div", { class: "review-comment-section", text: c.section }) : null,
        el("div", { class: "review-comment-text", text: c.text }),
        reqs,
        saveReqs,
        reqLine ? el("div", { class: "review-reqs-line", text: reqLine }) : null,
      );
      for (const ch of c.changes) card.append(changeRow(ch));
      for (const b of c.blocked) if (!inFlightEdit(b)) card.append(blockedRow(c.ref, b));
      box.append(card);
    }
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
      box.append(
        el(
          "div",
          { class: "review-ledger-item" },
          el("strong", { text: `${item.id} ${item.kind}` }),
          el("span", { text: ` ${item.text} [${item.check.type}] ${statusText} ` }),
          check,
          edit,
          remove,
        ),
      );
    }
    $("review-ledger-add").textContent = state.editingLedger
      ? `Save ${state.editingLedger}`
      : "Add";
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
    const d = pendingDecisions(state.revision);
    const box = $("review-decisions");
    box.replaceChildren(
      el("div", { text: d.total ? `Waiting for you: ${d.total}` : "Nothing waits for you." }),
    );
    for (const x of [...d.blocked, ...d.ledger, ...d.requirements]) {
      box.append(el("div", { class: "review-decision", text: x.text }));
    }
  }

  // L4: changes made outside the tool since the last refresh.
  function renderExternal() {
    const box = $("review-external");
    box.replaceChildren();
    for (const x of state.revision?.external ?? []) {
      const where = x.paragraph_index ?? x.old_index;
      const label =
        `${x.id} ${x.kind} paragraph ${where}` +
        (x.lost_facts.length ? `. Lost: ${x.lost_facts.join("; ")}` : "");
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

  function render() {
    renderDraft();
    renderComments();
    renderLedger();
    renderCandidates();
    renderDecisions();
    renderExternal();
    showProgress();
    $("review-file").textContent = state.path ? `Review file: ${state.path}` : "";
  }

  // ---- actions --------------------------------------------------------------
  // Start the review turn of one comment. `release` (from onTurnStarted) ends
  // the turn taken for it when the request fails.
  // useSelection: send the paragraphs Preview prompt would show
  // (promptParagraphs); Propose all sends none, a selection belongs to one
  // comment.
  async function proposeRef(ref, release, useSelection) {
    state.reviewTurnOpen = true;
    try {
      const paragraphs = useSelection ? await promptParagraphs() : [];
      // A fresh snapshot first: outside changes are caught before the turn
      // (L4) and the new words of the turn are checked against it (L2).
      await refresh();
      const done = state.turnsDone;
      const r = await ask("review_propose", { ref, paragraphs });
      if (state.turnsDone === done && isTurnInFlight()) startProgress(ref, r.started_at);
      else if (state.turnsDone !== done) state.reviewTurnOpen = false; // its turn_done came first
    } catch (e) {
      state.reviewTurnOpen = false;
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
    else say(`Propose all: ${tail}`);
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
  }

  const refresh = () => refreshWith(say);
  // report(text, bad): where the counts go. The message line for a refresh
  // the author or a turn asked for, the small line for an automatic one.
  // Returns the invariant report (review-view.js invariantReport).
  async function refreshWith(report) {
    const snapshot = await word.collectReviewSnapshot();
    const r = await ask("review_refresh", { snapshot });
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
      ? ` Outside changes: ${r.external.length} (list under Comments and changes).`
      : "";
    report(
      `Accepted ${c.accepted}, rejected ${c.rejected}, edited ${c.user_edited}, pending ${c.pending}, not found ${c.unlocated}. ` +
        `Changes ${inv.changes}, expected tool comments ${inv.expected_comments}, found ${inv.tool_comments}` +
        invariantText(inv) +
        outside +
        (unresolved.length
          ? ` ${unresolved.length} rationale comment(s) not marked resolved (Word: ${unresolved[0]}; one cause is a document in compatibility mode).`
          : ""),
      !inv.ok,
    );
    return inv; // turn_done: a mismatch ends a Propose all run
  }

  // L10: changed paragraphs only unless full.
  async function runLedgerCheck(ids, full = false) {
    const snapshot = await word.collectReviewSnapshot();
    const r = await askJob("ledger_check", { ids, snapshot, full });
    const box = $("review-violations");
    box.replaceChildren();
    for (const res of r.outcomes) {
      for (const v of res.violations) {
        const label =
          v.paragraph_index == null
            ? `${res.id}: required text is gone`
            : `${res.id}: paragraph ${v.paragraph_index}${v.match ? ` "${v.match}"` : " (pending deletion)"}`;
        const b = el("button", { type: "button", class: "review-violation", text: label });
        if (v.paragraph_index != null)
          b.addEventListener(
            "click",
            guard(() => word.selectParagraphByIndex(v.paragraph_index)),
          );
        box.append(b);
      }
      if (res.status === "error")
        box.append(el("div", { class: "form-error", text: `${res.id}: ${res.error}` }));
    }
  }

  // P15: the review text from a .txt or .md file instead of a paste. The
  // file is read in the pane (File.text()); the daemon is not involved.
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
      if (r.warnings?.length) say(`Split by ${r.method}. ${r.warnings.join(" ")}`);
      else say(`Split by ${r.method}: ${r.comments.length} comment(s). Edit, then Save.`);
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
        };
      });
      const r = await ask("review_save_comments", { comments });
      state.draft = [];
      // Comments a running turn or an open edit still uses are not removed.
      if (r.warnings?.length) say(r.warnings.join(" "));
    }),
  );
  $("review-refresh").addEventListener("click", guard(refresh));
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
      const r = ledgerItemFromForm({
        kind: $("ledger-kind").value,
        text: $("ledger-text").value,
        scope: $("ledger-scope").value,
        checkType: $("ledger-check-type").value,
        pattern: $("ledger-pattern").value,
        flags: $("ledger-flags").value,
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
  $("review-clean-tool").addEventListener(
    "click",
    guard(async () => {
      const r = await word.deleteToolComments();
      await markCleaned(r.deleted_ids);
      const rep = cleanupReport(r.before, r.after);
      say(rep.text, !rep.ok);
    }),
  );
  // Delete ALL comments: first press shows the counts, second press reads
  // the comments again, backs up that fresh list and deletes the backed-up
  // threads that did not change meanwhile. If the list changed since it was
  // shown, nothing is deleted and the new count is shown for another press. window.confirm is unreliable in Office panes
  // (taskpane.js context-error note), so the confirmation is inline.
  $("review-delete-all").addEventListener(
    "click",
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
    if (!document.hidden) auto.trigger("focus");
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
              if (failed) return failed;
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
      } else if (msg.event === "late_result_recorded") {
        // A tracked edit that timed out was applied after all and is now recorded.
        say(`${msg.ref}: an edit that timed out was applied; recorded as ${msg.change_id}.`);
        load()
          .catch((e) => say(e.message, true))
          .finally(render);
      } else if (msg.event === "record_failed") {
        // The load turns the attempt into an unanswered edit in the list.
        load()
          .catch(() => {})
          .finally(() => {
            render();
            say(
              `${msg.ref}: the edit was applied in Word but saving its record failed; press Refresh status to match them again.`,
              true,
            );
          });
      } else if (msg.event === "stale_target") {
        // A tracked edit was refused: its paragraph is no longer as the
        // daemon has it. A refresh gives the daemon the document as it is.
        load()
          .then(refresh)
          .catch((e) => say(e.message, true))
          .finally(render);
      } else if (msg.event === "change_recorded" || msg.event === "edit_blocked") {
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
      (state.reviewTurnOpen = false),
      stopProgress(),
      load())
        .catch((e) => say(e.message, true))
        .finally(render),
  };
}
