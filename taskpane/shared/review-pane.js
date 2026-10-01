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
  reviewerOfRef,
  progressText,
  reviewTextFromFile,
  REVIEW_FILE_MAX_BYTES,
} from "./review-view.js";

const UNKNOWN_HINT =
  "Word did not answer; this exception may already be applied. Refresh status, then check the paragraph.";
// An override waits for the Word tool, which may take up to the daemon's 60 s
// tool timeout (bridge.mjs TOOL_TIMEOUT_MS). A split or a ledger check is a
// job instead (askJob): no request timeout covers its model calls.
const OVERRIDE_WAIT = 120_000;
const CONNECTION_LOST =
  "The connection to the daemon was lost before the result came. Run it again.";
const APPLYING_HINT = "Applying this exception in Word. Wait for it to finish.";

export function initReviewPane({ sendRequest, word, onTurnStarted, isTurnInFlight }) {
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
      box.append(el("div", { class: "review-draft-row" }, ref, body, reqs, drop));
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
    const label = el("span", { text: `Blocked by ${b.blocked_by.join(", ")}: "${newText}" ` });
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
      el("span", { class: "review-unknown", text: UNKNOWN_HINT + " " }),
      allow,
      reapply,
    );
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
      propose.disabled = isTurnInFlight();
      // The turn is taken on the click, before the selection and refresh()
      // waits: a second click meanwhile is ignored (Codex review 8) and the
      // daemon, which queues the agent message before it answers, may send
      // turn_complete and turn_done before the answer (Codex review 7).
      // onTurnStarted returns a release that ends only this turn. The check
      // runs before guard() so an ignored click keeps the message line.
      const run = guard(async (release) => {
        let r;
        try {
          const paragraphs = $("review-use-selection").checked
            ? await word.selectionParagraphs()
            : [];
          // A fresh snapshot first: outside changes are caught before the turn
          // (L4) and the new words of the turn are checked against it (L2).
          await refresh();
          const done = state.turnsDone;
          r = await ask("review_propose", { ref: c.ref, paragraphs });
          if (state.turnsDone === done && isTurnInFlight()) startProgress(c.ref, r.started_at);
        } catch (e) {
          release?.();
          throw e;
        }
      });
      propose.addEventListener("click", () => {
        if (isTurnInFlight()) return;
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
          const r = await ask("review_preview", { ref: c.ref });
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
        el("div", { class: "review-comment-text", text: c.text }),
        reqs,
        saveReqs,
        reqLine ? el("div", { class: "review-reqs-line", text: reqLine }) : null,
      );
      for (const ch of c.changes) card.append(changeRow(ch));
      for (const b of c.blocked) card.append(blockedRow(c.ref, b));
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

  // L9: elapsed time of the running turn and comments not done.
  function showProgress() {
    const bar = $("review-progress");
    bar.hidden = !state.turn;
    if (state.turn) {
      bar.textContent = progressText({ ...state.turn, comments: state.revision?.comments ?? [] });
    }
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
    renderDecisions();
    renderExternal();
    showProgress();
    $("review-file").textContent = state.path ? `Review file: ${state.path}` : "";
  }

  // ---- actions --------------------------------------------------------------
  async function load() {
    const r = await ask("review_load");
    state.path = r.path;
  }

  async function refresh() {
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
    say(
      `Accepted ${c.accepted}, rejected ${c.rejected}, edited ${c.user_edited}, pending ${c.pending}, not found ${c.unlocated}. ` +
        `Changes ${inv.changes}, expected tool comments ${inv.expected_comments}, found ${inv.tool_comments}` +
        (inv.ok
          ? "."
          : ` (MISMATCH: missing ${inv.missing_comment.join(",") || "-"}, unrecorded ${inv.unrecorded_comment_ids.join(",") || "-"})`) +
        outside +
        (unresolved.length
          ? ` ${unresolved.length} rationale comment(s) not marked resolved (Word: ${unresolved[0]}; one cause is a document in compatibility mode).`
          : ""),
      !inv.ok,
    );
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

  render(); // loaded on "welcome" (taskpane.js calls reload once connected)

  return {
    onEvent(msg) {
      if (msg.event === "job_done") {
        jobDone(msg);
      } else if (msg.event === "turn_done") {
        state.turnsDone += 1;
        stopProgress();
        load()
          .then(refresh)
          .catch((e) => say(e.message, true))
          .finally(render);
      } else if (msg.event === "late_result_recorded") {
        // A tracked edit that timed out was applied after all and is now recorded.
        say(`${msg.ref}: an edit that timed out was applied; recorded as ${msg.change_id}.`);
        load()
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
      for (const b of proposeButtons) b.disabled = busy;
    },
    // A new connection (welcome) means no turn of this pane is running any
    // more: a daemon restart never sends turn_done, so clear the progress line.
    // Jobs of the old connection are ended with an error first.
    reload: () =>
      (dropJobs(), stopProgress(), load()).catch((e) => say(e.message, true)).finally(render),
  };
}
