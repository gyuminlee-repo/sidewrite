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
  pendingDecisions,
  progressText,
} from "./review-view.js";

const LONG = 120_000; // LLM split and LLM ledger checks exceed the 10 s default

export function initReviewPane({ sendRequest, word, onTurnStarted, isTurnInFlight }) {
  const root = document.getElementById("review-root");
  if (!root) return { onEvent() {} }; // Excel pane: no Review tab
  const $ = (id) => document.getElementById(id);
  const state = {
    revision: null,
    path: null,
    draft: [],
    backupShown: null,
    editingLedger: null, // L6: id of the ledger item loaded into the form
    confirmDone: null, // L5: ref whose "Done" was asked once with open requirements
    armedBlocked: new Set(), // ids of blocked edits whose "Allow exception" was pressed once
    turn: null, // L9: { ref, startedAt } of the running review turn
    timer: null,
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

  function blockedRow(ref, b) {
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
        const r = await ask("review_override", { ref, blocked_id: b.id });
        if (!r.applied) throw new Error(r.result?.message || "Not applied");
      }),
    );
    const newText = b.args.replace ?? b.args.insert_after ?? "";
    return el(
      "div",
      { class: "review-blocked" },
      el("span", { text: `Blocked by ${b.blocked_by.join(", ")}: "${newText}" ` }),
      allow,
    );
  }

  function renderComments() {
    const box = $("review-comments");
    box.replaceChildren();
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
      propose.disabled = isTurnInFlight();
      propose.addEventListener(
        "click",
        guard(async () => {
          const paragraphs = $("review-use-selection").checked
            ? await word.selectionParagraphs()
            : [];
          // A fresh snapshot first: outside changes are caught before the turn
          // (L4) and the new words of the turn are checked against it (L2).
          await refresh();
          const r = await ask("review_propose", { ref: c.ref, paragraphs });
          startProgress(c.ref, r.started_at);
          onTurnStarted(`Review ${c.ref}: proposing changes (tracked, one comment per change)...`);
        }),
      );
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
      reqs.value = (c.requirements ?? []).map((r) => r.text).join("\n");
      const saveReqs = el("button", {
        type: "button",
        class: "btn-secondary btn-small",
        text: "Save requirements",
      });
      saveReqs.addEventListener(
        "click",
        guard(() =>
          ask("review_set_requirements", {
            ref: c.ref,
            requirements: requirementsFromText(reqs.value),
          }),
        ),
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
      const statusText =
        item.status === "violations" ? `${item.violation_count} violation(s)` : item.status;
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
    for (const id of commentsToResolve(r.rows, state.revision, snapshot.comments)) {
      await word.resolveComment(id);
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
        outside,
      !inv.ok,
    );
  }

  // L10: changed paragraphs only unless full.
  async function runLedgerCheck(ids, full = false) {
    const snapshot = await word.collectReviewSnapshot();
    const r = await ask("ledger_check", { ids, snapshot, full }, LONG);
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

  $("review-split").addEventListener(
    "click",
    guard(async () => {
      const r = await ask("review_split", { text: $("review-input").value }, LONG);
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
        const m = /^R(\d+)\.(\d+)$/.exec(c.ref);
        if (!m) throw new Error(`Bad ref "${c.ref}": use R<reviewer>.<number>`);
        return {
          ref: c.ref,
          reviewer: Number(m[1]),
          text: c.text.trim(),
          requirements: requirementsFromText(c.requirementsText),
        };
      });
      await ask("review_save_comments", { comments });
      state.draft = [];
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
  $("review-clean-tool").addEventListener(
    "click",
    guard(async () => {
      const r = await word.deleteToolComments();
      const rep = cleanupReport(r.before, r.after);
      say(rep.text, !rep.ok);
    }),
  );
  // Delete ALL comments: first press shows the counts, second press reads
  // the comments again, backs up that fresh list and deletes. If the list
  // changed since it was shown, nothing is deleted and the new count is
  // shown for another press. window.confirm is unreliable in Office panes
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
      const cleared = await word.clearAllComments();
      const left = await word.listComments();
      const rep = deleteAllReport({
        shown: current.length,
        backedUp: b.count,
        deleted: cleared.cleared_comments,
        remaining: left.comment_count,
      });
      say(`${rep.text}. Backup: ${b.path}`, !rep.ok);
    }),
  );

  render(); // loaded on "welcome" (taskpane.js calls reload once connected)

  return {
    onEvent(msg) {
      if (msg.event === "turn_done") {
        stopProgress();
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
    // A new connection (welcome) means no turn of this pane is running any
    // more: a daemon restart never sends turn_done, so clear the progress line.
    reload: () => (stopProgress(), load()).catch((e) => say(e.message, true)).finally(render),
  };
}
