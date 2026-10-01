// Pure helpers for the Review tab. No Office.js and no DOM here, so
// node --test can exercise them (tests/review-view.test.mjs).

// Same rule as daemon/revision-classify.mjs normalizeParaText; the test
// runs both on one corpus so they cannot drift apart.
export function normalizeParaText(s) {
  return String(s ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

// The paragraph as it reads once this one change is accepted. Computed from
// the text before the edit, because while the change is pending the Word
// paragraph.text holds both the inserted and the deleted text.
export function computeAfterText(before, find, { replace, insert_after } = {}) {
  const b = normalizeParaText(before);
  // Keep leading/trailing spaces of `find`: " with shaking" deletes the space too.
  const f = String(find ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ");
  if (!f.trim()) return { error: "find is empty" };
  const i = b.indexOf(f);
  if (i === -1) return { error: "find is not in the paragraph text" };
  if (b.indexOf(f, i + 1) !== -1) return { error: "find occurs more than once in the paragraph" };
  if (typeof insert_after === "string") {
    return {
      after: normalizeParaText(b.slice(0, i + f.length) + insert_after + b.slice(i + f.length)),
    };
  }
  if (typeof replace === "string") {
    return { after: normalizeParaText(b.slice(0, i) + replace + b.slice(i + f.length)) };
  }
  return { error: "give replace or insert_after" };
}

// `[\s>/]` after the name excludes w:instrText, w:delText, w:delInstrText.
const W_INS = /<w:ins[\s>/]/g;
const W_DEL = /<w:del[\s>/]/g;

// getOoxml() of a paragraph is a pkg:package. Only the document part holds the
// paragraph's own revision marks: a rationale comment inserted while tracking
// was on carries a w:ins in /word/comments.xml that accepting the change in
// Word does not remove, and it made an accepted insertion read as pending.
const DOC_PART = '<pkg:part pkg:name="/word/document.xml"';

export function countRevisionMarks(xml) {
  let s = String(xml ?? "");
  const at = s.indexOf(DOC_PART);
  if (at !== -1) {
    const end = s.indexOf("</pkg:part>", at);
    s = s.slice(at, end === -1 ? undefined : end);
  }
  return { w_ins: (s.match(W_INS) || []).length, w_del: (s.match(W_DEL) || []).length };
}

// Same prefix as daemon/review-tools.mjs TOOL_COMMENT_PREFIX.
export const TOOL_COMMENT_RE = /^\s*\[(?:R\d+|E)\.\d+\]/;

// Reviewer number of a comment ref: R<n>.<m> -> n, E.<m> (editor) -> 0,
// anything else -> null.
export function reviewerOfRef(ref) {
  const m = /^(?:R(\d+)|E)\.\d+$/.exec(String(ref ?? ""));
  if (!m) return null;
  return m[1] === undefined ? 0 : Number(m[1]);
}

export function isToolComment(content) {
  return TOOL_COMMENT_RE.test(String(content ?? ""));
}

export function countComments(comments) {
  let tool = 0;
  for (const c of comments ?? []) if (isToolComment(c.content)) tool += 1;
  return { tool, user: (comments?.length ?? 0) - tool, total: comments?.length ?? 0 };
}

// Result line for "Remove tool comments": tool comments must reach 0 and
// the comments the user wrote must not change.
export function cleanupReport(before, after) {
  const ok = after.tool === 0 && after.user === before.user;
  return {
    ok,
    text:
      `Tool comments ${before.tool} -> ${after.tool}, your comments ${before.user} -> ${after.user}` +
      (ok ? " (unchanged)" : " (MISMATCH)"),
  };
}

// Deleting a comment deletes its replies, so the backup keeps them.
// count is the number of parent comments (what Word lists and deletes);
// reply_count is the number of replies across all threads.
export function buildCommentsBackup(comments, { manuscript = null, now = new Date() } = {}) {
  const rows = comments.map((c) => ({
    id: String(c.id),
    author: c.author ?? null,
    content: c.content ?? "",
    resolved: !!c.resolved,
    anchor_text: c.anchor_text ?? "",
    paragraph_index: c.paragraph_index ?? null,
    replies: (c.replies ?? []).map((r) => ({
      id: r.id == null ? null : String(r.id),
      author: r.author ?? null,
      content: r.content ?? "",
      created_at: r.created_at ?? null,
    })),
  }));
  return {
    version: 1,
    manuscript,
    created_at: now.toISOString(),
    count: rows.length,
    reply_count: rows.reduce((n, c) => n + c.replies.length, 0),
    comments: rows,
  };
}

// "Delete ALL comments" shows a list on the first press and deletes on the
// second. Before deleting, the pane reads the comments again and compares:
// any comment added, removed, or changed (text, author, resolved, anchor,
// replies) since the list was shown means no delete and a new confirmation.
function commentKey(c) {
  return JSON.stringify([
    c.author ?? null,
    c.content ?? "",
    !!c.resolved,
    c.anchor_text ?? "",
    (c.replies ?? []).map((r) => [r.author ?? null, r.content ?? ""]),
  ]);
}

export function compareCommentLists(shown, current) {
  const before = new Map((shown ?? []).map((c) => [String(c.id), commentKey(c)]));
  const now = new Map((current ?? []).map((c) => [String(c.id), commentKey(c)]));
  let added = 0;
  let edited = 0;
  for (const [id, key] of now) {
    if (!before.has(id)) added += 1;
    else if (before.get(id) !== key) edited += 1;
  }
  let removed = 0;
  for (const id of before.keys()) if (!now.has(id)) removed += 1;
  const same = added + removed + edited === 0;
  const parts = [];
  if (added) parts.push(`${added} added`);
  if (removed) parts.push(`${removed} removed`);
  if (edited) parts.push(`${edited} changed`);
  return { same, added, removed, edited, text: parts.join(", ") };
}

// The second "Delete ALL" press deletes only the threads in the backup, and
// only those still equal to their backed-up copy (text, author, resolved,
// anchor, replies). The Office function reads the comments and deletes in
// one Word.run with this decision in between. new_ids: comments added after
// the backup (kept); changed_ids: backed-up threads that changed (kept, the
// user presses again); missing_ids: backed-up threads already gone.
export function planBackupDelete(backupRows, current) {
  const backed = new Map((backupRows ?? []).map((c) => [String(c.id), commentKey(c)]));
  const plan = { delete_ids: [], changed_ids: [], new_ids: [], missing_ids: [] };
  const seen = new Set();
  for (const c of current ?? []) {
    const id = String(c.id);
    if (!backed.has(id)) {
      plan.new_ids.push(id);
      continue;
    }
    seen.add(id);
    (backed.get(id) === commentKey(c) ? plan.delete_ids : plan.changed_ids).push(id);
  }
  for (const id of backed.keys()) if (!seen.has(id)) plan.missing_ids.push(id);
  return plan;
}

// Verdict for "Delete ALL comments": the backup holds every comment shown;
// every backed-up thread is deleted or was already gone; none of the deleted
// ids is still in the document (stillThere, read again afterwards).
// Comments added after the backup are kept and named, not a mismatch. A
// backed-up thread that changed after the backup is kept and needs another
// press, so the result is not ok.
export function deleteAllReport({
  shown,
  backedUp,
  deleted,
  changed = 0,
  missing = 0,
  addedAfter = 0,
  stillThere = 0,
}) {
  const ok = backedUp === shown && changed === 0 && deleted + missing === shown && stillThere === 0;
  const parts = [`Backed up ${backedUp}, deleted ${deleted} of ${shown}`];
  if (missing) parts.push(`${missing} already gone`);
  if (changed) parts.push(`${changed} changed after the backup and not deleted (press again)`);
  if (stillThere) parts.push(`${stillThere} deleted comment(s) still present`);
  let text = parts.join(", ") + (ok ? "" : " (MISMATCH)");
  if (addedAfter) text += `. ${addedAfter} comment(s) added after the backup were not deleted`;
  return { ok, text };
}

// Recorded changes vs tool comments now in the document. Every change keeps
// its comment (a deleted-range comment survives accept and reject), except a
// rejected pure insertion, whose comment Word removes with the text. An
// extra [R comment means an edit that was not recorded.
export function invariantReport(revision, comments) {
  const ids = new Set(
    (comments ?? []).filter((c) => isToolComment(c.content)).map((c) => String(c.id)),
  );
  const changes = (revision?.comments ?? []).flatMap((c) => c.changes ?? []);
  // No comment is expected for a rejected insertion (Word removes it) or
  // for a change whose comment the tool comment cleanup removed.
  const expected = changes.filter(
    (ch) => !(ch.kind === "insert" && ch.state === "rejected") && !ch.comment_cleaned_at,
  );
  const recorded = new Set(changes.map((ch) => String(ch.word_comment_id)));
  const missing = expected.filter((ch) => !ids.has(String(ch.word_comment_id)));
  const orphans = [...ids].filter((id) => !recorded.has(id));
  return {
    changes: changes.length,
    expected_comments: expected.length,
    tool_comments: ids.size,
    missing_comment: missing.map((ch) => ch.id),
    unrecorded_comment_ids: orphans,
    ok: expected.length === ids.size && missing.length === 0 && orphans.length === 0,
  };
}

// Design 0 C1: once the author accepts or rejects a change, its rationale
// comment is marked resolved. Returns ids of open comments to resolve.
export function commentsToResolve(rows, revision, comments) {
  const open = new Set((comments ?? []).filter((c) => !c.resolved).map((c) => String(c.id)));
  const byId = new Map();
  for (const c of revision?.comments ?? []) for (const ch of c.changes ?? []) byId.set(ch.id, ch);
  const out = [];
  for (const r of rows ?? []) {
    if (r.state !== "accepted" && r.state !== "rejected") continue;
    const id = byId.get(r.change_id)?.word_comment_id;
    if (id != null && open.has(String(id))) out.push(String(id));
  }
  return out;
}

export const STATE_LABELS = {
  pending: "Pending",
  accepted: "Accepted",
  rejected: "Rejected",
  user_edited: "Edited by you",
  unlocated: "Not found",
};

export const STATUS_LABELS = { pending: "To do", in_progress: "In progress", done: "Done" };

// Ledger status shown on an item row. A skipped item is outside the
// manuscript scope and is checked at its own stage, not counted as clean.
export const SKIPPED_SCOPE_LABELS = {
  response: "Checked at the response stage",
  supplement: "Checked at the supplement stage",
};

const unsureText = (item) => `unsure on ${item.unsure.length} paragraph(s)`;

export function ledgerStatusText(item) {
  if (item?.status === "violations")
    return item.unsure?.length
      ? `${item.violation_count} violation(s), ${unsureText(item)}`
      : `${item.violation_count} violation(s)`;
  if (item?.status === "unsure") return unsureText(item);
  if (item?.status === "skipped")
    return SKIPPED_SCOPE_LABELS[item.scope] ?? `Checked at the ${item.scope} stage`;
  return item?.status;
}

// Parse the ledger form into the ledger item shape of the daemon.
export function ledgerItemFromForm({ kind, text, scope, checkType, pattern, flags, expect }) {
  const t = String(text ?? "").trim();
  if (!t) return { error: "Write the instruction text." };
  if (checkType === "regex") {
    const p = String(pattern ?? "");
    if (!p) return { error: "A regex check needs a pattern." };
    try {
      new RegExp(p, String(flags ?? "").replace(/g/g, ""));
    } catch (e) {
      return { error: `Invalid regex: ${e.message}` };
    }
    return {
      item: {
        kind,
        text: t,
        scope: scope || "manuscript",
        check: {
          type: "regex",
          pattern: p,
          flags: String(flags ?? "").replace(/g/g, ""),
          expect: expect || "absent",
        },
      },
    };
  }
  return { item: { kind, text: t, scope: scope || "manuscript", check: { type: "llm" } } };
}

// ---- L1-L3 badges, L5 requirements, L6 ledger drafts, L7, L9 ----------------

// Short labels shown on a change row (L1 lost facts, L2 unsourced, L3 narrowed)
// and a change applied without its rationale comment.
export function changeFlags(ch) {
  const out = [];
  if (ch?.comment_missing) out.push({ kind: "comment-missing", text: "No rationale comment" });
  if (ch?.lost_facts?.length) out.push({ kind: "lost", text: `Lost: ${ch.lost_facts.join("; ")}` });
  if (ch?.unsourced_terms?.length)
    out.push({ kind: "unsourced", text: `No source: ${ch.unsourced_terms.join(", ")}` });
  if (ch?.narrowed_from) out.push({ kind: "narrowed", text: "Narrowed" });
  return out;
}

// One requirement per non-empty line of the text box.
export function requirementsFromText(text) {
  return String(text ?? "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
}

// L5. Requirements of a comment that no change answers. A rejected change
// answers nothing.
export function unaddressedRequirements(comment) {
  const used = new Set();
  for (const ch of comment?.changes ?? []) {
    if (ch.state === "rejected") continue;
    for (const id of ch.rationale?.requirement_ids ?? []) used.add(id);
  }
  return (comment?.requirements ?? []).filter((r) => !used.has(r.id));
}

const escapeRegex = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// L6. Ledger form fields drafted from a change the author rejected: new text
// that was rejected becomes a forbid rule; a rejected deletion becomes a
// "must stay" rule on the deleted text. The author edits the draft before Add.
export function ledgerDraftFromChange(ch) {
  const newText = String(ch?.text ?? "").trim();
  const tag = ch?.id ? ` (rejected ${ch.id})` : "";
  if (ch?.kind === "delete" || !newText) {
    const kept = String(ch?.find ?? "").trim();
    return {
      kind: "instruction",
      text: `Keep "${kept}"${tag}`,
      scope: "manuscript",
      checkType: "regex",
      pattern: escapeRegex(kept),
      flags: "",
      expect: "present",
    };
  }
  return {
    kind: "forbid",
    text: `Do not write "${newText}"${tag}`,
    scope: "manuscript",
    checkType: "regex",
    pattern: escapeRegex(newText),
    flags: "i",
    expect: "absent",
  };
}

// L6. Form fields of an existing item, for editing it (ledger_update).
export function ledgerFormFromItem(item) {
  const c = item?.check ?? {};
  return {
    kind: item?.kind ?? "instruction",
    text: item?.text ?? "",
    scope: item?.scope ?? "manuscript",
    checkType: c.type === "regex" ? "regex" : "llm",
    pattern: c.type === "regex" ? c.pattern : "",
    flags: c.type === "regex" ? (c.flags ?? "") : "",
    expect: c.type === "regex" ? (c.expect ?? "absent") : "absent",
  };
}

// L6. Draft a check from what the author typed: a quoted phrase ("...",
// '...', or corner brackets) becomes a literal regex; without one the draft
// is an LLM check. Whether the phrase is forbidden or must stay is left to
// the author (the form keeps its current choice).
export function ledgerCheckFromText(text) {
  const m = /"([^"]+)"|'([^']+)'|「([^」]+)」|“([^”]+)”/.exec(String(text ?? ""));
  const phrase = m ? (m[1] ?? m[2] ?? m[3] ?? m[4]).trim() : "";
  return phrase
    ? { checkType: "regex", pattern: escapeRegex(phrase), flags: "i" }
    : { checkType: "llm", pattern: "", flags: "" };
}

// L7. Everything that waits for the author, in one list.
export function pendingDecisions(revision) {
  const blocked = [];
  const requirements = [];
  for (const c of revision?.comments ?? []) {
    for (const b of c.blocked ?? []) {
      blocked.push({
        ref: c.ref,
        id: b.id,
        // No forbid rule: a normal edit Word did not answer (outcome unknown).
        text: `${c.ref} ${b.blocked_by.length ? `blocked by ${b.blocked_by.join(", ")}` : "edit with no answer from Word"}: "${b.args?.replace ?? b.args?.insert_after ?? ""}"`,
      });
    }
    for (const r of unaddressedRequirements(c)) {
      requirements.push({
        ref: c.ref,
        id: r.id,
        text: `${c.ref}${c.section ? ` (${c.section})` : ""} ${r.id} has no change: ${r.text}`,
      });
    }
  }
  // An unsure paragraph (the LLM check abstained) is the author's call.
  const ledger = (revision?.ledger ?? [])
    .filter((i) => i.status === "unchecked" || i.status === "error" || i.unsure?.length)
    .map((i) => ({
      id: i.id,
      text: i.unsure?.length
        ? `${i.id} unsure on paragraph(s) ${i.unsure.map((u) => u.paragraph_index).join(", ")}`
        : `${i.id} ${i.status === "error" ? "check failed" : "not checked yet"}`,
    }));
  return {
    blocked,
    ledger,
    requirements,
    total: blocked.length + ledger.length + requirements.length,
  };
}

// L9. "R1.3 running 1:05. 2 of 3 comments not done."
export function progressText({ ref, startedAt, now = Date.now(), comments = [] }) {
  const t0 = typeof startedAt === "string" ? Date.parse(startedAt) : startedAt;
  const s = Math.max(0, Math.floor((now - t0) / 1000));
  const clock = `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
  const open = comments.filter((c) => c.status !== "done").length;
  return `${ref} running ${clock}. ${open} of ${comments.length} comments not done.`;
}

// P15. The comments "Propose all" runs, in list order: every one not done.
export function queueRefs(revision) {
  return (revision?.comments ?? []).filter((c) => c.status !== "done").map((c) => c.ref);
}

// ---- P15: review text from a file --------------------------------------------

export const REVIEW_FILE_MAX_BYTES = 2 * 1024 * 1024;

// A picked review file ({ name, size, text }) as text for the paste box:
// .txt or .md only, at most 2 MB, BOM removed, line ends as \n.
export function reviewTextFromFile({ name, size, text }) {
  if (!/\.(txt|md)$/i.test(String(name ?? ""))) {
    return { ok: false, error: `${name}: choose a .txt or .md file.` };
  }
  if (size > REVIEW_FILE_MAX_BYTES) {
    return { ok: false, error: `${name} is larger than 2 MB.` };
  }
  const t = String(text ?? "")
    .replace(/^\uFEFF/, "")
    .replace(/\r\n?/g, "\n");
  if (!t.trim()) return { ok: false, error: `${name} is empty.` };
  return { ok: true, text: t };
}

// ---- P15: ledger file --------------------------------------------------------

// The text of a picked ledger file as the ledger_import payload. Only the
// outer shape is checked here ({ items: [...] }, an exported ledger or a
// hand-written list); the daemon checks every rule.
export function ledgerImportFromJson(text) {
  let data;
  try {
    data = JSON.parse(String(text ?? "").replace(/^\uFEFF/, ""));
  } catch (e) {
    return { ok: false, error: `The ledger file is not JSON (${e.message}).` };
  }
  if (!Array.isArray(data?.items)) {
    return { ok: false, error: 'The ledger file needs an "items" list.' };
  }
  if (!data.items.length) return { ok: false, error: "The ledger file has no items." };
  return { ok: true, doc: { items: data.items }, count: data.items.length };
}

// ---- P15: ledger rule candidates from the review text ---------------------------

// The rule fields of a candidate (ledger_suggest), as ledger_import takes them.
export function candidateRule({ kind, text, scope, check }) {
  return { kind, text, scope, check };
}

// "Forbidden, must not match /novel/i: <sentence>"
export function candidateLabel(c) {
  const kind = c.kind === "forbid" ? "Forbidden" : "Instruction";
  const k = c.check ?? {};
  const check =
    k.type === "regex"
      ? `${k.expect === "present" ? "must keep" : "must not match"} /${k.pattern}/${k.flags ?? ""}`
      : "LLM check";
  return `${kind}, ${check}: ${c.text}`;
}
