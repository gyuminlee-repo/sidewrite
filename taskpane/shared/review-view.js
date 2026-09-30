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

export function countRevisionMarks(xml) {
  const s = String(xml ?? "");
  return { w_ins: (s.match(W_INS) || []).length, w_del: (s.match(W_DEL) || []).length };
}

// Same prefix as daemon/review-tools.mjs TOOL_COMMENT_PREFIX.
export const TOOL_COMMENT_RE = /^\s*\[R\d+\.\d+\]/;

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

export function buildCommentsBackup(comments, { manuscript = null, now = new Date() } = {}) {
  return {
    version: 1,
    manuscript,
    created_at: now.toISOString(),
    count: comments.length,
    comments: comments.map((c) => ({
      id: String(c.id),
      author: c.author ?? null,
      content: c.content ?? "",
      resolved: !!c.resolved,
      anchor_text: c.anchor_text ?? "",
      paragraph_index: c.paragraph_index ?? null,
    })),
  };
}

// Verdict for "Delete ALL comments": backup count and deleted count must both
// equal the count shown to the user, and nothing may remain.
export function deleteAllReport({ shown, backedUp, deleted, remaining }) {
  const ok = backedUp === shown && deleted === shown && remaining === 0;
  return {
    ok,
    text: `Backed up ${backedUp}, deleted ${deleted} of ${shown}, remaining ${remaining}${ok ? "" : " (MISMATCH)"}`,
  };
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
  const expected = changes.filter((ch) => !(ch.kind === "insert" && ch.state === "rejected"));
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

// Short labels shown on a change row (L1 lost facts, L2 unsourced, L3 narrowed).
export function changeFlags(ch) {
  const out = [];
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
        text: `${c.ref} blocked by ${b.blocked_by.join(", ")}: "${b.args?.replace ?? b.args?.insert_after ?? ""}"`,
      });
    }
    for (const r of unaddressedRequirements(c)) {
      requirements.push({
        ref: c.ref,
        id: r.id,
        text: `${c.ref} ${r.id} has no change: ${r.text}`,
      });
    }
  }
  const ledger = (revision?.ledger ?? [])
    .filter((i) => i.status === "unchecked" || i.status === "error")
    .map((i) => ({
      id: i.id,
      text: `${i.id} ${i.status === "error" ? "check failed" : "not checked yet"}`,
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
