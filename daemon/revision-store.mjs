// Per-manuscript review state: `<manuscript stem>.revision.json` next to the
// document. Holds the reviewer comments, every tracked edit the review flow
// made (with its rationale comment id and before/after sentences), edits
// that the forbid list refused, and the instruction ledger.
//
// Writes are atomic (temp file in the same folder, then rename) and
// serialized per file inside this process, so two tracked edits that land
// back to back cannot lose each other.

import { readFile, writeFile, rename, unlink } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { z } from "zod";
import { toLocalPath } from "./workspace.mjs";

export const REVISION_VERSION = 1;

const RegexCheck = z.object({
  type: z.literal("regex"),
  pattern: z.string().min(1),
  flags: z
    .string()
    .regex(/^[imsuy]*$/, "flags may use i, m, s, u, y (g is added by the checker)")
    .default(""),
  // absent: every match is a violation. present: the pattern must still
  // match somewhere (a "do not delete X" rule).
  expect: z.enum(["absent", "present"]).default("absent"),
});
const LlmCheck = z.object({ type: z.literal("llm") });

const Rationale = z.object({
  what: z.string().min(1),
  why: z.string().min(1),
  addresses: z.string().min(1),
  // L5: ids of the comment requirements this change answers (q1, q2 ...).
  requirement_ids: z.array(z.string()).default([]),
});

export const ChangeSchema = z.object({
  id: z.string(),
  word_comment_id: z.string().nullable(),
  paragraph_index: z.number().int().nullable(),
  paragraph_id: z.string().nullable().default(null),
  kind: z.enum(["replace", "delete", "insert"]),
  find: z.string(),
  // Replacement text (replace), "" (delete), or the inserted text (insert).
  text: z.string(),
  before: z.string(),
  after: z.string(),
  rationale: Rationale,
  override_of: z.array(z.string()).default([]),
  // L1: facts in the deleted text that the new text does not keep.
  lost_facts: z.array(z.string()).default([]),
  // L2: new facts and words found neither in the manuscript nor in the comment.
  unsourced_terms: z.array(z.string()).default([]),
  // L3: the find/replace the model asked for, when the applied range is smaller.
  narrowed_from: z.object({ find: z.string(), text: z.string() }).nullable().default(null),
  // Up to 3 words before and after the changed range at edit time, so one
  // change can be judged apart from others in the same paragraph. null for
  // changes recorded before this field (whole-paragraph rule).
  context: z.object({ left: z.string(), right: z.string() }).nullable().default(null),
  // The edit was applied but its rationale comment was not anchored: the
  // task pane error code (CommentAnchorMissing or CommentInsertFailed).
  // word_comment_id is then null.
  comment_missing: z.string().nullable().default(null),
  // When the task pane's comment cleanup (tool comments or Delete ALL)
  // removed the rationale comment: its absence is then no sign of a reject
  // and no missing comment (Codex review 7).
  comment_cleaned_at: z.string().nullable().default(null),
  created_at: z.string(),
  state: z.enum(["pending", "accepted", "rejected", "user_edited", "unlocated"]).default("pending"),
  checked_at: z.string().nullable().default(null),
});

// An approved exception whose tracked edit got no answer from the task pane
// (OutcomeUnknown): the change it would record and the document form it
// expects, so a late answer or a status refresh can record it once.
const Attempt = z.object({
  tool_call_id: z.string().nullable(),
  unknown_at: z.string(),
  comment_text: z.string(),
  change: z.record(z.string(), z.unknown()),
});

const Blocked = z.object({
  id: z.string(),
  args: z.record(z.string(), z.unknown()),
  blocked_by: z.array(z.string()).min(1),
  created_at: z.string(),
  // outcome_unknown: an override was sent and not answered; it may be in the
  // document, so a plain second override is refused (review-service.mjs).
  // applying: an override is running now; any second override is refused.
  state: z.enum(["blocked", "applying", "outcome_unknown"]).default("blocked"),
  attempt: Attempt.nullable().default(null),
});

// L5: one thing the reviewer asks for, e.g. { id: "q1", text: "state the unit" }.
const Requirement = z.object({
  id: z.string().regex(/^q\d+$/),
  text: z.string().min(1),
});

const Comment = z.object({
  // R<reviewer>.<n>; E.<n> is the editor letter with reviewer 0.
  ref: z.string().regex(/^(?:R\d+|E)\.\d+$/),
  reviewer: z.number().int().nonnegative(),
  text: z.string(),
  requirements: z.array(Requirement).default([]),
  status: z.enum(["pending", "in_progress", "done"]).default("pending"),
  changes: z.array(ChangeSchema).default([]),
  blocked: z.array(Blocked).default([]),
});

// A paragraph the LLM check could not judge (abstain). Not a violation; it
// waits for the author and is read again by the next incremental check.
const Unsure = z.object({
  paragraph_index: z.number().int(),
  h: z.string().nullable(),
  reason: z.string().default(""),
});

const Violation = z.object({
  paragraph_index: z.number().int().nullable(),
  h: z.string().nullable(),
  start: z.number().int().nullable().default(null),
  match: z.string().nullable().default(null),
  reason: z.string().optional(),
  missing: z.boolean().optional(),
});

export const LedgerItemSchema = z.object({
  id: z.string().regex(/^L\d+$/),
  kind: z.enum(["instruction", "forbid"]),
  text: z.string().min(1),
  scope: z.enum(["manuscript", "response", "supplement"]).default("manuscript"),
  check: z.discriminatedUnion("type", [RegexCheck, LlmCheck]),
  // skipped: outside the manuscript scope, not checked against the manuscript.
  // unsure: an LLM check could not judge some paragraphs and found no violation.
  status: z
    .enum(["unchecked", "ok", "violations", "error", "skipped", "unsure"])
    .default("unchecked"),
  violation_count: z.number().int().nonnegative().default(0),
  checked_at: z.string().nullable().default(null),
  // L10: violations of the last check keyed by paragraph hash, and the
  // hashes of every paragraph that check covered (null: never checked).
  violations: z.array(Violation).default([]),
  unsure: z.array(Unsure).default([]),
  checked_hashes: z.array(z.string()).nullable().default(null),
});

const External = z.object({
  id: z.string().regex(/^X\d+$/),
  detected_at: z.string(),
  kind: z.enum(["modified", "inserted", "deleted"]),
  paragraph_index: z.number().int().nullable(),
  old_index: z.number().int().nullable(),
  before: z.string(),
  after: z.string(),
  lost_facts: z.array(z.string()).default([]),
});

export const RevisionSchema = z.object({
  version: z.literal(REVISION_VERSION),
  manuscript: z.string(),
  updated_at: z.string(),
  comments: z.array(Comment).default([]),
  ledger: z.array(LedgerItemSchema).default([]),
  // L4: paragraph texts (reviewed view) at the last status refresh, and the
  // changes found since then that no review change explains.
  snapshot: z
    .object({ taken_at: z.string(), paragraphs: z.array(z.string()) })
    .nullable()
    .default(null),
  external: z.array(External).default([]),
});

// ---- paths ----------------------------------------------------------------

// Local path of the document in this pane, or null when the daemon cannot
// open: unsaved document, SharePoint/OneDrive URL, UNC share outside WSL.
export function localDocPath(activeDoc, env = process.env) {
  if (typeof activeDoc !== "string" || !activeDoc.trim()) return null;
  const local = toLocalPath(activeDoc.trim(), env);
  if (local === null) return null;
  if (/^file:\/\//i.test(local)) {
    try {
      return fileURLToPath(local);
    } catch {
      return null;
    }
  }
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(local)) return null;
  return local;
}

const DOC_EXT = /\.(docx?|docm|dotx?|dotm)$/i;

export function sidecarPath(docPath, suffix) {
  const stem = basename(docPath).replace(DOC_EXT, "") || "document";
  return join(dirname(docPath), `${stem}${suffix}`);
}

export const revisionPathFor = (docPath) => sidecarPath(docPath, ".revision.json");
export const commentsBackupPathFor = (docPath) => sidecarPath(docPath, ".comments_backup.json");
// LLM decision receipts (decision-log.mjs), append only.
export const decisionsPathFor = (docPath) => sidecarPath(docPath, ".decisions.jsonl");

// ---- load / save ------------------------------------------------------------

export function emptyRevision(manuscript, now = new Date()) {
  return {
    version: REVISION_VERSION,
    manuscript,
    updated_at: now.toISOString(),
    comments: [],
    ledger: [],
    snapshot: null,
    external: [],
  };
}

export async function loadRevision(path, manuscript) {
  let raw;
  try {
    raw = await readFile(path, "utf8");
  } catch (e) {
    if (e.code === "ENOENT") return emptyRevision(manuscript);
    throw e;
  }
  return RevisionSchema.parse(JSON.parse(raw));
}

export async function writeJsonAtomic(path, data) {
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
  await writeFile(tmp, JSON.stringify(data, null, 2) + "\n", "utf8");
  try {
    await rename(tmp, path);
  } catch (e) {
    await unlink(tmp).catch(() => {});
    throw e;
  }
}

export async function saveRevision(path, revision, now = new Date()) {
  const data = RevisionSchema.parse({ ...revision, updated_at: now.toISOString() });
  await writeJsonAtomic(path, data);
  return data;
}

const chains = new Map();

// Load, apply `fn`, save; one at a time per path. `fn` returns the next
// revision, or null to leave the file untouched.
export function withRevision(path, manuscript, fn) {
  const prev = chains.get(path) ?? Promise.resolve();
  const run = prev
    .catch(() => {})
    .then(async () => {
      const current = await loadRevision(path, manuscript);
      const next = await fn(current);
      return next ? saveRevision(path, next) : current;
    });
  chains.set(path, run);
  return run;
}

// ---- pure operations (return a new revision) --------------------------------

const clone = (x) => structuredClone(x);

export function findComment(rev, ref) {
  return rev.comments.find((c) => c.ref === ref) ?? null;
}

function requireComment(rev, ref) {
  const c = findComment(rev, ref);
  if (!c) throw new Error(`No reviewer comment ${ref}`);
  return c;
}

// Replace the comment list with the user-edited one. Changes, blocked edits
// and status of a ref that survives are kept. Dropping a ref that already
// has changes is refused: its edits are in the document. A dropped ref in
// `keep` (work on it is still open, review-service.mjs heldRefs) stays, after
// the listed ones.
export function replaceComments(rev, incoming, keep = new Set()) {
  const next = clone(rev);
  const seen = new Set();
  const list = [];
  for (const c of incoming) {
    if (seen.has(c.ref)) throw new Error(`Duplicate comment ref ${c.ref}`);
    seen.add(c.ref);
    const old = findComment(rev, c.ref);
    list.push({
      ref: c.ref,
      reviewer: c.reviewer,
      text: c.text,
      requirements: Array.isArray(c.requirements)
        ? numberRequirements(old?.requirements ?? [], c.requirements, citedIds(old))
        : clone(old?.requirements ?? []),
      status: c.status ?? old?.status ?? "pending",
      changes: old ? clone(old.changes) : [],
      blocked: old ? clone(old.blocked) : [],
    });
  }
  for (const old of rev.comments) {
    if (seen.has(old.ref)) continue;
    if (old.changes.length > 0) {
      throw new Error(`Comment ${old.ref} has ${old.changes.length} recorded change(s); keep it`);
    }
    if (keep.has(old.ref)) list.push(clone(old));
  }
  next.comments = list;
  return next;
}

// Put a comment back so a change or blocked edit can be recorded under it
// when the comment was removed anyway. from: { reviewer, text, requirements }
// the daemon still holds, or null for the ref alone.
export function restoreComment(rev, ref, from = null) {
  const m = /^(?:R(\d+)|E)\.\d+$/.exec(ref);
  if (!m) throw new Error(`Bad comment ref ${ref}`);
  const next = clone(rev);
  next.comments.push(
    Comment.parse({
      ref,
      reviewer: from?.reviewer ?? (m[1] ? Number(m[1]) : 0),
      text: from?.text ?? "",
      requirements: clone(from?.requirements ?? []),
      status: "in_progress",
    }),
  );
  return next;
}

// Requirement ids that recorded changes of this comment cite, and those that
// blocked edits cite (blocked, applying or outcome_unknown, with the change
// an unanswered attempt would record): an approval or a late answer records
// them later, so their ids stay taken (Codex review 9).
function citedIds(comment) {
  const ids = (r) => (Array.isArray(r?.requirement_ids) ? r.requirement_ids : []);
  return [
    ...(comment?.changes ?? []).flatMap((ch) => ids(ch.rationale)),
    ...(comment?.blocked ?? []).flatMap((b) => [
      ...ids(b.args?.rationale),
      ...ids(b.attempt?.change?.rationale),
    ]),
  ].filter((id) => typeof id === "string");
}

// L5. texts: strings or { text }. A text that was already listed keeps its
// id, so changes that point at it stay linked; new texts get the next q-id.
// The next q-id is above every listed id and every id in `used` (the ids
// recorded changes cite), so a removed requirement's id is never given to a
// new one that no change answers.
export function numberRequirements(old, texts, used = []) {
  const byText = new Map(old.map((r) => [r.text, r.id]));
  let n = nextNumber([...old.map((r) => r.id), ...used], "q");
  const out = [];
  const seen = new Set();
  for (const t of texts) {
    const text = String(typeof t === "string" ? t : (t?.text ?? "")).trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    out.push({ id: byText.get(text) ?? `q${n++}`, text });
  }
  return out;
}

export function setRequirements(rev, ref, texts) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  c.requirements = numberRequirements(c.requirements, texts, citedIds(c));
  return next;
}

export function setCommentStatus(rev, ref, status) {
  const next = clone(rev);
  requireComment(next, ref).status = status;
  return next;
}

function nextNumber(ids, prefix) {
  let max = 0;
  for (const id of ids) {
    if (id.startsWith(prefix)) {
      const n = Number(id.slice(prefix.length));
      if (Number.isInteger(n) && n > max) max = n;
    }
  }
  return max + 1;
}

export function addChange(rev, ref, change, now = new Date()) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  const id = `${ref}-${nextNumber(
    c.changes.map((x) => x.id),
    `${ref}-`,
  )}`;
  const row = ChangeSchema.parse({ ...change, id, created_at: now.toISOString() });
  c.changes.push(row);
  if (c.status === "pending") c.status = "in_progress";
  return { revision: next, change: row };
}

// ids: Word comment ids the task pane's cleanup deleted. Marks the changes
// whose rationale comment is one of them; returns null when none is.
export function markCommentsCleaned(rev, ids, now = new Date()) {
  const gone = new Set((ids ?? []).map(String));
  const next = clone(rev);
  let hit = false;
  for (const c of next.comments) {
    for (const ch of c.changes) {
      if (ch.word_comment_id == null || ch.comment_cleaned_at) continue;
      if (!gone.has(String(ch.word_comment_id))) continue;
      ch.comment_cleaned_at = now.toISOString();
      hit = true;
    }
  }
  return hit ? next : null;
}

// rows: [{ ref, change_id, state, paragraph_index }] from classifyRevision.
export function applyStates(rev, rows, now = new Date()) {
  const next = clone(rev);
  const at = now.toISOString();
  for (const r of rows) {
    const ch = findComment(next, r.ref)?.changes.find((x) => x.id === r.change_id);
    if (!ch) continue;
    ch.state = r.state;
    if (r.paragraph_index != null) ch.paragraph_index = r.paragraph_index;
    ch.checked_at = at;
  }
  return next;
}

export function addBlocked(rev, ref, { args, blocked_by }, now = new Date()) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  const id = `${ref}-b${nextNumber(
    c.blocked.map((x) => x.id),
    `${ref}-b`,
  )}`;
  const row = { id, args: clone(args), blocked_by: [...blocked_by], created_at: now.toISOString() };
  c.blocked.push(row);
  return { revision: next, blocked: row };
}

export function markBlockedUnknown(rev, ref, blockedId, attempt) {
  const next = clone(rev);
  const b = requireComment(next, ref).blocked.find((x) => x.id === blockedId);
  if (!b) throw new Error(`No blocked edit ${blockedId} under ${ref}`);
  b.state = "outcome_unknown";
  b.attempt = Attempt.parse(clone(attempt));
  return next;
}

// Start an override: the item turns applying so an overlapping override of
// the same item is refused instead of editing Word twice. An unanswered one
// (outcome_unknown) goes again only with force.
export function claimBlocked(rev, ref, blockedId, { force = false } = {}) {
  const next = clone(rev);
  const b = findComment(next, ref)?.blocked.find((x) => x.id === blockedId);
  if (!b) throw new Error(`No blocked edit ${blockedId} under ${ref}`);
  if (b.state === "applying") {
    throw new Error(`${b.id}: this exception is already being applied. Wait for it to finish.`);
  }
  if (b.state === "outcome_unknown" && !force) {
    throw new Error(
      `${b.id}: Word did not answer the last attempt, so it may already be applied. Refresh status, then check the paragraph. Re-apply only if the change is not there.`,
    );
  }
  const before = clone(b);
  b.state = "applying";
  return { revision: next, blocked: before };
}

// End an override that did not apply: back to the state it had before.
export function releaseBlocked(rev, ref, blockedId, state) {
  const next = clone(rev);
  const b = findComment(next, ref)?.blocked.find((x) => x.id === blockedId);
  if (!b || b.state !== "applying") return null;
  b.state = state;
  return next;
}

// applying items no running override owns (the daemon stopped mid-call) may
// be in Word: they turn outcome_unknown. `active`: "<ref>\0<id>" keys.
export function markStaleApplying(rev, active) {
  const next = clone(rev);
  let n = 0;
  for (const c of next.comments) {
    for (const b of c.blocked) {
      if (b.state === "applying" && !active.has(`${c.ref}\0${b.id}`)) {
        b.state = "outcome_unknown";
        n++;
      }
    }
  }
  return n ? next : null;
}

export function takeBlocked(rev, ref, blockedId) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  const i = c.blocked.findIndex((b) => b.id === blockedId);
  if (i === -1) throw new Error(`No blocked edit ${blockedId} under ${ref}`);
  const [blocked] = c.blocked.splice(i, 1);
  return { revision: next, blocked };
}

export function addLedgerItem(rev, item) {
  const next = clone(rev);
  const id = `L${nextNumber(
    next.ledger.map((x) => x.id),
    "L",
  )}`;
  const row = LedgerItemSchema.parse({ ...item, id });
  next.ledger.push(row);
  return { revision: next, item: row };
}

export function updateLedgerItem(rev, id, patch) {
  const next = clone(rev);
  const i = next.ledger.findIndex((x) => x.id === id);
  if (i === -1) throw new Error(`No ledger item ${id}`);
  const old = next.ledger[i];
  // A new rule text or check invalidates the last check (L10 needs a full run).
  const ruleChanged =
    ("text" in patch && patch.text !== old.text) ||
    ("check" in patch && JSON.stringify(patch.check) !== JSON.stringify(old.check));
  const reset = ruleChanged
    ? {
        status: "unchecked",
        violation_count: 0,
        violations: [],
        unsure: [],
        checked_hashes: null,
        checked_at: null,
      }
    : {};
  next.ledger[i] = LedgerItemSchema.parse({ ...old, ...patch, ...reset, id });
  return next;
}

export function removeLedgerItem(rev, id) {
  const next = clone(rev);
  const before = next.ledger.length;
  next.ledger = next.ledger.filter((x) => x.id !== id);
  if (next.ledger.length === before) throw new Error(`No ledger item ${id}`);
  return next;
}

// ---- P15: the ledger as a JSON file -------------------------------------------
// A file holds rules only ({ kind, text, scope, check }): ids and check
// results belong to one manuscript. tests/fixtures/review_ledger.json has the
// same item shape, so a hand-written list imports too.

const RULE_KEYS = ["kind", "text", "scope", "check"];
const ruleOf = (item) => Object.fromEntries(RULE_KEYS.map((k) => [k, item?.[k]]));
const ruleKey = (item) => JSON.stringify(RULE_KEYS.map((k) => item[k]));

export function ledgerExportDoc(rev, now = new Date()) {
  return {
    format: "manuscript-pane-ledger",
    version: 1,
    exported_at: now.toISOString(),
    items: rev.ledger.map(ruleOf),
  };
}

// mode "append": rules the ledger already holds are skipped. "replace": the
// ledger is emptied first. Every row is checked before anything is added, so
// a bad row adds nothing. Ids come from addLedgerItem.
export function importLedgerItems(rev, items, { mode = "append" } = {}) {
  if (!Array.isArray(items)) throw new Error("ledger file needs an items array");
  if (mode !== "append" && mode !== "replace") throw new Error(`unknown import mode ${mode}`);
  const rows = items.map((it, n) => {
    const parsed = LedgerItemSchema.safeParse({ ...ruleOf(it), id: "L0" });
    if (!parsed.success) {
      const issue = parsed.error.issues[0];
      const where = issue?.path?.length ? `${issue.path.join(".")}: ` : "";
      throw new Error(`item ${n + 1}: ${where}${issue?.message ?? "invalid"}`);
    }
    return ruleOf(parsed.data);
  });
  let next = mode === "replace" ? { ...clone(rev), ledger: [] } : rev;
  const have = new Set(next.ledger.map(ruleKey));
  const added = [];
  const skipped = [];
  for (const row of rows) {
    if (have.has(ruleKey(row))) {
      skipped.push(row);
      continue;
    }
    const r = addLedgerItem(next, row);
    next = r.revision;
    have.add(ruleKey(row));
    added.push(r.item);
  }
  return { revision: next, added, skipped };
}

// ---- L4: snapshot and outside changes ----------------------------------------

export const EXTERNAL_KEEP = 100;

export function setSnapshot(rev, paragraphs, now = new Date()) {
  const next = clone(rev);
  next.snapshot = { taken_at: now.toISOString(), paragraphs: [...paragraphs] };
  return next;
}

// Append outside changes with X-ids; only the newest EXTERNAL_KEEP stay.
export function addExternal(rev, items, now = new Date()) {
  const next = clone(rev);
  let n = nextNumber(
    next.external.map((x) => x.id),
    "X",
  );
  const added = items.map((x) => ({ ...clone(x), id: `X${n++}`, detected_at: now.toISOString() }));
  next.external = [...next.external, ...added].slice(-EXTERNAL_KEEP);
  return { revision: next, added };
}

export function dismissExternal(rev, id) {
  const next = clone(rev);
  const before = next.external.length;
  next.external = next.external.filter((x) => x.id !== id);
  if (next.external.length === before) throw new Error(`No outside change ${id}`);
  return next;
}
