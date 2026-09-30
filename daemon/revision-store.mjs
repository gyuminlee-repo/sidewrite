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
  created_at: z.string(),
  state: z.enum(["pending", "accepted", "rejected", "user_edited", "unlocated"]).default("pending"),
  checked_at: z.string().nullable().default(null),
});

const Blocked = z.object({
  id: z.string(),
  args: z.record(z.string(), z.unknown()),
  blocked_by: z.array(z.string()).min(1),
  created_at: z.string(),
});

// L5: one thing the reviewer asks for, e.g. { id: "q1", text: "state the unit" }.
const Requirement = z.object({
  id: z.string().regex(/^q\d+$/),
  text: z.string().min(1),
});

const Comment = z.object({
  ref: z.string().regex(/^R\d+\.\d+$/),
  reviewer: z.number().int().positive(),
  text: z.string(),
  requirements: z.array(Requirement).default([]),
  status: z.enum(["pending", "in_progress", "done"]).default("pending"),
  changes: z.array(ChangeSchema).default([]),
  blocked: z.array(Blocked).default([]),
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
  status: z.enum(["unchecked", "ok", "violations", "error"]).default("unchecked"),
  violation_count: z.number().int().nonnegative().default(0),
  checked_at: z.string().nullable().default(null),
  // L10: violations of the last check keyed by paragraph hash, and the
  // hashes of every paragraph that check covered (null: never checked).
  violations: z.array(Violation).default([]),
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
// has changes is refused: its edits are in the document.
export function replaceComments(rev, incoming) {
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
    if (!seen.has(old.ref) && old.changes.length > 0) {
      throw new Error(`Comment ${old.ref} has ${old.changes.length} recorded change(s); keep it`);
    }
  }
  next.comments = list;
  return next;
}

// Requirement ids that recorded changes of this comment cite.
function citedIds(comment) {
  return (comment?.changes ?? []).flatMap((ch) => ch.rationale?.requirement_ids ?? []);
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
