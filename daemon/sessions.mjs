// Per-(host, workspace) Agent SDK session bookkeeping.
//
// The SDK persists each session's transcript to
// ~/.claude/projects/<hash>/<session_id>.jsonl. To resume via
// query({ options: { resume } }) we need that session_id. Word and Excel
// are independent surfaces — opening the same folder in Word vs Excel is
// a different conversation — so the session_id is keyed by **host AND
// cwd**, not cwd alone. (A single cwd-keyed session is why Excel used to
// replay Word's history.)
//
// The folder list itself (for the workspace switcher UI) stays
// cwd-deduped and host-agnostic — you pick a folder, not a folder+host.
//
// Two documents in one folder share that hint. For the files of one revision
// (a manuscript and its supplementary documents, revision-parts.mjs) that is
// wrong: the pane of the supplement would resume the session of the
// manuscript, the text of the manuscript would be in its turn, and two turns
// could write to one transcript. Such a document has a hint of its own, by
// its path (`docs`). It never reads the hint of the folder: that one may be
// the session of the other file. A document that is in no such revision
// keeps the hint of the folder, so every hint stored before is still read.
//
// Stored at ~/.claude/office-addins/sessions.json (version 2):
//
//   {
//     "version": 2,
//     "folders": {
//       "/Users/x/Proj": {
//         "last_used": "2026-05-16T…Z",
//         "display_name": "Proj",
//         "sessions": { "word": "uuid…", "excel": "uuid…" },
//         "docs": { "/Users/x/Proj/SI.docx": { "word": "uuid…" } }
//       }
//     }
//   }
//
// `docs` is optional: a file without it is read and written as before.

import { readFile, writeFile, mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { join, dirname, basename } from "node:path";
import { isSystemHomeChild } from "./system-paths.mjs";

const FILE = join(homedir(), ".claude", "office-addins", "sessions.json");
const VERSION = 2;

// Don't persist or surface OS-managed $HOME children (e.g. ~/Library) as
// a recent workspace.
function isAllowedMatterPath(cwd) {
  return !isSystemHomeChild(cwd);
}

function normalizeHost(host) {
  return host === "word" || host === "excel" ? host : null;
}

async function readState() {
  let parsed;
  try {
    parsed = JSON.parse(await readFile(FILE, "utf8"));
  } catch (e) {
    if (e.code === "ENOENT") return { version: VERSION, folders: {} };
    throw e;
  }
  if (!parsed || typeof parsed.folders !== "object") {
    return { version: VERSION, folders: {} };
  }
  if (parsed.version === VERSION) return parsed;
  // Migrate v1 ({ session_id } per folder): the old id can't be attributed
  // to a host, so drop it (one conversation lost across the upgrade), but
  // keep the folder so it still shows in recents.
  const folders = {};
  for (const [cwd, info] of Object.entries(parsed.folders ?? {})) {
    folders[cwd] = {
      last_used: info?.last_used ?? new Date(0).toISOString(),
      display_name: info?.display_name ?? basename(cwd),
      sessions: {},
    };
  }
  return { version: VERSION, folders };
}

async function writeState(state) {
  await mkdir(dirname(FILE), { recursive: true });
  await writeFile(FILE, JSON.stringify(state, null, 2) + "\n", "utf8");
}

function ensureFolder(state, cwd) {
  let f = state.folders[cwd];
  if (!f) {
    f = { last_used: new Date(0).toISOString(), display_name: basename(cwd), sessions: {} };
    state.folders[cwd] = f;
  }
  if (typeof f.sessions !== "object" || f.sessions === null) f.sessions = {};
  return f;
}

// The resumable session_id for this host in this folder, or null.
// doc: the path of a document that has a hint of its own (a file of a
// revision with parts); only that hint is read then, never the folder's.
export async function getSessionId(host, cwd, doc = null) {
  const h = normalizeHost(host);
  if (!h) return null;
  const state = await readState();
  const f = state.folders[cwd];
  if (doc) return f?.docs?.[doc]?.[h] ?? null;
  return f?.sessions?.[h] ?? null;
}

// Record the session_id for (host, cwd), or for (host, doc) when the document
// has a hint of its own, and bump last_used.
export async function saveSessionId(host, cwd, sessionId, doc = null) {
  const h = normalizeHost(host);
  if (!h) return;
  const state = await readState();
  const f = ensureFolder(state, cwd);
  if (doc) {
    if (typeof f.docs !== "object" || f.docs === null) f.docs = {};
    f.docs[doc] = { ...f.docs[doc], [h]: sessionId };
  } else f.sessions[h] = sessionId;
  f.last_used = new Date().toISOString();
  f.display_name = basename(cwd);
  await writeState(state);
}

// One line of the daemon log per turn: the document of the pane and the SDK
// session the turn runs on. The CLI sends an init at the start of every turn.
//   started    a session nobody asked to resume (or not the one asked for)
//   resumed    the session the stored hint named, at the first turn of a loop
//   continued  a later turn of the same loop
// asked is printed as well: should the SDK give a resumed session a new id,
// the line still says whose transcript the turn continued.
// doc: the file name of the document, or null. asked: the id the loop was
// started with (null: none). known: the id the loop already had from an
// earlier init (null: this is its first).
export function turnSessionLine({ doc = null, sessionId, asked = null, known = null }) {
  const how =
    known && known === sessionId
      ? "continued"
      : asked && asked === sessionId
        ? "resumed"
        : "started";
  return `[agent] turn doc=${doc || "-"} session=${sessionId} asked=${asked || "-"} (${how})`;
}

// Touch a folder's bookkeeping (last_used / display_name) without changing
// any session id. Ensures the folder entry exists for the per-host session
// store. Skips OS-managed $HOME children so they never get persisted.
export async function touchFolder(cwd) {
  if (!isAllowedMatterPath(cwd)) return;
  const state = await readState();
  const f = ensureFolder(state, cwd);
  f.last_used = new Date().toISOString();
  f.display_name = basename(cwd);
  await writeState(state);
}
