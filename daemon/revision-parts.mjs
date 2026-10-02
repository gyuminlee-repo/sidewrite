// Supplementary documents in the revision of a manuscript: everything that
// reads the review file of ANOTHER document. Each document keeps a review
// file of its own (revision-store.mjs); the file of a supplement names its
// manuscript (part_of) and the file of the manuscript lists its supplements
// (parts).
//
// Rules for the callers (review-service.mjs):
//   - a read of another file is a plain loadRevision, never a place in its
//     write chain (a save is an atomic rename, so a read sees a whole file);
//   - no update of one file runs inside the update of another;
//   - the only write to two files is join and leave: the supplement first,
//     then the manuscript. Either end alone is put right by the next load of
//     the supplement.

import { access, readdir, stat } from "node:fs/promises";
import { basename, dirname, join, relative, resolve, sep } from "node:path";
import { loadRevision, localDocPath, revisionPathFor } from "./revision-store.mjs";
import { partLabel } from "./response-quotes.mjs";

export { partLabel };

const SIDECAR = /\.revision\.json$/i;
const under = (dir, path) => path.startsWith(dir + sep);
// The file name of a path written on any system.
const nameOf = (p) =>
  String(p ?? "")
    .split(/[\\/]/)
    .pop();

async function exists(p) {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

// The path of another document as a review file stores it: relative to the
// folder `dir` of that review file ("/" separators) when the document is
// under it, absolute otherwise (the rule of the figure files).
export const storedPartPath = (dir, abs) =>
  under(dir, abs) ? relative(dir, abs).split(sep).join("/") : abs;
// And back.
export const resolvePart = (dir, stored) => resolve(dir, stored);

// The document a review file belongs to. The file name of the document is in
// the review file; its folder is where the review file is now (the folder
// written in the file may be that of another computer).
function docOfSidecar(path, rev) {
  const named = join(dirname(path), nameOf(rev?.manuscript));
  if (rev?.manuscript && revisionPathFor(named) === path) return named;
  return join(dirname(path), basename(path).replace(SIDECAR, ".docx"));
}

// The other review files in the folder of `doc`: [{ doc, revision }]. A file
// that cannot be read is no candidate.
async function siblings(doc) {
  const dir = dirname(doc);
  const own = revisionPathFor(doc);
  let names;
  try {
    names = await readdir(dir);
  } catch {
    return [];
  }
  const out = [];
  for (const name of names.filter((n) => SIDECAR.test(n)).sort()) {
    const path = join(dir, name);
    if (path === own) continue;
    try {
      const revision = await loadRevision(path, "");
      out.push({ doc: docOfSidecar(path, revision), revision });
    } catch {
      /* not a review file this daemon reads */
    }
  }
  return out;
}

// Manuscripts in the folder of `doc` whose revision it could join: review
// files that hold comments and are no supplement themselves.
// covers: every comment `own` (the review file of `doc`) holds is in the list
// of that manuscript, with the same text. It is how a supplement that left
// its revision with changes of its own is told from a manuscript with a
// review of its own: the pane offers the first a way back.
export async function linkCandidates(doc, own = null) {
  const mine = own?.comments ?? [];
  return (await siblings(doc))
    .filter((s) => s.revision.comments.length > 0 && !s.revision.part_of)
    .map((s) => {
      const listed = new Map(s.revision.comments.map((c) => [c.ref, c.text]));
      return {
        doc: s.doc,
        name: basename(s.doc),
        comments: s.revision.comments.length,
        updated_at: s.revision.updated_at,
        covers: mine.length > 0 && mine.every((c) => listed.get(c.ref) === c.text),
      };
    });
}

// The link a supplement lost (an earlier daemon saved its file without
// part_of): the manuscript in its folder that lists it, as a part_of value.
export async function claimedBy(doc) {
  for (const s of await siblings(doc)) {
    const part = (s.revision.parts ?? []).find((p) => resolvePart(dirname(s.doc), p.doc) === doc);
    if (part) {
      return {
        manuscript: storedPartPath(dirname(doc), s.doc),
        label: part.label,
        joined_at: part.joined_at,
      };
    }
  }
  return null;
}

// What the author typed or picked as the manuscript to join: the path of the
// document or of its review file, absolute or relative to the folder of
// `doc`, in the form of this system or of Windows.
export async function mainDocFrom(input, doc, env = process.env) {
  const text = String(input ?? "")
    .trim()
    .replace(/^"(.*)"$/, "$1");
  const local = text ? localDocPath(text, env) : null;
  if (!local) throw new Error("Type the path of the manuscript, or pick it from the list.");
  const abs = resolve(dirname(doc), local);
  if (!SIDECAR.test(abs)) return abs;
  return docOfSidecar(abs, (await exists(abs)) ? await loadRevision(abs, "") : null);
}

// The manuscript of a supplement: { doc, path, name, revision }. revision is
// null when the manuscript has no review file where the link points.
export async function loadMain(doc, partOf) {
  const mainDoc = resolvePart(dirname(doc), partOf.manuscript);
  const path = revisionPathFor(mainDoc);
  return {
    doc: mainDoc,
    path,
    name: basename(mainDoc),
    revision: (await exists(path)) ? await loadRevision(path, mainDoc) : null,
  };
}

// The supplements of a manuscript: [{ doc, name, label, revision }], revision
// null when that review file is not there or cannot be read.
export async function loadParts(mainDoc, mainRev) {
  const out = [];
  for (const p of mainRev.parts ?? []) {
    const doc = resolvePart(dirname(mainDoc), p.doc);
    const path = revisionPathFor(doc);
    let revision = null;
    try {
      if (await exists(path)) revision = await loadRevision(path, doc);
    } catch {
      revision = null;
    }
    out.push({ doc, name: basename(doc), label: p.label, joined_at: p.joined_at, revision });
  }
  return out;
}

// The file of a document as it is on the disk now, in the form a status
// refresh records it (review-service.mjs review_refresh doc_state.file), or
// null when it cannot be read.
export async function fileOnDisk(doc, statFile = stat) {
  try {
    const s = await statFile(doc);
    return { mtime: s.mtime.toISOString(), bytes: s.size };
  } catch {
    return null;
  }
}

// Can what the review file of a supplement records still be trusted? The
// manuscript reads a supplement from its review file only, never from Word,
// so the record must be of the file as it is on the disk now:
//   - the supplement was refreshed after it joined (doc_state, not older
//     than the link),
//   - on a document without unsaved changes (doc_state.saved === true; not
//     known counts as unsaved),
//   - and the file still has the time and the size recorded then.
// Returns null when it can, else why not: "no_record" (no review file, or
// never refreshed since the join), "unsaved", "changed" (the file is not the
// one that was read, or cannot be read now).
export function partUnconfirmed({ revision, joined_at }, onDisk) {
  const ds = revision?.doc_state ?? null;
  if (!ds || (joined_at && ds.at < joined_at)) return "no_record";
  if (ds.saved !== true) return "unsaved";
  if (!ds.file || !onDisk || ds.file.mtime !== onDisk.mtime || ds.file.bytes !== onDisk.bytes)
    return "changed";
  return null;
}

// The supplements of a manuscript as the response letter, its checks and the
// submission checklist read them: loadParts, each with
//   label       what the letter calls the file (partLabel)
//   paragraphs  the text of the file at its last status refresh (the
//               accepted view, as its review file stores it)
//   confirmed   false: the record cannot be trusted (why: partUnconfirmed)
// Plain reads of the review files and one stat per file; nothing of Word.
export async function partsForLetter(mainDoc, mainRev, { statFile = stat } = {}) {
  const out = [];
  for (const p of await loadParts(mainDoc, mainRev)) {
    const why = partUnconfirmed(p, await fileOnDisk(p.doc, statFile));
    out.push({
      ...p,
      label: partLabel(p.label, p.name),
      paragraphs: p.revision?.snapshot?.paragraphs ?? [],
      confirmed: why === null,
      why,
    });
  }
  return out;
}

// Comment refs that hold recorded changes in a supplement, each with the
// name of that file: ref -> name.
export function partChangeRefs(parts) {
  const refs = new Map();
  for (const p of parts) {
    for (const c of p.revision?.comments ?? []) {
      if (c.changes.length > 0 && !refs.has(c.ref)) refs.set(c.ref, p.name);
    }
  }
  return refs;
}

// The key of the session hint of a document (sessions.mjs): its path when it
// is a file of a revision with parts (a supplement, or a manuscript that has
// one), null for any other document, which keeps the hint of its folder.
// A document without a review file, or with one that cannot be read, is
// such another document.
export async function sessionDocKey(doc) {
  if (!doc) return null;
  try {
    const path = revisionPathFor(doc);
    if (!(await exists(path))) return null;
    const rev = await loadRevision(path, doc);
    return rev.part_of || rev.parts?.length ? resolve(doc) : null;
  } catch {
    return null;
  }
}

// For the pane of a supplement: which of its changes the letter of the
// manuscript cites, as the review file of the manuscript has them
// (response-quotes.mjs partCitations wrote them at its last refresh):
//   [{ ref, kind, section, quote, change_ids }]
// name: the file name of the supplement. A citation with a problem is not
// printed by the letter and is left out.
export function citedInMain(mainRev, name) {
  return (mainRev?.comments ?? []).flatMap((c) =>
    (c.response?.citations ?? [])
      .filter((cit) => cit.doc === name && !cit.problem)
      .map((cit) => ({
        ref: c.ref,
        kind: cit.kind,
        section: cit.section,
        quote: cit.quote,
        change_ids: cit.change_ids,
      })),
  );
}

// The link of a document as a reply tells it to the pane (never saved):
//   part_of: { doc, name, label } of its manuscript, or null
//   parts: [{ doc, name, label }] of its supplements
export function linkedSummary(doc, rev) {
  const dir = dirname(doc);
  const main = rev.part_of ? resolvePart(dir, rev.part_of.manuscript) : null;
  // label: what the response letter calls the supplement (never empty).
  return {
    part_of: main
      ? { doc: main, name: basename(main), label: partLabel(rev.part_of.label, basename(doc)) }
      : null,
    parts: (rev.parts ?? []).map((p) => {
      const part = resolvePart(dir, p.doc);
      return { doc: part, name: basename(part), label: partLabel(p.label, basename(part)) };
    }),
  };
}

// The other files of the revision a document is in: [{ doc, name, revision }].
// For a supplement: its manuscript, then the other supplements of that
// manuscript. For a manuscript: its supplements. For any other document: [].
// revision is null for a file whose review file is not there. Plain reads.
export async function otherFiles(doc, rev) {
  const own = resolve(doc);
  if (rev?.part_of) {
    const main = await loadMain(own, rev.part_of);
    const rest = main.revision ? await loadParts(main.doc, main.revision) : [];
    return [
      { doc: main.doc, name: main.name, revision: main.revision },
      ...rest
        .filter((p) => p.doc !== own)
        .map(({ doc, name, revision }) => ({ doc, name, revision })),
    ];
  }
  if (!rev?.parts?.length) return [];
  return (await loadParts(own, rev)).map(({ doc, name, revision }) => ({ doc, name, revision }));
}

// The changes on record in those files, as the pane of this document shows
// where a request stands elsewhere (never saved):
//   [{ file, ref, id, state, requirement_ids }]
// state is the one the last status refresh of THAT file wrote into its
// review file (a change no refresh has seen yet is pending).
export function linkedChanges(files) {
  return files.flatMap((f) =>
    (f.revision?.comments ?? []).flatMap((c) =>
      c.changes.map((ch) => ({
        file: f.name,
        ref: c.ref,
        id: ch.id,
        state: ch.state,
        requirement_ids: ch.rationale?.requirement_ids ?? [],
      })),
    ),
  );
}

// What the author answered to questions asked in those files:
//   [{ file, ref, id, question, answer }]
export function linkedAnswers(files) {
  return files.flatMap((f) =>
    (f.revision?.comments ?? []).flatMap((c) =>
      (c.questions ?? [])
        .filter((q) => q.answer)
        .map((q) => ({
          file: f.name,
          ref: c.ref,
          id: q.id,
          question: q.question,
          answer: q.answer,
        })),
    ),
  );
}

// The requests of one comment that the turn of the manuscript noted as
// belonging to another file: [{ id, text, why }], in the order of the list.
export function outsideNotes(mainRev, ref) {
  const c = (mainRev?.comments ?? []).find((x) => x.ref === ref);
  const why = new Map((c?.requirements_outside ?? []).map((n) => [n.id, n.why]));
  return (c?.requirements ?? [])
    .filter((r) => why.has(r.id))
    .map((r) => ({ id: r.id, text: r.text, why: why.get(r.id) ?? "" }));
}

// What belongs to the manuscript is changed in the manuscript: a request
// that would change it in the file of a supplement is refused.
export function refuseInPart(rev) {
  if (rev?.part_of) throw new Error(`Edit this in ${nameOf(rev.part_of.manuscript)}.`);
}
