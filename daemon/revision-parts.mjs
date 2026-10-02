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

import { access, readdir } from "node:fs/promises";
import { basename, dirname, join, relative, resolve, sep } from "node:path";
import { loadRevision, localDocPath, revisionPathFor } from "./revision-store.mjs";

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
    out.push({ doc, name: basename(doc), label: p.label, revision });
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

// The link of a document as a reply tells it to the pane (never saved):
//   part_of: { doc, name, label } of its manuscript, or null
//   parts: [{ doc, name, label }] of its supplements
export function linkedSummary(doc, rev) {
  const dir = dirname(doc);
  const main = rev.part_of ? resolvePart(dir, rev.part_of.manuscript) : null;
  return {
    part_of: main ? { doc: main, name: basename(main), label: rev.part_of.label } : null,
    parts: (rev.parts ?? []).map((p) => {
      const part = resolvePart(dir, p.doc);
      return { doc: part, name: basename(part), label: p.label };
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
