// Figure lane: a reviewer request that is answered by changing a figure.
//
// The figure files are made by hand in any tool. This service records which
// file took the place of which under a reviewer comment, renders both for a
// comparison, and on Adopt copies the new file into the submission folder.
// It calls no model and no figure tool.
//
// Where it writes, and nowhere else (all under the manuscript folder):
//   <figures>/.renders/      pictures and facts of the files, by content hash
//   <figures>/submission/    the adopted copies and FILE_SOURCES.md
//   <manuscript>.revision.json
// <figures> is "figures" next to the manuscript unless revision.json names
// another folder (figure_lane.dir). The files the author names (old, new,
// their views) are only read; their hashes are taken when a change is
// proposed and compared again before a render and before Adopt.
//
// Every function takes the local path of the manuscript (`doc`). Which
// document a caller may name is the caller's question (figure-http.mjs).

import { createHash, randomUUID } from "node:crypto";
import { createReadStream, constants as FS } from "node:fs";
import {
  appendFile,
  copyFile,
  mkdir,
  readFile,
  readdir,
  realpath,
  rename,
  stat,
  unlink,
  writeFile,
} from "node:fs/promises";
import { basename, dirname, extname, isAbsolute, join, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import {
  revisionPathFor,
  loadRevision,
  withRevision,
  findComment,
  addFigureChange,
  updateFigureChange,
  setFigureState,
} from "./revision-store.mjs";
import { toLocalPath } from "./workspace.mjs";
import { runPy as defaultRunPy } from "./py-tool.mjs";
import { outputName, MAX_OUTPUT_NAMES } from "./output-name.mjs";

export const FIGURE_DIR = "figures";
export const RENDER_DPI = 150;
// What tools/figure_render.py opens.
export const FIGURE_TYPES = [".pdf", ".ai", ".svg", ".png", ".jpg", ".jpeg", ".tif", ".tiff"];
export const SOURCES_FILE = "FILE_SOURCES.md";

const RENDER_TOOL = join(
  dirname(fileURLToPath(import.meta.url)),
  "..",
  "tools",
  "figure_render.py",
);

// An error the page shows as it is. status: the HTTP status of the reply.
export class FigureError extends Error {
  constructor(message, status = 400) {
    super(message);
    this.status = status;
  }
}
const refuse = (message, status) => {
  throw new FigureError(message, status);
};

// What the store refuses (a state that cannot follow, a request the comment
// does not list) is said to the page as it is.
function stored(fn) {
  try {
    return fn();
  } catch (e) {
    throw e instanceof FigureError ? e : new FigureError(e.message, 409);
  }
}

export function sha256File(path) {
  return new Promise((done, fail) => {
    const h = createHash("sha256");
    createReadStream(path)
      .on("error", fail)
      .on("data", (chunk) => h.update(chunk))
      .on("end", () => done(h.digest("hex")));
  });
}

const under = (dir, path) => path.startsWith(dir + sep);
const posix = (p) => p.split(sep).join("/");

// The stored path of a file: relative to the manuscript folder when under it.
const storedPath = (docDir, abs) => (under(docDir, abs) ? posix(relative(docDir, abs)) : abs);
// And back. A relative path that leaves the manuscript folder is refused.
function absPath(docDir, stored) {
  if (isAbsolute(stored)) return stored;
  const abs = resolve(docDir, stored);
  if (!under(docDir, abs)) refuse(`${stored} is outside the manuscript folder.`);
  return abs;
}

async function fileRef(docDir, abs) {
  const st = await stat(abs);
  return {
    path: storedPath(docDir, abs),
    sha256: await sha256File(abs),
    bytes: st.size,
    mtime: st.mtime.toISOString(),
  };
}

// What a figure change says about its two files, in plain words. Pure.
// Reads the facts and renders stored with the change.
export function figureFindings(f) {
  const out = [];
  const side = (s) => (s ? (s.view ?? s.source) : null);
  const b = side(f.before);
  const a = side(f.after);
  if (b && a && b.sha256 === a.sha256) out.push("The old and the new file are identical.");
  const rb = f.renders?.before;
  const ra = f.renders?.after;
  if (rb && ra && (rb.width !== ra.width || rb.height !== ra.height))
    out.push(
      `The two pictures have different pixel sizes: ${rb.width} x ${rb.height} and ${ra.width} x ${ra.height}. The overlay and the difference view put their top left corners together.`,
    );
  const fb = f.facts?.before;
  const fa = f.facts?.after;
  if (fb && fa && fb.type !== fa.type)
    out.push(`The file type changed from ${fb.type} to ${fa.type}.`);
  if (fa && fa.pages > 1) out.push(`The new file has ${fa.pages} pages. Only the first is shown.`);
  if (fb && fa && fb.words > 0 && fa.words === 0)
    out.push("The old file has live text and the new file has none.");
  const loose = (fa?.fonts ?? []).filter((x) => !x.embedded).map((x) => x.name);
  if (loose.length) out.push(`Fonts not embedded in the new file: ${loose.join(", ")}.`);
  return out;
}

const cell = (s) =>
  String(s ?? "")
    .replace(/\|/g, "\\|")
    .replace(/\s+/g, " ");
const SOURCES_HEAD = [
  "# Figure files for submission",
  "",
  "One row per decision on the Sidewrite figures page. A file in this folder is a copy; the source file stays where it was.",
  "",
  "| Date | Figure | Comment | Decision | Source | Source SHA-256 | File in this folder | SHA-256 |",
  "| --- | --- | --- | --- | --- | --- | --- | --- |",
  "",
].join("\n");

export function createFigureService({
  runPy = defaultRunPy,
  env = process.env,
  log = console.log,
  now = () => new Date(),
} = {}) {
  const docDirOf = (doc) => dirname(doc);
  const load = (doc) => loadRevision(revisionPathFor(doc), doc);
  const update = (doc, fn) => withRevision(revisionPathFor(doc), doc, fn);

  // The figures folder and the two folders this service writes to.
  function dirs(doc, rev) {
    const docDir = docDirOf(doc);
    const figures = resolve(docDir, rev.figure_lane?.dir ?? FIGURE_DIR);
    if (!under(docDir, figures)) refuse("The figures folder must be inside the manuscript folder.");
    return {
      docDir,
      figures,
      renders: join(figures, ".renders"),
      submission: join(figures, "submission"),
    };
  }

  // A folder this service writes to, made when it is not there. A link that
  // leads out of the manuscript folder is refused.
  async function writable(docDir, dir) {
    await mkdir(dir, { recursive: true });
    const [root, real] = await Promise.all([realpath(docDir), realpath(dir)]);
    if (!under(root, real)) refuse(`${dir} leads outside the manuscript folder.`);
    return dir;
  }

  // A file the author named: a path in the list of the page, or a pasted
  // one (a Windows path when Word runs on Windows and the daemon in WSL).
  async function namedFile(d, input, what) {
    const raw = String(input ?? "").trim();
    if (!raw) refuse(`Choose ${what}.`);
    const local = toLocalPath(raw, env);
    if (local === null || /^[a-z][a-z0-9+.-]*:\/\//i.test(local))
      refuse(`${raw} is not a file this computer can open.`);
    const abs = isAbsolute(local) ? resolve(local) : resolve(d.docDir, local);
    if (abs === d.renders || under(d.renders, abs))
      refuse("A picture made for the comparison cannot be a figure file.");
    if (!FIGURE_TYPES.includes(extname(abs).toLowerCase()))
      refuse(`${basename(abs)} is not a figure file this page reads (${FIGURE_TYPES.join(", ")}).`);
    const st = await stat(abs).catch(() => null);
    if (!st?.isFile()) refuse(`No file at ${raw}.`);
    return abs;
  }

  async function sideOf(d, source, view, what) {
    const src = await namedFile(d, source, what);
    const v = String(view ?? "").trim() ? await namedFile(d, view, `the view of ${what}`) : null;
    return { source: await fileRef(d.docDir, src), view: v ? await fileRef(d.docDir, v) : null };
  }

  function figureOf(rev, ref, id) {
    const f = findComment(rev, ref)?.figure_changes?.find((x) => x.id === id);
    if (!f) refuse(`No figure change ${id} under ${ref}.`, 404);
    return f;
  }

  // The file of a side that is shown and submitted, checked against its
  // record. Throws when it is gone or was changed since.
  async function recorded(d, ref, label) {
    const abs = absPath(d.docDir, ref.path);
    const sha = await sha256File(abs).catch(() => null);
    if (sha === null) refuse(`${label} is not there any more: ${ref.path}`, 409);
    if (sha !== ref.sha256)
      refuse(
        `${label} changed since it was proposed: ${ref.path}. Save the new version under a new name and propose it again.`,
        409,
      );
    return abs;
  }

  // ---- renders ---------------------------------------------------------------

  let rendererOk = false;
  // { available, reason }. A yes is remembered; a no is asked again, so
  // installing PyMuPDF needs no restart.
  async function renderer() {
    if (rendererOk) return { available: true, reason: null };
    const r = await runPy(RENDER_TOOL, ["--probe"]);
    rendererOk = r.ok && r.json?.ok === true;
    return rendererOk
      ? { available: true, reason: null }
      : {
          available: false,
          reason:
            r.json?.code === "no_pymupdf"
              ? "PyMuPDF is not installed (pip install pymupdf). Figure changes are still recorded, adopted and cited, without the pictures."
              : `The figure renderer could not run: ${r.json?.error ?? r.error ?? "unknown"}`,
        };
  }

  // The picture and the facts of one side, from the cache or made now.
  // { render, facts, png } or { error }.
  async function renderSide(d, f, side) {
    const s = f[side];
    if (!s) return { error: `This figure change has no ${side} file.` };
    const ref = s.view ?? s.source;
    const src = await recorded(d, ref, side === "before" ? "The old file" : "The new file");
    const stem = `${ref.sha256}-${RENDER_DPI}`;
    const png = join(d.renders, `${stem}.png`);
    const factsPath = join(d.renders, `${stem}.json`);
    let facts = null;
    try {
      facts = JSON.parse(await readFile(factsPath, "utf8"));
      await stat(png);
    } catch {
      facts = null;
    }
    if (!facts) {
      await writable(d.docDir, d.renders);
      const tmp = join(d.renders, `.${stem}.${process.pid}.${randomUUID()}.tmp.png`);
      const r = await runPy(RENDER_TOOL, ["--in", src, "--out", tmp, "--dpi", String(RENDER_DPI)]);
      if (!r.json?.ok) {
        await unlink(tmp).catch(() => {});
        if (r.json?.code === "no_pymupdf") return { error: (await renderer()).reason };
        return { error: `${basename(src)} could not be rendered: ${r.json?.error ?? r.error}` };
      }
      rendererOk = true;
      facts = r.json;
      await rename(tmp, png);
      const tmpJson = `${tmp}.json`;
      await writeFile(tmpJson, JSON.stringify(facts) + "\n", "utf8");
      await rename(tmpJson, factsPath);
    }
    return {
      png,
      render: {
        path: storedPath(d.docDir, png),
        of: ref.sha256,
        dpi: RENDER_DPI,
        width: facts.width_px,
        height: facts.height_px,
      },
      facts: {
        type: facts.type,
        pages: facts.pages,
        width_pt: facts.width_pt ?? null,
        height_pt: facts.height_pt ?? null,
        words: facts.words ?? 0,
        images: facts.images ?? 0,
        paths: facts.paths ?? 0,
        fonts: facts.fonts ?? [],
      },
    };
  }

  // Renders the sides of a figure change and stores what was read. A side
  // that cannot be rendered is left as it is; the reasons are returned.
  async function renderAll(doc, d, ref, id) {
    const f = figureOf(await load(doc), ref, id);
    const renders = { ...f.renders };
    const facts = { ...f.facts };
    const problems = [];
    const pngs = {};
    for (const side of ["before", "after"]) {
      if (!f[side]) continue;
      let r;
      try {
        r = await renderSide(d, f, side);
      } catch (e) {
        r = { error: e.message };
      }
      if (r.error) {
        problems.push(r.error);
        continue;
      }
      renders[side] = r.render;
      facts[side] = r.facts;
      pngs[side] = r.png;
    }
    if (
      JSON.stringify(renders) !== JSON.stringify(f.renders) ||
      JSON.stringify(facts) !== JSON.stringify(f.facts)
    )
      await update(doc, (rev) => updateFigureChange(rev, ref, id, { renders, facts }));
    return { pngs, problems: [...new Set(problems)] };
  }

  // ---- the submission folder ---------------------------------------------------

  async function sourcesRow(d, row) {
    const path = join(d.submission, SOURCES_FILE);
    await writeFile(path, SOURCES_HEAD, { encoding: "utf8", flag: "wx" }).catch((e) => {
      if (e.code !== "EEXIST") throw e;
    });
    await appendFile(path, `| ${row.map(cell).join(" | ")} |\n`, "utf8");
  }

  // Copies `src` (sha256 `sha`) into the submission folder under its own
  // name, or under the next free "-2", "-3" name when that one is taken by
  // another file. A file of that name with the same content is the copy
  // already. Returns the path of the copy.
  async function submit(d, src, sha) {
    await writable(d.docDir, d.submission);
    const wanted = join(d.submission, basename(src));
    for (let n = 1; n <= MAX_OUTPUT_NAMES; n++) {
      const path = outputName(wanted, n);
      try {
        await copyFile(src, path, FS.COPYFILE_EXCL);
      } catch (e) {
        if (e.code !== "EEXIST") throw e;
        if ((await sha256File(path).catch(() => null)) === sha) return path;
        continue;
      }
      if ((await sha256File(path)) !== sha) {
        await unlink(path).catch(() => {});
        refuse(`${basename(src)} changed while it was copied. Propose it again.`, 409);
      }
      return path;
    }
    throw new Error(`No free name for ${wanted}.`);
  }

  // ---- what the page shows -----------------------------------------------------

  // A figure change with what is computed from it and from the disk.
  async function shown(d, f) {
    const problems = [];
    if (f.state === "adopted") {
      const there = await stat(absPath(d.docDir, f.adopted.path)).catch(() => null);
      if (!there) problems.push(`The adopted copy is missing: ${f.adopted.path}`);
    } else {
      // Size and time are enough for a look; Adopt and a render read the hash.
      for (const [label, s] of [
        ["The old file", f.before],
        ["The new file", f.after],
      ]) {
        const ref = s ? (s.view ?? s.source) : null;
        if (!ref) continue;
        const st = await stat(absPath(d.docDir, ref.path)).catch(() => null);
        if (!st) problems.push(`${label} is not there any more: ${ref.path}`);
        else if (st.size !== ref.bytes || st.mtime.toISOString() !== ref.mtime)
          problems.push(`${label} was saved again since it was proposed: ${ref.path}`);
      }
    }
    return { ...f, findings: figureFindings(f), problems };
  }

  const api = {
    renderer,

    // The comments of the manuscript with their requests and figure changes.
    async state(doc) {
      const rev = await load(doc);
      const d = dirs(doc, rev);
      const comments = [];
      for (const c of rev.comments) {
        const outside = new Map((c.requirements_outside ?? []).map((n) => [n.id, n.why]));
        const figures = [];
        for (const f of c.figure_changes ?? []) figures.push(await shown(d, f));
        comments.push({
          ref: c.ref,
          reviewer: c.reviewer,
          section: c.section ?? null,
          text: c.text,
          requirements: c.requirements.map((q) => ({
            id: q.id,
            text: q.text,
            outside: outside.has(q.id),
            why: outside.get(q.id) ?? "",
          })),
          figures,
        });
      }
      return {
        name: basename(doc),
        folder: storedPath(d.docDir, d.figures),
        renderer: await renderer(),
        comments,
      };
    },

    // Figure files in the figures folder and one level below it, newest
    // first. The pictures of the comparison are left out.
    async files(doc) {
      const d = dirs(doc, await load(doc));
      const out = [];
      const read = async (dir, deep) => {
        for (const e of await readdir(dir, { withFileTypes: true }).catch(() => [])) {
          if (e.name.startsWith(".")) continue;
          const path = join(dir, e.name);
          if (e.isDirectory()) {
            if (deep) await read(path, false);
          } else if (e.isFile() && FIGURE_TYPES.includes(extname(e.name).toLowerCase())) {
            const st = await stat(path);
            out.push({
              path: storedPath(d.docDir, path),
              bytes: st.size,
              mtime: st.mtime.toISOString(),
              submitted: under(d.submission, path),
            });
          }
        }
      };
      await read(d.figures, true);
      out.sort((a, b) => (a.mtime < b.mtime ? 1 : a.mtime > b.mtime ? -1 : 0));
      return { folder: storedPath(d.docDir, d.figures), files: out };
    },

    // input: { ref, figure, kind, requirement_ids, summary, before, after,
    // before_view, after_view }; the four files are paths.
    async propose(doc, input = {}) {
      const rev = await load(doc);
      const d = dirs(doc, rev);
      const ref = String(input.ref ?? "");
      if (!findComment(rev, ref)) refuse(`No reviewer comment ${ref}.`, 404);
      const kind = input.kind === "new" ? "new" : "revised";
      const figure = String(input.figure ?? "").trim();
      if (!figure) refuse('Enter the figure label, for example "Figure 2".');
      if (kind === "new" && String(input.before ?? "").trim())
        refuse("A new figure has no old file.");
      const before =
        kind === "revised"
          ? await sideOf(d, input.before, input.before_view, "the old file")
          : null;
      const after = await sideOf(d, input.after, input.after_view, "the new file");
      if (before && before.source.path === after.source.path)
        refuse(
          "The old and the new file are the same file. Save the new figure under a new name; the old file stays.",
        );
      let id;
      await update(doc, (cur) => {
        const r = stored(() =>
          addFigureChange(
            cur,
            ref,
            {
              figure,
              kind,
              requirement_ids: (Array.isArray(input.requirement_ids)
                ? input.requirement_ids
                : []
              ).map(String),
              summary: String(input.summary ?? "").trim(),
              before,
              after,
              made_by: "author",
            },
            now(),
          ),
        );
        id = r.figure.id;
        return r.revision;
      });
      const { problems } = await renderAll(doc, d, ref, id);
      log(`[figures] propose ${id} kind=${kind} rendered=${problems.length ? "no" : "yes"}`);
      return { id, problems };
    },

    // patch: { figure?, summary?, requirement_ids? }, while it is proposed.
    async update(doc, { ref, id, patch = {} } = {}) {
      const take = {};
      if ("figure" in patch) {
        take.figure = String(patch.figure ?? "").trim();
        if (!take.figure) refuse('Enter the figure label, for example "Figure 2".');
      }
      if ("summary" in patch) take.summary = String(patch.summary ?? "").trim();
      if ("requirement_ids" in patch) {
        if (!Array.isArray(patch.requirement_ids)) refuse("requirement_ids must be a list.");
        take.requirement_ids = patch.requirement_ids.map(String);
      }
      await update(doc, (rev) => {
        figureOf(rev, ref, id);
        return stored(() => updateFigureChange(rev, ref, id, take));
      });
      return { id };
    },

    // The picture of one side as a PNG file: { png }. The path comes from
    // the record of the figure change, never from the caller.
    async render(doc, { ref, id, side } = {}) {
      if (side !== "before" && side !== "after") refuse("side must be before or after.");
      const rev = await load(doc);
      const d = dirs(doc, rev);
      const f = figureOf(rev, ref, id);
      if (!f[side]) refuse(`This figure change has no ${side} file.`, 404);
      const r = await renderSide(d, f, side);
      if (r.error) refuse(r.error, 503);
      if (f.renders?.[side]?.of !== r.render.of || !f.facts?.[side])
        await update(doc, (rev) => {
          const cur = figureOf(rev, ref, id);
          return updateFigureChange(rev, ref, id, {
            renders: { ...cur.renders, [side]: r.render },
            facts: { ...cur.facts, [side]: r.facts },
          });
        });
      return { png: r.png };
    },

    // Copies the new file into the submission folder and records the copy.
    async adopt(doc, { ref, id } = {}) {
      let copy = null;
      await update(doc, async (rev) => {
        const d = dirs(doc, rev);
        const f = figureOf(rev, ref, id);
        if (f.state !== "proposed") refuse(`This figure change is ${f.state} already.`, 409);
        // The source and, when there is one, the view that is submitted.
        await recorded(d, f.after.source, "The new file");
        const file = f.after.view ?? f.after.source;
        const src = await recorded(d, file, "The new file");
        const path = await submit(d, src, file.sha256);
        const at = now();
        copy = { path: storedPath(d.docDir, path), sha256: file.sha256, at: at.toISOString() };
        await sourcesRow(d, [
          copy.at.slice(0, 10),
          f.figure,
          ref,
          "adopted",
          f.after.source.path,
          f.after.source.sha256,
          basename(path),
          copy.sha256,
        ]);
        return setFigureState(rev, ref, id, "adopted", { adopted: copy }, at);
      });
      log(`[figures] adopt ${id} file=${basename(copy.path)}`);
      return { id, adopted: copy };
    },

    async reject(doc, { ref, id } = {}) {
      await update(doc, (rev) => {
        figureOf(rev, ref, id);
        return stored(() => setFigureState(rev, ref, id, "rejected", {}, now()));
      });
      log(`[figures] reject ${id}`);
      return { id };
    },

    // Back to proposed. The copy of an adopted one stays in the submission
    // folder (this service removes no file); FILE_SOURCES.md says that it
    // was withdrawn.
    async reopen(doc, { ref, id } = {}) {
      await update(doc, async (rev) => {
        const f = figureOf(rev, ref, id);
        const next = stored(() => setFigureState(rev, ref, id, "proposed", {}, now()));
        if (f.state === "adopted") {
          const d = dirs(doc, rev);
          await writable(d.docDir, d.submission);
          await sourcesRow(d, [
            now().toISOString().slice(0, 10),
            f.figure,
            ref,
            "withdrawn",
            f.after.source.path,
            f.after.source.sha256,
            basename(f.adopted.path),
            f.adopted.sha256,
          ]);
        }
        return next;
      });
      log(`[figures] reopen ${id}`);
      return { id };
    },
  };
  return api;
}
