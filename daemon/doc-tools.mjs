// What the wrappers around the Python document tools share: reading the
// result of the daemon's Python runner, and naming an output file next to
// the document.
//
// The runner is injected as runPy(script, args). Its result is read in
// either shape: { code, json, stderr, error } (the last line of stdout
// already parsed) or { code, stdout, stderr }.

import { constants, copyFile, link, unlink } from "node:fs/promises";
import { dirname, extname, basename, join } from "node:path";
import { fileURLToPath } from "node:url";

export { SAVE_FIRST } from "./doc-bytes.mjs";

// The repository root: tools/ and scripts/ are found from here.
export const REPO_ROOT = fileURLToPath(new URL("..", import.meta.url));

// The JSON a tool printed: the runner's parsed value, else the whole stdout,
// else its last line. undefined when there is none.
export function pyJson(result) {
  if (result?.json !== undefined) return result.json;
  const out = String(result?.stdout ?? "").trim();
  if (!out) return undefined;
  for (const text of [out, out.split("\n").pop()]) {
    try {
      return JSON.parse(text);
    } catch {
      /* try the next reading */
    }
  }
  return undefined;
}

// Why a tool run failed, in the tool's own words when it said any.
export function pyError(result) {
  const said = String(result?.stderr ?? "")
    .trim()
    .split("\n")
    .pop();
  return said || result?.error || `the tool ended with code ${result?.code}`;
}

const two = (n) => String(n).padStart(2, "0");
export const stamp = (d) =>
  `${d.getFullYear()}${two(d.getMonth() + 1)}${two(d.getDate())}-${two(d.getHours())}${two(d.getMinutes())}${two(d.getSeconds())}`;

// The first of "<stem><ext>", "<stem>-2<ext>", "<stem>-3<ext>", ... that
// is not taken.
export function freeName(stem, ext, exists = () => false) {
  for (let n = 1; ; n++) {
    const path = n === 1 ? `${stem}${ext}` : `${stem}-${n}${ext}`;
    if (!exists(path)) return path;
  }
}

// "<folder>/<name>_<tag><ext>" next to the document. When that name is
// taken, a time stamp is added, and a counter when that is taken too. This
// only picks a name that is free now; that an existing file is never
// overwritten is the job of whoever writes the file (the Python tools
// create their output exclusively, moveExclusive below does the same).
export function siblingPath(docPath, tag, { exists = () => false, now = new Date() } = {}) {
  if (!docPath) return null;
  const ext = extname(docPath);
  const stem = join(dirname(docPath), `${basename(docPath, ext)}_${tag}`);
  return exists(`${stem}${ext}`) ? freeName(`${stem}_${stamp(now)}`, ext, exists) : `${stem}${ext}`;
}

// Did a Python tool refuse because its output file already exists?
export const outputTaken = (result) =>
  result?.code === 2 && /already exists/.test(String(result?.stderr ?? result?.error ?? ""));

// Move a file to a name that must not exist yet. Unlike rename, this fails
// with EEXIST when the name is taken, also when it was taken a moment ago
// by someone else. A hard link is tried first; on a file system without
// hard links the file is copied with the exclusive flag.
export async function moveExclusive(from, to) {
  try {
    await link(from, to);
  } catch (e) {
    if (e?.code === "EEXIST") throw e;
    await copyFile(from, to, constants.COPYFILE_EXCL);
  }
  await unlink(from);
}
