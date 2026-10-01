// What the wrappers around the Python document tools share: reading the
// result of the daemon's Python runner, and naming an output file next to
// the document.
//
// The runner is injected as runPy(script, args). Its result is read in
// either shape: { code, json, stderr, error } (the last line of stdout
// already parsed) or { code, stdout, stderr }.

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

// "<folder>/<name>_<tag><ext>" next to the document. When that name is
// taken, a time stamp is added: an existing file is never overwritten.
export function siblingPath(docPath, tag, { exists = () => false, now = new Date() } = {}) {
  if (!docPath) return null;
  const ext = extname(docPath);
  const stem = join(dirname(docPath), `${basename(docPath, ext)}_${tag}`);
  const plain = `${stem}${ext}`;
  return exists(plain) ? `${stem}_${stamp(now)}${ext}` : plain;
}
