// Look up a journal's citation style in the public CSL styles repository.
// Lookup and information only: nothing here formats a reference. Off unless
// the caller passes enabled: true (the user's consent switch). fetchFn and
// the cache folder are injected. Never throws.
//
// Checked (2026-10-01): one style is one file at <RAW>/<slug>.csl, a
// dependent style at <RAW>/dependent/<slug>.csl with a link
// rel="independent-parent". The repository is licensed CC BY-SA 3.0 and asks
// software that uses its styles to name the CSL project and link to
// https://citationstyles.org/. A fetched style is written to the user's
// cache folder byte for byte (its author and contributor list stays as it
// is) and is never put into this repository.
//
// Not wired yet. The caller reads the switch from the settings, passes
// <config dir>/csl as cacheDir and shows CSL_NOTICE next to the lookup, also
// before the switch is on.

import { mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { FETCH_TIMEOUT_MS, deadline } from "./fetch-timeout.mjs";

export const CSL_RAW = "https://raw.githubusercontent.com/citation-style-language/styles/master";

export const CSL_NOTICE = [
  "Citation styles come from the Citation Style Language project (https://citationstyles.org/).",
  "A style is downloaded from the public CSL styles repository on github.com when you ask for one; the name of the style is sent there.",
  "The styles are licensed CC BY-SA 3.0 (https://creativecommons.org/licenses/by-sa/3.0/).",
  "A downloaded style is kept unchanged in your user cache folder on this PC. It is not included in this tool and is not redistributed by it.",
].join(" ");

export const CSL_OFF =
  "The style lookup is off. Turn it on in Setup; the style name leaves this PC.";

const SLUG = /^[a-z0-9][a-z0-9-]{0,99}$/;

const squash = (s) =>
  String(s ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();

// index: [{ title, titleShort?, name, dependent }]. An exact title (or short
// title) first, then titles that contain the words; five at most.
export function findStyles(query, index, { limit = 5 } = {}) {
  const q = squash(query);
  if (!q || !Array.isArray(index)) return [];
  const exact = index.filter((s) => squash(s.title) === q || squash(s.titleShort) === q);
  const part = index.filter((s) => !exact.includes(s) && squash(s.title).includes(q));
  return [...exact, ...part].slice(0, limit);
}

export function styleUrl({ name, dependent } = {}) {
  return `${CSL_RAW}/${dependent ? "dependent/" : ""}${name}.csl`;
}

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const decode = (s) =>
  String(s ?? "")
    .replace(/&(amp|lt|gt|quot|apos);/g, (_, n) => ENTITIES[n])
    .replace(/\s+/g, " ")
    .trim();
const attr = (tag, name) => new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`).exec(tag ?? "")?.[1] ?? null;

// The <info> of a CSL style, or null when the text is not one.
export function parseStyleInfo(xml) {
  const text = String(xml ?? "");
  if (!/<style\b[^>]*\bxmlns\s*=\s*"http:\/\/purl\.org\/net\/xbiblio\/csl"/.test(text)) return null;
  const info = /<info>([\s\S]*?)<\/info>/.exec(text)?.[1];
  const title = /<title>([\s\S]*?)<\/title>/.exec(info ?? "")?.[1];
  if (!info || !title) return null;
  const names = (tag) =>
    [...info.matchAll(new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`, "g"))]
      .map((m) => /<name>([\s\S]*?)<\/name>/.exec(m[1])?.[1])
      .filter(Boolean)
      .map(decode);
  const parentLink = (info.match(/<link\b[^>]*>/g) ?? []).find(
    (l) => attr(l, "rel") === "independent-parent",
  );
  const rights = /<rights\b([^>]*)>([\s\S]*?)<\/rights>/.exec(info);
  return {
    title: decode(title),
    id: decode(/<id>([\s\S]*?)<\/id>/.exec(info)?.[1]) || null,
    parent: parentLink ? (attr(parentLink, "href") ?? "").split("/").pop() || null : null,
    format: attr(/<category\b[^>]*citation-format[^>]*>/.exec(info)?.[0], "citation-format"),
    rights: rights ? decode(rights[2]) : null,
    license: rights ? attr(rights[1], "license") : null,
    authors: names("author"),
    contributors: names("contributor"),
  };
}

// slug: the style's file name without ".csl". dependent: true, false, or
// unknown (the top folder is asked first, then dependent/).
export async function fetchStyle({
  slug,
  dependent,
  enabled = false,
  fetchFn,
  cacheDir,
  timeoutMs = FETCH_TIMEOUT_MS,
  log = () => {},
} = {}) {
  const fail = (error) => ({ ok: false, error, notice: CSL_NOTICE });
  if (enabled !== true) return fail(CSL_OFF);
  if (typeof slug !== "string" || !SLUG.test(slug)) {
    return fail("A style name has lower-case letters, digits and hyphens only.");
  }

  // { bytes, info, cached } of one style, or { error }.
  const load = async (name, folders) => {
    const path = join(cacheDir, `${name}.csl`);
    try {
      const bytes = await readFile(path);
      const info = parseStyleInfo(bytes.toString("utf8"));
      if (info) return { bytes, info, path, cached: true };
    } catch {
      /* not cached */
    }
    for (const folder of folders) {
      let res;
      try {
        res = await fetchFn(styleUrl({ name, dependent: folder === "dependent" }), {
          signal: deadline(timeoutMs),
        });
      } catch (e) {
        log(`[csl] style=${name} status=failed`);
        return { error: `The styles repository could not be reached: ${String(e?.message ?? e)}` };
      }
      log(`[csl] style=${name} status=${res.status}`);
      if (res.status === 404) continue;
      if (!res.ok) return { error: `The styles repository answered ${res.status}` };
      // The connection can still break, or the time run out, while the body is read.
      let bytes;
      try {
        bytes = Buffer.from(await res.arrayBuffer());
      } catch (e) {
        log(`[csl] style=${name} status=failed`);
        return { error: `The style could not be read to its end: ${String(e?.message ?? e)}` };
      }
      const info = parseStyleInfo(bytes.toString("utf8"));
      if (!info) return { error: `${name} is not a CSL style` };
      return { bytes, info, path, cached: false };
    }
    return { missing: true };
  };

  const folders =
    dependent === true ? ["dependent"] : dependent === false ? ["top"] : ["top", "dependent"];
  const style = await load(slug, folders);
  if (style.missing) return fail(`The styles repository has no style named ${slug}.`);
  if (style.error) return fail(style.error);
  let parent = null;
  if (style.info.parent) {
    if (!SLUG.test(style.info.parent))
      return fail(`${slug} names a parent style that is not a style name.`);
    parent = await load(style.info.parent, ["top"]);
    if (parent.missing || parent.error) {
      return fail(
        `The parent style ${style.info.parent} could not be read: ${parent.error ?? "not found"}`,
      );
    }
  }
  const got = [style, parent].filter(Boolean);
  try {
    await mkdir(cacheDir, { recursive: true, mode: 0o700 });
    for (const s of got) if (!s.cached) await writeFile(s.path, s.bytes);
  } catch (e) {
    return fail(`The style could not be saved: ${String(e?.message ?? e)}`);
  }
  return {
    ok: true,
    info: { ...style.info, slug },
    parent: parent ? { ...parent.info, slug: style.info.parent } : null,
    files: got.map((s) => s.path),
    notice: CSL_NOTICE,
  };
}
