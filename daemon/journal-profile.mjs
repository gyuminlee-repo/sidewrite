// Journal profiles: one module reads, merges and fills them.
//
// A profile is journal_profiles/<name>.json. The format is the profile of the
// journal-revision-package layout (reviewer_head, comment_head, ref_head,
// ref_item, ref_sep, ref_item_cont, ref_caption, supp_ref_corpus, colors ...)
// plus two optional keys for the response letter: response_letter and
// rationale_lang. Keys this module does not know are passed through.
//
// Values of one paper (measured colours, the language of this round) are not
// written into a profile file. They live in profile_overrides of
// revision.json, or of a submission.json next to the manuscript.

import { readFile, readdir } from "node:fs/promises";
import { homedir } from "node:os";
import { delimiter, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

export const BUILTIN_DIR = join(
  dirname(fileURLToPath(import.meta.url)),
  "..",
  "profiles",
  "journal_profiles",
);
export const DEFAULT_PROFILE = "generic";

// A name is a file stem and nothing else: no path part, no leading "_"
// (files that start with "_" are templates, not profiles).
const NAME_RE = /^[a-z0-9][a-z0-9_-]{0,63}$/;
const HEX = /^[0-9A-Fa-f]{6}$/;
const isObj = (x) => !!x && typeof x === "object" && !Array.isArray(x);

// Response letter keys a profile may leave out. The language is not here:
// responseLang() decides it.
export const LETTER_DEFAULTS = {
  reviewer_label: "Reviewer #{n}",
  editor_label: "Editor",
  comment_label: "Comment {n}",
  // Label of a group of comments in the review ("Major comments"); null: none.
  section_head: "{section}",
  location_line: "[Revised in {section}, {lines}]",
  removed_line: "[Removed from {section}, {lines}]",
  // The same two lines for a citation whose section is not known (a
  // manuscript without heading styles, or text before its first heading).
  location_line_no_section: "[Revised, {lines}]",
  removed_line_no_section: "[Removed]",
  declined_line: "[No change made]",
  line_one: "line {a}",
  line_range: "lines {a}-{b}",
  section_level: null,
  rules: [],
  opening_regex: null,
  max_words: null,
  max_quote_words: null,
  forbid_cross_reviewer: false,
  gloss_enabled: false,
  gloss_lang: null,
  gloss_prefix: "",
};
export const COLOR_DEFAULTS = { response_body: null, response_revised: null, response_kr: null };
// italics.statistics: check that statistics symbols (p, n, t ...) are italic.
// Off unless a profile turns it on. The other italics keys are name lists.
export const ITALICS_DEFAULTS = { statistics: false };
// first_line_indent: null (not checked), one of these, or { expect: one of
// these, ... }. "consistent": every body paragraph like the most common one.
export const FIRST_LINE_INDENT = ["indented", "none", "consistent"];

// Dict values merge one level by key, everything else is replaced. A null
// override removes nothing here: the caller deletes keys before it saves.
export function mergeProfile(base, overrides = {}) {
  const out = structuredClone(base ?? {});
  for (const [k, v] of Object.entries(overrides ?? {}))
    out[k] = isObj(v) && isObj(out[k]) ? { ...out[k], ...structuredClone(v) } : structuredClone(v);
  return out;
}

// The profile with every response letter and response colour key present.
export function withDefaults(profile) {
  const p = structuredClone(profile ?? {});
  p.colors = { ...COLOR_DEFAULTS, ...(isObj(p.colors) ? p.colors : {}) };
  p.response_letter = {
    ...structuredClone(LETTER_DEFAULTS),
    ...(isObj(p.response_letter) ? p.response_letter : {}),
  };
  // Format keys of the manuscript checks. A value of the wrong shape is left
  // as it is; regexProblems names it.
  if (p.italics == null || isObj(p.italics)) p.italics = { ...ITALICS_DEFAULTS, ...p.italics };
  p.first_line_indent ??= null;
  // The rule packs of a review turn (review-prompt.mjs POLICY_PACKS): on
  // unless a profile says false.
  p.policy_packs = { review_scope: true, ...(isObj(p.policy_packs) ? p.policy_packs : {}) };
  return p;
}

// Where profiles are looked up, first hit wins: the manuscript folder, the
// folders of CC_OFFICE_PROFILE_DIRS (PATH-style list), the user's config
// folder, the profiles shipped with this repository.
export function profileDirs({ docDir = null, env = process.env } = {}) {
  const extra = String(env.CC_OFFICE_PROFILE_DIRS ?? "")
    .split(delimiter)
    .filter(Boolean);
  const config = env.CC_OFFICE_CONFIG_DIR || join(homedir(), ".claude", "office-addins");
  return [
    ...(docDir ? [join(docDir, "journal_profiles")] : []),
    ...extra,
    join(config, "journal_profiles"),
    BUILTIN_DIR,
  ];
}

// null when the file is not there; an error that names the file otherwise.
async function readJson(path) {
  let raw;
  try {
    raw = await readFile(path, "utf8");
  } catch (e) {
    if (e.code === "ENOENT" || e.code === "ENOTDIR") return null;
    throw new Error(`${path}: ${e.message}`);
  }
  let data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    throw new Error(`${path}: ${e.message}`);
  }
  if (!isObj(data)) throw new Error(`${path}: not a JSON object`);
  return data;
}

// The file as written, nothing filled in.
export async function loadProfile(name, { dirs = profileDirs() } = {}) {
  if (typeof name !== "string" || !NAME_RE.test(name))
    throw new Error(`Bad profile name ${JSON.stringify(name)}`);
  for (const dir of dirs) {
    const path = join(dir, `${name}.json`);
    const profile = await readJson(path);
    if (profile) return { name, path, profile };
  }
  throw new Error(`No profile "${name}" in ${dirs.join(", ")}`);
}

export async function listProfiles({ dirs = profileDirs() } = {}) {
  const names = new Set();
  for (const dir of dirs) {
    for (const f of await readdir(dir).catch(() => [])) {
      if (!f.endsWith(".json")) continue;
      const name = f.slice(0, -5);
      if (NAME_RE.test(name)) names.add(name);
    }
  }
  return [...names].sort();
}

// Language of the rationale comments and of the response text. Order: the
// manuscript override (already merged into `profile`), the profile, the
// settings file default, the environment (rationale only), English.
export const rationaleLang = (profile, { env = process.env, settingsDefault = null } = {}) =>
  profile?.rationale_lang || settingsDefault || env.CC_OFFICE_RATIONALE_LANG || "English";
export const responseLang = (profile, { settingsDefault = null } = {}) =>
  profile?.response_letter?.lang || settingsDefault || "English";

// paper: { journal_profile, profile_overrides } of revision.json. A
// submission.json next to the manuscript gives the same two keys as a
// fallback; an override of the paper wins over one of submission.json.
export async function effectiveProfile(
  paper,
  { docDir = null, env = process.env, settingsDefault = null, dirs } = {},
) {
  const sub = docDir ? await readJson(join(docDir, "submission.json")) : null;
  const name =
    paper?.journal_profile || sub?.journal_profile || env.CC_OFFICE_PROFILE || DEFAULT_PROFILE;
  const overrides = mergeProfile(
    isObj(sub?.profile_overrides) ? sub.profile_overrides : {},
    isObj(paper?.profile_overrides) ? paper.profile_overrides : {},
  );
  const base = await loadProfile(name, { dirs: dirs ?? profileDirs({ docDir, env }) });
  const merged = mergeProfile(base.profile, overrides);
  const profile = withDefaults(merged);
  profile.rationale_lang = rationaleLang(merged, { env, settingsDefault });
  profile.response_letter.lang = responseLang(merged, { settingsDefault });
  return { name, path: base.path, overrides, profile, problems: regexProblems(merged) };
}

// Profile regexes are written for Python `re`. This reports the ones
// JavaScript cannot compile and response colours that are not RRGGBB, as
// "<key>: <reason>" strings.
const REGEX_KEYS = [
  "reviewer_head",
  "comment_head",
  "ref_item",
  "ref_sep",
  "ref_item_cont",
  "ref_caption",
  "body_section_head",
];
export function regexProblems(profile) {
  const p = profile ?? {};
  const out = [];
  const tryRe = (key, src) => {
    try {
      new RegExp(src);
    } catch (e) {
      out.push(`${key}: ${e.message}`);
    }
  };
  for (const k of REGEX_KEYS) if (typeof p[k] === "string") tryRe(k, p[k]);
  for (const [i, h] of (Array.isArray(p.ref_head) ? p.ref_head : []).entries())
    tryRe(`ref_head[${i}]`, h);
  if (p.response_letter?.opening_regex)
    tryRe("response_letter.opening_regex", p.response_letter.opening_regex);
  for (const [k, v] of Object.entries(isObj(p.colors) ? p.colors : {}))
    if (k.startsWith("response_") && v != null && !HEX.test(String(v)))
      out.push(`colors.${k}: not RRGGBB`);
  if (p.italics != null && !isObj(p.italics)) out.push("italics: not an object");
  else if (p.italics?.statistics != null && typeof p.italics.statistics !== "boolean")
    out.push("italics.statistics: not true or false");
  const indent = isObj(p.first_line_indent) ? p.first_line_indent.expect : p.first_line_indent;
  if (
    (indent != null && !FIRST_LINE_INDENT.includes(indent)) ||
    (p.first_line_indent != null && !isObj(p.first_line_indent) && typeof indent !== "string")
  )
    out.push(`first_line_indent: not one of ${FIRST_LINE_INDENT.join(", ")}`);
  return out;
}

// Number of capture groups, 0 for a regex that does not compile.
const groups = (src) => {
  try {
    return new RegExp(`${src}|`).exec("").length - 1;
  } catch {
    return 0;
  }
};

// Head patterns for the review splitter, which reads the reviewer and
// comment number from a capture group. A profile regex without a group is
// left out, so the splitter keeps its own default for that head.
export function splitPatterns(profile) {
  const out = {};
  if (typeof profile?.reviewer_head === "string" && groups(profile.reviewer_head) > 0)
    out.reviewerHead = profile.reviewer_head;
  if (typeof profile?.comment_head === "string" && groups(profile.comment_head) > 0)
    out.commentHead = profile.comment_head;
  return out;
}
