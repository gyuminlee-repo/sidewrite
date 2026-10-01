import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  mergeProfile,
  profileDirs,
  loadProfile,
  listProfiles,
  effectiveProfile,
  withDefaults,
  rationaleLang,
  responseLang,
  splitPatterns,
  regexProblems,
  LETTER_DEFAULTS,
} from "../daemon/journal-profile.mjs";

const tmp = () => mkdtemp(join(tmpdir(), "prof-"));
async function dirWith(files) {
  const d = await tmp();
  for (const [n, v] of Object.entries(files)) await writeFile(join(d, n), JSON.stringify(v));
  return d;
}

test("merge: dict values one level, the rest replaced", () => {
  assert.deepEqual(
    mergeProfile(
      { colors: { a: "1", b: "2" }, ref_head: ["x"], name: "n" },
      { colors: { b: "9" }, ref_head: ["y"], marked_mode: "color" },
    ),
    { colors: { a: "1", b: "9" }, ref_head: ["y"], name: "n", marked_mode: "color" },
  );
});

test("merge: one level only, a dict over a non-dict replaces it, the base is not changed", () => {
  // The rule of the profile loader this file format comes from:
  // {**base[k], **v} when both are dicts, else v.
  assert.deepEqual(
    mergeProfile(
      { legend_heads: { main: { a: 1, b: 2 } }, x: "s" },
      { legend_heads: { main: { a: 9 } }, x: { y: 1 } },
    ),
    { legend_heads: { main: { a: 9 } }, x: { y: 1 } },
  );
  const base = { colors: { a: "1" } };
  mergeProfile(base, { colors: { a: "2" } });
  assert.deepEqual(base, { colors: { a: "1" } });
});

test("first folder wins, _template is not listed, bad names refused", async () => {
  const user = await dirWith({ "generic.json": { name: "mine" }, "_template.json": {} });
  const builtin = await dirWith({ "generic.json": { name: "built-in" }, "other.json": {} });
  const dirs = [user, builtin];
  assert.equal((await loadProfile("generic", { dirs })).profile.name, "mine");
  assert.deepEqual(await listProfiles({ dirs }), ["generic", "other"]);
  await assert.rejects(loadProfile("../x", { dirs }), /profile name/);
  await assert.rejects(loadProfile("Generic", { dirs }), /profile name/);
  await assert.rejects(loadProfile("_template", { dirs }), /profile name/);
  await assert.rejects(loadProfile("nope", { dirs }), /No profile "nope"/);
});

test("a profile file that is not a JSON object is an error that names the file", async () => {
  const d = await tmp();
  await writeFile(join(d, "broken.json"), "{");
  await writeFile(join(d, "list.json"), "[]");
  await assert.rejects(loadProfile("broken", { dirs: [d] }), /broken\.json/);
  await assert.rejects(loadProfile("list", { dirs: [d] }), /list\.json.*not a JSON object/);
});

test("folder order: manuscript folder, CC_OFFICE_PROFILE_DIRS, user config folder, built-in", () => {
  const d = profileDirs({
    docDir: "/m",
    env: { CC_OFFICE_PROFILE_DIRS: "/a:/b", CC_OFFICE_CONFIG_DIR: "/c" },
  });
  assert.deepEqual(d.slice(0, 4), ["/m/journal_profiles", "/a", "/b", "/c/journal_profiles"]);
  assert.match(d[4], /profiles[\\/]journal_profiles$/);
  assert.equal(d.length, 5);
});

test("effective profile: revision.json, then submission.json, then generic; response defaults filled", async () => {
  const doc = await tmp();
  await writeFile(
    join(doc, "submission.json"),
    JSON.stringify({
      journal_profile: "example-elsevier",
      profile_overrides: { colors: { response_body: "AA0000", response_revised: "00AA00" } },
    }),
  );
  const a = await effectiveProfile(
    { journal_profile: null, profile_overrides: { colors: { response_body: "BB0000" } } },
    { docDir: doc, env: {} },
  );
  assert.equal(a.name, "example-elsevier");
  assert.deepEqual(
    [a.profile.colors.response_body, a.profile.colors.response_revised],
    ["BB0000", "00AA00"],
  );
  assert.deepEqual(a.problems, []);
  const g = await effectiveProfile(
    { journal_profile: null, profile_overrides: {} },
    { docDir: await tmp(), env: {} },
  );
  assert.equal(g.name, "generic");
  assert.deepEqual(
    [
      g.profile.colors.response_body,
      g.profile.colors.response_revised,
      g.profile.colors.response_kr,
    ],
    [null, null, null],
  );
  assert.equal(g.profile.response_letter.location_line, LETTER_DEFAULTS.location_line);
  assert.deepEqual(g.profile.response_letter.rules, []);
  assert.equal(g.profile.response_letter.section_head, "{section}");
  assert.deepEqual(
    [g.profile.rationale_lang, g.profile.response_letter.lang],
    ["English", "English"],
  );
  const e = await effectiveProfile(null, {
    docDir: null,
    env: { CC_OFFICE_PROFILE: "example-elsevier" },
  });
  assert.equal(e.name, "example-elsevier");
});

test("withDefaults fills the letter and colour keys and keeps what the profile set", () => {
  const p = withDefaults({
    colors: { response_body: "AA0000" },
    response_letter: { max_words: 80 },
  });
  assert.deepEqual(p.colors, {
    response_body: "AA0000",
    response_revised: null,
    response_kr: null,
  });
  assert.equal(p.response_letter.max_words, 80);
  assert.equal(p.response_letter.comment_label, "Comment {n}");
  assert.equal("lang" in LETTER_DEFAULTS, false);
});

test("stage 4 keys pass through untouched", async () => {
  const d = await dirWith({
    "j.json": {
      name: "j",
      rules: [{ id: "r1" }],
      policy_packs: { facts_only: false },
      unit_packs: ["bio"],
      italics: { genera: ["Escherichia"] },
      subscripts: ["CO2"],
      csl_style: "x",
    },
  });
  const { profile } = await loadProfile("j", { dirs: [d] });
  assert.deepEqual(
    [
      profile.rules,
      profile.policy_packs,
      profile.unit_packs,
      profile.italics,
      profile.subscripts,
      profile.csl_style,
    ],
    [[{ id: "r1" }], { facts_only: false }, ["bio"], { genera: ["Escherichia"] }, ["CO2"], "x"],
  );
  const eff = await effectiveProfile({ journal_profile: "j" }, { dirs: [d], env: {} });
  assert.deepEqual([eff.profile.rules, eff.profile.csl_style], [[{ id: "r1" }], "x"]);
});

test("format keys: italics.statistics is off unless set; first_line_indent takes consistent", async () => {
  const p = withDefaults({});
  assert.deepEqual([p.italics, p.first_line_indent], [{ statistics: false }, null]);
  // What the profile set is kept, next to its other italics lists.
  const set = withDefaults({
    italics: { statistics: true, genera: ["Escherichia"] },
    first_line_indent: "consistent",
  });
  assert.deepEqual(set.italics, { statistics: true, genera: ["Escherichia"] });
  assert.equal(set.first_line_indent, "consistent");
  // One level merge: an override of one italics key keeps the others.
  const merged = withDefaults(
    mergeProfile({ italics: { genera: ["Escherichia"] } }, { italics: { statistics: true } }),
  );
  assert.deepEqual(merged.italics, { statistics: true, genera: ["Escherichia"] });

  const keys = (profile) => regexProblems(profile).map((s) => s.split(":")[0]);
  for (const ok of [
    null,
    "indented",
    "none",
    "consistent",
    { expect: "consistent", min_pt: 18 },
    {},
  ])
    assert.deepEqual(keys({ first_line_indent: ok }), [], JSON.stringify(ok));
  for (const bad of ["Consistent", "hanging", 12, true, ["none"], { expect: "hanging" }])
    assert.deepEqual(keys({ first_line_indent: bad }), ["first_line_indent"], JSON.stringify(bad));
  assert.deepEqual(keys({ italics: { statistics: "yes" } }), ["italics.statistics"]);
  assert.deepEqual(keys({ italics: ["E. coli"] }), ["italics"]);
  assert.deepEqual(keys({ italics: { statistics: true } }), []);
  // A value of the wrong shape is reported, not replaced.
  assert.deepEqual(withDefaults({ italics: ["E. coli"] }).italics, ["E. coli"]);

  for (const name of ["generic", "example-elsevier"]) {
    const { profile, problems } = await effectiveProfile({ journal_profile: name }, { env: {} });
    assert.deepEqual(problems, [], name);
    assert.equal(profile.italics.statistics, false, name);
    assert.ok("first_line_indent" in (await loadProfile(name)).profile, name);
    assert.ok("statistics" in (await loadProfile(name)).profile.italics, name);
  }
});

test("language: manuscript override, profile, settings default, env, English", () => {
  assert.equal(rationaleLang({}, { env: {} }), "English");
  assert.equal(rationaleLang({}, { env: { CC_OFFICE_RATIONALE_LANG: "Korean" } }), "Korean");
  assert.equal(
    rationaleLang({}, { env: { CC_OFFICE_RATIONALE_LANG: "Korean" }, settingsDefault: "French" }),
    "French",
  );
  assert.equal(
    rationaleLang({ rationale_lang: "German" }, { env: {}, settingsDefault: "French" }),
    "German",
  );
  assert.equal(responseLang({ response_letter: {} }, {}), "English");
  assert.equal(responseLang({ response_letter: { lang: "German" } }, {}), "German");
});

test("splitPatterns: only regexes with a capture group reach the review splitter", () => {
  assert.deepEqual(
    splitPatterns({ reviewer_head: "^Reviewer #(\\d+)", comment_head: "^Comment \\d+ " }),
    { reviewerHead: "^Reviewer #(\\d+)" },
  );
  assert.deepEqual(splitPatterns({ reviewer_head: "^Reviewer #(?:\\d+)", comment_head: "(" }), {});
  assert.deepEqual(splitPatterns({}), {});
});

test("regexProblems names the broken key", () => {
  assert.deepEqual(
    regexProblems({ comment_head: "(", colors: { response_body: "red" } }).map(
      (s) => s.split(":")[0],
    ),
    ["comment_head", "colors.response_body"],
  );
  assert.deepEqual(
    regexProblems({ ref_head: ["ok", "["], response_letter: { opening_regex: "(" } }).map(
      (s) => s.split(":")[0],
    ),
    ["ref_head[1]", "response_letter.opening_regex"],
  );
});
