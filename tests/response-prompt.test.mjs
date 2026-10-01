import { test } from "node:test";
import assert from "node:assert/strict";
import {
  buildResponsePrompt,
  buildGlossPrompt,
  parseDraft,
  RESPONSE_SYSTEM,
  RESPONSE_RULES,
} from "../daemon/response-prompt.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";

const change = (id, state, over = {}) => ({
  id,
  kind: "replace",
  find: "30 degrees",
  text: "30 °C",
  state,
  rationale: {
    what: "Unit written as °C.",
    why: "The reviewer asked for it.",
    addresses: "unit",
    requirement_ids: ["q1"],
  },
  ...over,
});
const comment = (over = {}) => ({
  ref: "R1.1",
  reviewer: 1,
  text: "Please state the unit.\nAlso name the supplier of the flasks.",
  requirements: [
    { id: "q1", text: "state the unit" },
    { id: "q2", text: "name the supplier" },
  ],
  declined: null,
  changes: [change("R1.1-1", "accepted")],
  response: null,
  ...over,
});
const profile = (letter = {}) => {
  const p = withDefaults({ response_letter: letter });
  p.response_letter.lang = letter.lang ?? "English";
  return p;
};
const ledger = [
  {
    id: "L1",
    kind: "forbid",
    text: "Do not write 'unfortunately' in the letter.",
    scope: "response",
  },
  { id: "L2", kind: "instruction", text: "Thank the reviewer once, briefly.", scope: "response" },
  { id: "L3", kind: "instruction", text: "Write in vitro without a hyphen." },
  { id: "L4", kind: "forbid", text: "No first person in the manuscript.", scope: "manuscript" },
  { id: "L5", kind: "instruction", text: "Number supplementary notes.", scope: "supplement" },
];
const text = (args) => buildResponsePrompt({ profile: profile(), ...args }).prompt;

test("the prompt holds the comment, its requests and the accepted changes with their rationale", () => {
  const built = buildResponsePrompt({ comment: comment(), profile: profile(), ledger });
  assert.equal(built.system, RESPONSE_SYSTEM);
  const p = built.prompt;
  assert.ok(
    p.includes("Reviewer 1, comment R1.1 (verbatim):\n<<<\nPlease state the unit.\nAlso name"),
  );
  assert.ok(p.includes("- [q1] state the unit"));
  assert.ok(p.includes("- [q2] name the supplier"));
  assert.match(p, /- \[R1\.1-1\] replaced <<<30 degrees>>> with <<<30 °C>>> \(answers q1\)/);
  assert.ok(p.includes("What: Unit written as °C. Why: The reviewer asked for it."));
  assert.deepEqual(built.changeIds, ["R1.1-1"]);
  assert.match(p, /Write in English\./);
  assert.ok(p.trim().endsWith('Reply with JSON only: {"response": "<text>"}'));
});

test("only response-scope ledger items and the profile rules go in", () => {
  const built = buildResponsePrompt({
    comment: comment(),
    profile: profile({ rules: ["Open with one sentence that says what was done."] }),
    ledger,
  });
  const p = built.prompt;
  assert.ok(p.includes("- [L2] Thank the reviewer once, briefly."));
  assert.ok(p.includes("- [L1] Do not write 'unfortunately' in the letter."));
  assert.ok(p.includes("- Open with one sentence that says what was done."));
  for (const other of ["in vitro", "first person", "supplementary notes"])
    assert.ok(!p.includes(other), other);
  assert.deepEqual([built.instructionIds, built.forbidIds], [["L2"], ["L1"]]);
});

test("the language is the profile's; no language is built in", () => {
  const p = text({ comment: comment(), profile: profile({ lang: "German" }) });
  assert.match(p, /Write in German\./);
  assert.doesNotMatch(p, /Korean|English/);
  assert.equal(RESPONSE_RULES.filter((r) => r.includes("{lang}")).length, 1);
});

test("only changes that are in the manuscript may be reported; the others are named as not to claim", () => {
  const c = comment({
    changes: [
      change("R1.1-1", "accepted"),
      change("R1.1-2", "rejected", { find: "cell-free extract", text: "clarified lysate" }),
      change("R1.1-3", "pending", {
        kind: "delete",
        find: "The flasks came from a supplier.",
        text: "",
      }),
      change("R1.1-4", "unlocated", {
        kind: "insert",
        find: "optical density",
        text: " at 600 nm",
      }),
      change("R1.1-5", "user_edited", { find: "novel", text: "new" }),
    ],
  });
  const built = buildResponsePrompt({ comment: c, profile: profile() });
  const p = built.prompt;
  assert.deepEqual(built.changeIds, ["R1.1-1", "R1.1-5"]);
  const [done, notDone] = p.split("Not in the manuscript");
  assert.ok(done.includes("[R1.1-1] replaced <<<30 degrees>>> with <<<30 °C>>>"));
  // The author rewrote the new text: its first wording must not be reported.
  assert.ok(done.includes("[R1.1-5] revised by the author by hand at <<<novel>>>"));
  assert.ok(!done.includes("<<<new>>>"));
  assert.ok(!done.includes("clarified lysate") && !done.includes("600 nm"));
  assert.ok(
    notDone.includes(
      "[R1.1-2] rejected by the author; do not claim: replaced <<<cell-free extract>>> with <<<clarified lysate>>>",
    ),
  );
  assert.ok(
    notDone.includes(
      "[R1.1-3] not decided in Word yet; do not claim: deleted <<<The flasks came from a supplier.>>>",
    ),
  );
  assert.ok(
    notDone.includes(
      "[R1.1-4] not found in the manuscript; do not claim: inserted <<< at 600 nm>>> after <<<optical density>>>",
    ),
  );
  assert.ok(p.includes("Never claim a change that is not listed as in the manuscript."));
});

test("a comment with pending changes only tells the model that nothing was changed yet", () => {
  const c = comment({ changes: [change("R1.1-1", "pending")] });
  const built = buildResponsePrompt({ comment: c, profile: profile() });
  assert.deepEqual(built.changeIds, []);
  assert.ok(built.prompt.includes("Changes in the manuscript for this comment"));
  assert.match(built.prompt, /for this comment[^\n]*:\n- \(none\)\n/);
  assert.ok(
    built.prompt.includes(
      "If no change is listed as in the manuscript, do not say that anything was revised.",
    ),
  );
});

test("decline reasons are printed by the tool: the prompt says so and gives no reason to copy", () => {
  const c = comment({
    declined: { reason: "REASON-OF-THE-COMMENT" },
    response: { text: "", declined_requirements: [{ id: "q2", reason: "REASON-OF-Q2" }] },
  });
  const p = text({ comment: c });
  assert.ok(p.includes("- [q2] name the supplier (declined by the author)"));
  assert.ok(p.includes("The author declined this comment as a whole."));
  assert.ok(!p.includes("REASON-OF-Q2") && !p.includes("REASON-OF-THE-COMMENT"));
  assert.ok(
    p.includes(
      "Do not give reasons for a declined request or a declined comment. The tool prints the author's reasons after your text.",
    ),
  );
  // Without a decline nothing is said about one.
  assert.ok(!text({ comment: comment() }).includes("declined by the author)"));
});

test("positions and quotes are not the model's to write", () => {
  const p = text({ comment: comment() });
  assert.ok(
    p.includes(
      "Do not write location lines, line numbers, section names or quotes of the manuscript. The tool adds them after your text.",
    ),
  );
  assert.ok(p.includes("Use numbers only as they appear below."));
});

test("word cap and earlier draft are given when there are any", () => {
  const p = text({
    comment: comment(),
    profile: profile({ max_words: 120 }),
    previous: "We added the unit in round one.",
  });
  assert.ok(p.includes("At most 120 words."));
  assert.ok(
    p.includes(
      "Earlier draft (keep its facts unless a change above contradicts them):\n<<<\nWe added the unit in round one.\n>>>",
    ),
  );
  const none = text({ comment: comment() });
  assert.ok(!none.includes("At most") && !none.includes("Earlier draft"));
});

test("an editor comment and a review section are named", () => {
  const p = text({ comment: comment({ ref: "E.2", reviewer: 0, section: "Major comments" }) });
  assert.ok(p.includes("Editor, comment E.2 (verbatim):"));
  assert.ok(p.includes("Section of the review this comment is under: Major comments"));
});

test("parseDraft: trimmed text, or the reason there is none", () => {
  assert.deepEqual(parseDraft({ ok: true, json: { response: " a " } }), { text: "a" });
  assert.match(parseDraft({ ok: true, json: { response: "  " } }).error, /no response text/);
  assert.match(parseDraft({ ok: true, json: { other: "x" } }).error, /no response text/);
  assert.match(parseDraft({ ok: true, json: { response: ["a"] } }).error, /no response text/);
  assert.match(parseDraft({ ok: true, text: "not json" }).error, /no response text/);
  assert.deepEqual(parseDraft({ ok: false, error: "timed out" }), { error: "timed out" });
  assert.ok(parseDraft(null).error);
});

test("the translation prompt asks for plain text in the given language", () => {
  const g = buildGlossPrompt("We agree.", "Korean");
  assert.deepEqual(
    [g.prompt, g.expectJson, /Translate into Korean/.test(g.system)],
    ["We agree.", false, true],
  );
});
