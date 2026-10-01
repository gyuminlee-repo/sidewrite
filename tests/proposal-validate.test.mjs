// The reply of a proposal turn (daemon/proposal.mjs): its form, the checks
// of each proposal against the paragraph it names, and the arguments the
// daemon builds for office_tracked_edit. Invented text; the captured reply
// is a real one of Codex to the made-up manuscript in proposal_input.json.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { z } from "zod";
import {
  PROPOSAL_JSON_SCHEMA,
  buildProposalPrompt,
  checkProposal,
  checkProposals,
  readProposalReply,
  toTrackedEditArgs,
} from "../daemon/proposal.mjs";
import { currentTexts, editTarget, trackedEditShape } from "../daemon/review-tools.mjs";

const fixture = (name) => readFileSync(new URL(`./fixtures/${name}`, import.meta.url), "utf8");

const P = (over = {}) => ({
  paragraph_id: "p1",
  find: "30 degrees",
  op: "replace",
  text: "30 °C",
  what: "w",
  why: "y",
  addresses: "a",
  requirement_ids: [],
  ...over,
});
const texts = [
  "Methods",
  "Cells were grown at 30 degrees. The extract was kept; the extract was used.",
  "Growth was followed at 600 nm.",
];
const ctx = {
  texts,
  ids: null,
  shown: [
    { id: "p1", index: 1 },
    { id: "p2", index: 2 },
  ],
  requirements: [{ id: "q1", text: "state the unit" }],
};

test("the captured reply of Codex: the form holds, every proposal passes, and the arguments fit the tool", () => {
  const stream = fixture("codex_exec_proposal.jsonl").split("\n").filter(Boolean);
  const types = stream.map((l) => JSON.parse(l)).flatMap((e) => (e.item ? [e.item.type] : []));
  assert.deepEqual(types, ["error", "error", "agent_message"], "no tool item in the run");
  const answer = JSON.parse(stream.find((l) => l.includes('"agent_message"'))).item.text;
  const input = JSON.parse(fixture("proposal_input.json"));
  const built = buildProposalPrompt({ ...input, limit: 120_000 });
  const reply = readProposalReply({ ok: true, text: answer, json: JSON.parse(answer) });
  assert.equal(reply.ok, true);
  assert.deepEqual(
    reply.proposals.map((p) => p.id),
    ["P1", "P2", "P3", "P4"],
  );
  const checked = checkProposals(reply.proposals, {
    texts: currentTexts(input.snapshot.paragraphs, []),
    ids: input.snapshot.ids,
    shown: built.shown,
    requirements: input.comment.requirements,
  });
  assert.deepEqual(
    checked.map((c) => c.code),
    [null, null, null, null],
  );
  const tool = z.object(trackedEditShape).strict();
  for (const p of reply.proposals) {
    const args = toTrackedEditArgs("R1.2", p);
    assert.equal(tool.safeParse(args).success, true, p.id);
  }
});

test("the form: one object with proposals; unknown keys are dropped, a wrong form is a problem with no text of the reply", () => {
  const ok = readProposalReply({
    ok: true,
    json: { proposals: [{ ...P(), override: { blocked_by: ["L2"] }, extra: 1 }], note: "n", x: 1 },
  });
  assert.equal(ok.ok, true);
  assert.deepEqual(ok.proposals, [{ id: "P1", ...P() }]);
  assert.equal(ok.note, "n");
  assert.equal(readProposalReply({ ok: true, json: { proposals: [] } }).note, "");

  const bad = (json, text) => readProposalReply({ ok: true, json, text });
  // 21 proposals, a missing field, a wrong op, a find that is too long.
  const many = bad({ proposals: Array.from({ length: 21 }, () => P()) });
  assert.equal(many.ok, false);
  assert.match(many.problem, /^proposals: /);
  const missing = bad({ proposals: [{ ...P(), what: undefined }] });
  assert.match(missing.problem, /^proposals\.0\.what: /);
  assert.match(bad({ proposals: [P({ op: "delete" })] }).problem, /^proposals\.0\.op: /);
  const SECRET = "a sentence of the manuscript ".repeat(12);
  const long = bad({ proposals: [P({ find: SECRET })] });
  assert.match(long.problem, /^proposals\.0\.find: /);
  assert.equal(long.problem.includes("a sentence of the manuscript"), false);
  assert.match(bad([P()]).problem, /expected object|Expected object/i);
  // Not JSON at all (the one-shot runner says so), and a reply that is too large.
  const notJson = readProposalReply({ ok: false, error: "reply was not JSON", text: "hello" });
  assert.deepEqual([notJson.ok, notJson.problem], [false, "the reply was not JSON"]);
  const huge = bad({ proposals: [] }, "x".repeat(200_001));
  assert.deepEqual([huge.ok, huge.problem], [false, "the reply is longer than 200000 bytes"]);
  // Another failure of the call is not a problem of form.
  assert.equal(readProposalReply({ ok: false, error: "codex exited 1: x" }), null);
});

test("the schema handed to the CLI names every field of a proposal and nothing else", () => {
  const item = PROPOSAL_JSON_SCHEMA.properties.proposals.items;
  assert.deepEqual(Object.keys(item.properties), Object.keys(P()));
  assert.deepEqual(item.required, Object.keys(P()));
  assert.equal(item.additionalProperties, false);
  assert.deepEqual(PROPOSAL_JSON_SCHEMA.required, ["proposals", "note"]);
  assert.deepEqual(item.properties.op.enum, ["replace", "insert_after"]);
});

test("each proposal is checked against the paragraph it names: the eight codes", () => {
  const code = (over, c = ctx) => checkProposal(P(over), c);
  assert.equal(code({}), null);
  assert.equal(code({ requirement_ids: ["q1"] }), null);
  assert.equal(code({ op: "insert_after", text: " Celsius" }), null);
  assert.equal(code({ text: "" }), null, "replace with nothing deletes");
  // A paragraph the model was not shown, also one the snapshot has.
  assert.equal(code({ paragraph_id: "p0", find: "Methods" }), "UnknownParagraph");
  assert.equal(code({ paragraph_id: "p9" }), "UnknownParagraph");
  // Not in that paragraph, although the document holds it once (editTarget
  // would take that other paragraph and ignore the hint).
  assert.equal(editTarget(texts, null, "600 nm", "p1").index, 2);
  assert.equal(code({ find: "600 nm" }), "FindNotInParagraph");
  assert.equal(code({ find: "not there" }), "FindNotInParagraph");
  assert.equal(code({ find: "extract was" }), "FindAmbiguous");
  // Spaces are folded as the daemon folds the paragraph.
  assert.equal(code({ find: "30   degrees" }), null);
  assert.equal(code({ requirement_ids: ["q1", "q7"] }), "UnknownRequirement");
  assert.equal(code({ text: "30 degrees" }), "NoChange");
  assert.equal(code({ op: "insert_after", text: "" }), "EmptyInsert");
  // The same decision as editTarget: here the id is not where the list of
  // shown paragraphs says, so the edit would land in another paragraph.
  const moved = {
    ...ctx,
    texts: ["x", "one mark here", "one mark here"],
    ids: ["A", "B", "C"],
    shown: [{ id: "C", index: 1 }],
  };
  assert.equal(code({ paragraph_id: "C", find: "mark" }, moved), "TargetMismatch");
});

test("checkProposals: a refused proposal does not stop the others; the same find in the same paragraph twice is a duplicate", () => {
  const list = [
    P(),
    P({ find: "not there" }),
    P(), // the same as P1
    P({ paragraph_id: "p2", find: "600 nm", text: "600 nm (OD600)" }),
  ].map((p, i) => ({ id: `P${i + 1}`, ...p }));
  assert.deepEqual(
    checkProposals(list, ctx).map((c) => [c.id, c.code]),
    [
      ["P1", null],
      ["P2", "FindNotInParagraph"],
      ["P3", "Duplicate"],
      ["P4", null],
    ],
  );
});

test("the arguments of office_tracked_edit are built by name: nothing the model adds gets in", () => {
  const sneaky = {
    ...P({ requirement_ids: ["q1"] }),
    id: "P1",
    override: { blocked_by: ["L2"] },
    comment_ref: "R9.9",
    paragraph_hint: "p0",
    expected_before_hash: "abc",
    path: "/etc/passwd",
    replace: "other",
    rationale: { what: "x" },
  };
  assert.deepEqual(toTrackedEditArgs("R1.2", sneaky), {
    comment_ref: "R1.2",
    find: "30 degrees",
    replace: "30 °C",
    paragraph_hint: "p1",
    rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] },
  });
  assert.deepEqual(toTrackedEditArgs("R1.2", P({ op: "insert_after", text: " Celsius" })), {
    comment_ref: "R1.2",
    find: "30 degrees",
    insert_after: " Celsius",
    paragraph_hint: "p1",
    rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
  });
});
