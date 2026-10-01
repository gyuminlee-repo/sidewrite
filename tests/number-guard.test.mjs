// A change that puts a new number into the manuscript states a fact only the
// author can confirm (the replay of a real revision: a duration the records
// gave was replaced by one computed from other values of the paper). Such a
// change is refused and the model is told to ask the author instead.
// All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import { numbersIn, newNumbers } from "../daemon/number-guard.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";

test("numbersIn: the numbers a text states, thousands commas and decimal commas read as one number", () => {
  assert.deepEqual(
    [...numbersIn("ran for 3,360 hours at 30.5 °C with 0,5 g and 12 flasks")],
    ["3360", "30.5", "0.5", "12"],
  );
  // Not numbers of the text: the digits of a formula or a name, citation
  // brackets, and the number of a figure, table, section, equation or reference.
  assert.deepEqual([...numbersIn("CO2 and H2O in strain K12 [3] [4, 5] [7-9]")], []);
  assert.deepEqual(
    [
      ...numbersIn(
        "see Fig. 2, Figure 3A, Table 1, Tables 2 and 3, Section 4, Eq. 5, Ref. 6 and Fig. S7",
      ),
    ],
    [],
  );
  assert.deepEqual([...numbersIn("Figure 2 shows 15 % more")], ["15"]);
  assert.deepEqual([...numbersIn("")], []);
});

test("newNumbers: numbers the change adds that neither the replaced text nor a source holds", () => {
  const sources = ["The reviewer asks for the temperature in Celsius.", "n = 3"];
  // The made-up form of the judged case: a recorded value replaced by a computed one.
  assert.deepEqual(
    newNumbers({
      kind: "replace",
      find: "The culture ran for 3,360 hours.",
      text: "The culture ran for about 3,020 hours.",
      sources,
    }),
    ["3,020"],
  );
  // Notation only: the same numbers.
  assert.deepEqual(
    newNumbers({ kind: "replace", find: "held at 30 C", text: "held at 30 °C", sources }),
    [],
  );
  assert.deepEqual(
    newNumbers({ kind: "replace", find: "3360 hours", text: "3,360 h", sources }),
    [],
  );
  // A number of the comment, of an author instruction or answer, or of the paragraph itself.
  assert.deepEqual(
    newNumbers({ kind: "insert", find: "were used", text: " (n = 3)", sources }),
    [],
  );
  assert.deepEqual(
    newNumbers({
      kind: "insert",
      find: "were used",
      text: " for 48 h",
      sources: [...sources, "The run lasted 48 h and was repeated twice."],
    }),
    [],
  );
  assert.deepEqual(newNumbers({ kind: "insert", find: "were used", text: " for 48 h", sources }), [
    "48",
  ]);
  // Taking a number out is not adding one.
  assert.deepEqual(newNumbers({ kind: "delete", find: "for 48 h", text: "", sources }), []);
  assert.deepEqual(
    newNumbers({ kind: "replace", find: "for 48 h at 30 °C", text: "at 30 °C", sources }),
    [],
  );
  // A reference to a figure or a citation is not a number of the paper.
  assert.deepEqual(
    newNumbers({ kind: "insert", find: "as shown", text: " in Fig. 4 [12]", sources }),
    [],
  );
});

const R = { what: "w", why: "y", addresses: "a" };
function deps(ctx = {}) {
  const d = {
    calls: [],
    lines: [],
    log: (s) => d.lines.push(s),
    getLedger: async () => [],
    getReviewContext: async () => ({
      requirements: [],
      comment_text: "The run time in the Methods does not match the Results. Please check.",
      paragraphs: ["Title", "The culture ran for 3,360 hours at 30 C in 12 flasks."],
      paragraph_ids: ["a", "b"],
      recent_changes: [],
      author_texts: [],
      ...ctx,
    }),
    call: async (name, args) => {
      d.calls.push([name, args]);
      return {
        ok: true,
        word_comment_id: 1,
        paragraph_index: 1,
        before: "b",
        after: "a",
        target_checked: true,
      };
    },
    onEdit: async () => ({ id: "R1.1-1" }),
    paneChecksTarget: () => true,
  };
  return d;
}
const edit = (find, replace) => ({ comment_ref: "R1.1", find, replace, rationale: R });

test("applyTrackedEdit refuses a change that puts a new number in, before Word is called", async () => {
  const d = deps();
  const r = await applyTrackedEdit(d, edit("3,360 hours", "about 3,020 hours"), { ref: "R1.1" });
  assert.deepEqual([r.ok, r.code, r.numbers], [false, "NumberNeedsAuthor", ["3,020"]]);
  assert.match(r.message, /^Not applied: the change puts a number into the manuscript \(3,020\)/);
  assert.match(r.message, /office_ask_author/);
  assert.match(r.message, /Do not try another number/);
  assert.deepEqual(d.calls, [], "nothing was sent to Word");
  // The log holds the count, not the number.
  assert.deepEqual(d.lines, ["[review] number needs author R1.1 count=1"]);
});

test("applyTrackedEdit lets notation, numbers of the paragraph and numbers the author gave through", async () => {
  // Notation.
  let d = deps();
  assert.equal((await applyTrackedEdit(d, edit("at 30 C", "at 30 °C"), { ref: "R1.1" })).ok, true);
  // A number that is already in the paragraph (moved within it).
  d = deps();
  assert.equal(
    (await applyTrackedEdit(d, edit("ran for", "ran, in 12 flasks, for"), { ref: "R1.1" })).ok,
    true,
  );
  // The author answered the question: the number is theirs now.
  d = deps({ author_texts: ["The run lasted 3,020 hours; 3,360 was a typing error."] });
  assert.equal(
    (await applyTrackedEdit(d, edit("3,360 hours", "3,020 hours"), { ref: "R1.1" })).ok,
    true,
  );
  // The reviewer gave it.
  d = deps({ comment_text: "The run time should read 3,020 hours." });
  assert.equal(
    (await applyTrackedEdit(d, edit("3,360 hours", "3,020 hours"), { ref: "R1.1" })).ok,
    true,
  );
  // A deletion of a number goes by the rules that were there before.
  d = deps();
  assert.equal((await applyTrackedEdit(d, edit(" in 12 flasks", ""), { ref: "R1.1" })).ok, true);
});
