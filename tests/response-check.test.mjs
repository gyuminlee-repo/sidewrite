import { test } from "node:test";
import assert from "node:assert/strict";
import { checkResponses, summarize, wordCount } from "../daemon/response-check.mjs";
import { snapshotHash } from "../daemon/response-quotes.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";

const QUOTE = "The incubation temperature was held at 30 °C.";
const SECTION = "2. Materials and Methods";

// A complete letter for one comment: answered, its one request cited, the
// quote in the manuscript, the line range from the current PDF, no ledger hit.
// mutate(revision, input) plants one defect.
function good(mutate = () => {}, letter = {}) {
  const snapshot = {
    paragraphs: [
      { index: 0, text: SECTION, w_ins: 0, w_del: 0 },
      {
        index: 1,
        text: `Flasks were shaken. ${QUOTE} Samples were drawn daily.`,
        w_ins: 0,
        w_del: 0,
      },
    ],
  };
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Please state the unit.",
        requirements: [{ id: "q1", text: "state the unit" }],
        declined: null,
        changes: [
          {
            id: "R1.1-1",
            kind: "replace",
            find: "degrees",
            text: "°C",
            state: "accepted",
            paragraph_index: 1,
            rationale: { requirement_ids: ["q1"] },
          },
        ],
        response: {
          text: "We now state the unit.",
          previous: null,
          declined_requirements: [],
          stale: [],
          citations: [
            {
              change_ids: ["R1.1-1"],
              kind: "revised",
              section: SECTION,
              paragraph_index: 1,
              quote: QUOTE,
              lines: { start: 9, end: 9, start_page: 1, end_page: 1, match_count: 1 },
              lines_pdf: "sha",
              problem: null,
            },
          ],
        },
      },
    ],
    ledger: [
      {
        id: "L1",
        kind: "forbid",
        text: "do not write 'unfortunately' in the letter",
        scope: "response",
        check: { type: "regex", pattern: String.raw`\bunfortunately\b`, flags: "i" },
      },
      // A manuscript rule: it must not be applied to the letter ("We" is in every answer).
      {
        id: "L2",
        kind: "forbid",
        text: "no first person in the manuscript",
        check: { type: "regex", pattern: String.raw`\bWe\b`, flags: "" },
      },
    ],
    response_render: {
      pdf_sha: "sha",
      snapshot_hash: snapshotHash(snapshot),
      at: "2000-01-01T00:00:00.000Z",
      lines_indexed: 40,
    },
  };
  const input = {
    revision,
    snapshot,
    headings: [{ index: 0, level: 1, text: SECTION }],
    profile: withDefaults({ response_letter: letter }),
  };
  mutate(revision, input);
  return input;
}
const ids = (rows, level = "fail") =>
  rows.filter((r) => r.level === level).map((r) => `${r.id}:${r.ref ?? ""}`);

test("a complete letter passes every axis", () => {
  const rows = checkResponses(good());
  assert.deepEqual(rows, []);
  assert.deepEqual(summarize(rows), { fail: 0, warn: 0 });
});

test("each axis fails alone", () => {
  assert.deepEqual(ids(checkResponses(good((r) => (r.comments[0].response.text = "")))), [
    "answered:R1.1",
  ]);
  assert.deepEqual(ids(checkResponses(good((r) => (r.comments[0].response = null)))), [
    "answered:R1.1",
  ]);
  assert.deepEqual(
    ids(
      checkResponses(
        good((r) => r.comments[0].requirements.push({ id: "q2", text: "name the supplier" })),
      ),
    ),
    ["requirements:R1.1"],
  );
  assert.deepEqual(
    ids(checkResponses(good((r) => (r.comments[0].response.citations[0].quote += "x")))),
    ["quote_exact:R1.1"],
  );
  assert.deepEqual(
    ids(checkResponses(good((r) => (r.comments[0].response.citations[0].lines = null)))),
    ["lines:R1.1"],
  );
  assert.deepEqual(
    ids(checkResponses(good((r) => (r.comments[0].response.citations[0].lines.match_count = 2)))),
    ["lines:R1.1"],
  );
  assert.deepEqual(
    ids(checkResponses(good((r) => (r.comments[0].response.citations[0].section = "Methods")))),
    ["section:R1.1"],
  );
  assert.deepEqual(
    ids(checkResponses(good((r) => (r.comments[0].response.text += " Unfortunately so.")))),
    ["ledger:R1.1"],
  );
  assert.deepEqual(ids(checkResponses(good((r) => (r.response_render.snapshot_hash = "old")))), [
    "lines:R1.1",
  ]);
});

test("line ranges: from another PDF, never looked up, or running backwards", () => {
  assert.deepEqual(
    ids(checkResponses(good((r) => (r.comments[0].response.citations[0].lines_pdf = "other")))),
    ["lines:R1.1"],
  );
  assert.deepEqual(ids(checkResponses(good((r) => (r.response_render = null)))), ["lines:R1.1"]);
  assert.deepEqual(
    ids(checkResponses(good((r) => (r.comments[0].response.citations[0].lines.start = 12)))),
    ["lines:R1.1"],
  );
});

test("the quote must be in its own paragraph as the reviewer will read it", () => {
  // Edited in Word since the draft: not in the manuscript any more.
  const edited = good((r, input) => {
    input.snapshot.paragraphs[1].text = input.snapshot.paragraphs[1].text.replace("held", "kept");
    r.response_render.snapshot_hash = snapshotHash(input.snapshot);
  });
  assert.deepEqual(ids(checkResponses(edited)), ["quote_exact:R1.1"]);
  // A pending deletion still shows in text; reviewed_text is what counts.
  const pending = good((r, input) => {
    const p = input.snapshot.paragraphs[1];
    p.reviewed_text = p.text;
    p.text = p.text.replace("30 °C", "30 °Cdegrees");
  });
  assert.deepEqual(ids(checkResponses(pending)), []);
  // A citation that could not be placed.
  const lost = good((r) => {
    Object.assign(r.comments[0].response.citations[0], {
      problem: "sentence not found",
      quote: "",
      lines: null,
    });
  });
  assert.deepEqual(ids(checkResponses(lost)), ["quote_exact:R1.1"]);
});

test("the cited changes must be the accepted changes", () => {
  const rejected = good((r) => (r.comments[0].changes[0].state = "rejected"));
  assert.deepEqual(ids(checkResponses(rejected)), ["requirements:R1.1", "stale:R1.1"]);
  const pending = good((r) => (r.comments[0].changes[0].state = "pending"));
  const rows = checkResponses(pending);
  assert.deepEqual(ids(rows), ["requirements:R1.1", "stale:R1.1"]);
  assert.match(rows[0].detail, /q1 has only pending changes in Word/);
});

test("a declined requirement or comment with a reason counts as answered", () => {
  const declinedReq = good((r) => {
    r.comments[0].requirements.push({ id: "q2", text: "name the supplier" });
    r.comments[0].response.declined_requirements = [{ id: "q2", reason: "Not recorded." }];
  });
  assert.deepEqual(ids(checkResponses(declinedReq)), []);
  const declinedComment = good((r) => {
    Object.assign(r.comments[0], {
      declined: { reason: "Outside the scope of this study." },
      response: null,
      changes: [],
    });
  });
  assert.deepEqual(checkResponses(declinedComment), []);
});

test("a removal needs a section but no quote and no line range", () => {
  const removed = good((r) => {
    Object.assign(r.comments[0].response.citations[0], {
      kind: "removed",
      quote: "",
      lines: null,
      lines_pdf: null,
    });
  });
  assert.deepEqual(checkResponses(removed), []);
  const nowhere = good((r) => {
    Object.assign(r.comments[0].response.citations[0], {
      kind: "removed",
      quote: "",
      lines: null,
      section: null,
    });
  });
  assert.deepEqual(ids(checkResponses(nowhere)), ["section:R1.1"]);
});

test("profile rules: word cap, opening, cross-reviewer; counts come from code", () => {
  assert.equal(wordCount("We agree. 30 °C was used."), 6);
  assert.equal(wordCount(""), 0);
  assert.equal(wordCount(null), 0);
  assert.deepEqual(ids(checkResponses(good(() => {}, { max_words: 3 }))), ["rule_words:R1.1"]);
  assert.match(checkResponses(good(() => {}, { max_words: 3 }))[0].detail, /^5 words, limit 3$/);
  assert.deepEqual(ids(checkResponses(good(() => {}, { max_words: 5 }))), []);
  assert.deepEqual(ids(checkResponses(good(() => {}, { opening_regex: "^We " }))), []);
  assert.deepEqual(ids(checkResponses(good(() => {}, { opening_regex: "^Thank you" }))), [
    "rule_opening:R1.1",
  ]);
  const xref = (text) =>
    ids(
      checkResponses(
        good((r) => (r.comments[0].response.text = text), { forbid_cross_reviewer: true }),
      ),
    );
  assert.deepEqual(xref("We now state the unit, as Reviewer 2 also asked."), ["rule_xref:R1.1"]);
  assert.deepEqual(xref("We now state the unit, as Referee #3 asked."), ["rule_xref:R1.1"]);
  assert.deepEqual(xref("We thank the reviewer and now state the unit."), []);
  // Off by default.
  assert.deepEqual(
    ids(
      checkResponses(
        good((r) => (r.comments[0].response.text = "As Reviewer 2 asked, we state the unit.")),
      ),
    ),
    [],
  );
  // A broken opening regex is reported, not thrown.
  assert.deepEqual(ids(checkResponses(good(() => {}, { opening_regex: "(" })), "warn"), [
    "rule_opening:R1.1",
  ]);
});

test("positions belong to the tool: a location line or a line number in the prose fails", () => {
  const prose = (text) => ids(checkResponses(good((r) => (r.comments[0].response.text = text))));
  assert.deepEqual(prose("We now state the unit.\n[Revised in 3. Results, lines 4-5]"), [
    "prose_position:R1.1",
  ]);
  assert.deepEqual(prose("We now state the unit (line 12)."), ["prose_position:R1.1"]);
  assert.deepEqual(prose("We now state the unit in lines 12-14."), ["prose_position:R1.1"]);
  // Not a line number: other words, a number that is no position.
  assert.deepEqual(prose("We redrew the baseline 5 times and the cell line 3T3 [see ref. 4]."), []);
  assert.deepEqual(prose("The outline and the guidelines were followed."), []);
  // The wording comes from the profile.
  const de = good((r) => (r.comments[0].response.text = "Siehe Zeile 12."), {
    line_one: "Zeile {a}",
    line_range: "Zeilen {a}-{b}",
  });
  assert.deepEqual(ids(checkResponses(de)), ["prose_position:R1.1"]);
});

test("warnings: numbers lost since the previous draft, pending tracked changes, long quote", () => {
  const lost = good((r) => {
    r.comments[0].response.previous = "We now state the unit as 30 °C in 5 flasks.";
  });
  const rows = checkResponses(lost);
  assert.deepEqual(ids(rows), []);
  assert.deepEqual(ids(rows, "warn"), ["lost_facts:R1.1"]);
  assert.match(rows[0].detail, /30 °C/);

  const pending = good((r, input) => {
    input.snapshot.paragraphs[1].w_ins = 1;
  });
  assert.deepEqual(ids(checkResponses(pending), "warn"), ["pending_changes:"]);

  assert.deepEqual(ids(checkResponses(good(() => {}, { max_quote_words: 3 })), "warn"), [
    "rule_quote_words:R1.1",
  ]);
  assert.deepEqual(ids(checkResponses(good(() => {}, { max_quote_words: 8 })), "warn"), []);
  assert.deepEqual(summarize(checkResponses(lost)), { fail: 0, warn: 1 });
});

test("ledger: a broken response-scope regex is a warning; a must-have phrase is checked per answer", () => {
  const broken = good((r) => {
    r.ledger[0].check.pattern = "(";
  });
  assert.deepEqual(ids(checkResponses(broken), "warn"), ["ledger:R1.1"]);
  const mustHave = good((r) => {
    r.ledger.push({
      id: "L3",
      kind: "instruction",
      text: "thank the reviewer",
      scope: "response",
      check: { type: "regex", pattern: "thank", flags: "i", expect: "present" },
    });
  });
  const rows = checkResponses(mustHave);
  assert.deepEqual(ids(rows), ["ledger:R1.1"]);
  assert.match(rows[0].detail, /^L3: missing$/);
});

test("several comments: every row names its comment", () => {
  const two = good((r) => {
    r.comments.push({
      ref: "E.1",
      reviewer: 0,
      text: "Shorten the title.",
      requirements: [],
      declined: null,
      changes: [],
      response: null,
    });
  });
  assert.deepEqual(ids(checkResponses(two)), ["answered:E.1"]);
});
