import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  extractGuidelineRules,
  candidatesFromLabels,
  guidelinePrompt,
  numbersIn,
  htmlToText,
  fetchGuidelineText,
  GUIDELINE_SYSTEM,
  GUIDELINE_WARNING,
} from "../daemon/guideline-extract.mjs";

// The real splitter is injected by the caller; the tests use a plain one.
const splitSentences = (text) =>
  String(text)
    .split(/(?<=\.)\s+/)
    .map((s) => s.trim())
    .filter(Boolean);

const text =
  "The abstract must not exceed 250 words. Provide 4 to 6 keywords. We welcome submissions. Species names are set in italics.";
const sentences = splitSentences(text);
const reply = {
  sentences: [
    { s: 1, type: "count", rule: { target: "section_words", section: "abstract", max: 250 } },
    { s: 2, type: "count", rule: { target: "keywords", min: 4, max: 6 } },
    { s: 3, type: "other" },
    { s: 4, type: "style" },
  ],
};
const tail = reply.sentences.slice(1);
const head3 = reply.sentences.slice(0, 3);
const answering = (json) => async () => ({ ok: true, json });
const others = (n, from = 1) =>
  Array.from({ length: n }, (_, i) => ({ s: from + i, type: "other" }));
const count = (rule) => ({ s: 1, type: "count", rule });

test("every rule is the guideline sentence as written and cites it", async () => {
  const r = await extractGuidelineRules(text, { oneshot: answering(reply), splitSentences });
  assert.deepEqual(r.warnings, []);
  assert.equal(r.candidates.length, 3);
  assert.deepEqual(r.candidates[0], {
    kind: "instruction",
    text: "The abstract must not exceed 250 words.",
    scope: "manuscript",
    check: {
      type: "count",
      target: "section_words",
      section: "^(?:[\\d.]+\\s+)?(?:abstract|summary)\\b",
      max: 250,
    },
    source: { kind: "guideline", sentence: "The abstract must not exceed 250 words." },
  });
  assert.deepEqual(r.candidates[1].check, { type: "count", target: "keywords", min: 4, max: 6 });
  assert.deepEqual(r.candidates[2].check, { type: "llm" });
  for (const c of r.candidates) {
    assert.equal(c.text, c.source.sentence);
    assert.ok(text.includes(c.text));
  }
});

test("the section pattern matches numbered and plain headings, not other sections", async () => {
  const r = await extractGuidelineRules(text, { oneshot: answering(reply), splitSentences });
  const re = new RegExp(r.candidates[0].check.section, "i");
  for (const title of ["Abstract", "ABSTRACT", "1. Abstract", "Summary"]) {
    assert.ok(re.test(title), title);
  }
  for (const title of ["Graphical abstract", "Abstracts of talks", "Results"]) {
    assert.ok(!re.test(title), title);
  }
});

test("the model is sent numbered sentences and asked for labels only", async () => {
  const seen = [];
  await extractGuidelineRules(text, {
    oneshot: async (req) => (seen.push(req), { ok: true, json: reply }),
    splitSentences,
  });
  assert.equal(seen.length, 1);
  assert.equal(seen[0].system, GUIDELINE_SYSTEM);
  assert.match(seen[0].system, /Do not return any text of the guidelines/);
  assert.equal(seen[0].prompt, guidelinePrompt(sentences));
  assert.match(seen[0].prompt, /\[s1\] The abstract must not exceed 250 words\./);
  assert.match(seen[0].prompt, /\[s4\] Species names are set in italics\./);
});

test("a number that is not in the cited sentence rejects the count rule; the sentence stays as a model-checked rule with a warning", () => {
  for (const rule of [
    { target: "section_words", section: "abstract", max: 300 },
    // 25 is inside 250; it is not a number of the sentence.
    { target: "section_words", section: "abstract", max: 25 },
    { target: "section_words", section: "abstract", min: 250, max: 2500 },
  ]) {
    const r = candidatesFromLabels(sentences, { sentences: [count(rule), ...others(3, 2)] });
    assert.equal(r.candidates.length, 1);
    assert.deepEqual(r.candidates[0].check, { type: "llm" });
    assert.equal(r.candidates[0].text, sentences[0]);
    assert.equal(r.warnings.length, 1);
    assert.match(r.warnings[0], /s1\b/);
  }
  const r = candidatesFromLabels(sentences, {
    sentences: [count({ target: "section_words", section: "abstract", max: 300 }), ...others(3, 2)],
  });
  assert.match(r.warnings[0], /300/);
});

test("a thousands separator does not change a number: 1,000, 1 000 and 1000 are one number", () => {
  assert.deepEqual(numbersIn("at most 1,000 words"), [1000]);
  assert.deepEqual(numbersIn("at most 1 000 words or 12 000 characters"), [1000, 12000]);
  assert.deepEqual(
    numbersIn("between 4 and 6 keywords, 250 words, 1,250,000 bytes"),
    [4, 6, 250, 1250000],
  );
  // Decimals and version-like numbers are not whole numbers of the sentence.
  assert.deepEqual(numbersIn("a 2.5 cm margin, size 1,5 and section 3.2.1"), []);
  const rule = (max) => ({ target: "section_words", section: "abstract", max });
  const check = (sentence, max) =>
    candidatesFromLabels([sentence], { sentences: [count(rule(max))] }).candidates[0].check.type;
  for (const written of ["1,000", "1 000", "1 000", "1 000", "1000"]) {
    assert.equal(check(`The abstract has at most ${written} words.`, 1000), "count", written);
  }
  // A part of a grouped number is not that number.
  for (const max of [100, 1, 0, 10000]) {
    assert.equal(check("The abstract has at most 1,000 words.", max), "llm", String(max));
    assert.equal(check("The abstract has at most 1 000 words.", max), "llm", String(max));
  }
  assert.equal(check("The abstract has at most 1.000 words.", 1000), "llm");
  assert.equal(check("The abstract has at most 2.5 pages.", 2), "llm");
});

test("a section order must name sections that the sentence names", () => {
  const two = ["The manuscript should be arranged as Introduction, Methods, Results.", "b"];
  const ok = candidatesFromLabels(two, {
    sentences: [
      count({ target: "section_order", order: ["introduction", "methods", "results"] }),
      { s: 2, type: "other" },
    ],
  });
  assert.deepEqual(ok.candidates[0].check, {
    type: "count",
    target: "section_order",
    order: ["introduction", "methods", "results"],
  });
  const invented = candidatesFromLabels(two, {
    sentences: [
      count({ target: "section_order", order: ["introduction", "discussion"] }),
      { s: 2, type: "other" },
    ],
  });
  assert.deepEqual(invented.candidates[0].check, { type: "llm" });
  assert.match(invented.warnings[0], /discussion/);
});

// ---- replies that must never yield a rule ----
const BAD_REPLIES = {
  "a sentence missing": { sentences: tail },
  "a sentence labelled twice": { sentences: [...head3, { s: 1, type: "other" }] },
  "one label too many": { sentences: [...reply.sentences, { s: 5, type: "other" }] },
  "a number out of range": { sentences: [...head3, { s: 9, type: "other" }] },
  "a number as a string": { sentences: [{ ...reply.sentences[0], s: "1" }, ...tail] },
  "an unknown type": { sentences: [...head3, { s: 4, type: "rule" }] },
  "text beside the label": {
    sentences: [...head3, { s: 4, type: "style", text: "Use italics." }],
  },
  "text inside the rule": {
    sentences: [count({ target: "keywords", max: 6, note: "at most six" }), ...tail],
  },
  "a count label without a rule": { sentences: [{ s: 1, type: "count" }, ...tail] },
  "a rule on a style label": {
    sentences: [...head3, { s: 4, type: "style", rule: { target: "keywords", max: 6 } }],
  },
  "an unknown target": { sentences: [count({ target: "pages", max: 250 }), ...tail] },
  "a section outside the list": {
    sentences: [count({ target: "section_words", section: ".*", max: 250 }), ...tail],
  },
  "a word limit without a section": {
    sentences: [count({ target: "section_words", max: 250 }), ...tail],
  },
  "a section on a keyword count": {
    sentences: [count({ target: "keywords", section: "abstract", max: 6 }), ...tail],
  },
  "a count without a bound": { sentences: [count({ target: "keywords" }), ...tail] },
  "a bound that is a string": { sentences: [count({ target: "keywords", max: "250" }), ...tail] },
  "a bound that is a fraction": { sentences: [count({ target: "keywords", max: 2.5 }), ...tail] },
  "a negative bound": { sentences: [count({ target: "keywords", max: -1 }), ...tail] },
  "min above max": { sentences: [count({ target: "keywords", min: 6, max: 4 }), ...tail] },
  "an order of one": {
    sentences: [count({ target: "section_order", order: ["abstract"] }), ...tail],
  },
  "an order with a bound": {
    sentences: [
      count({ target: "section_order", order: ["abstract", "methods"], max: 2 }),
      ...tail,
    ],
  },
  "a list instead of an object": reply.sentences,
  "prose instead of labels": "Here are the rules: the abstract must not exceed 250 words.",
  nothing: null,
  "sentences given as text": { sentences: "The abstract must not exceed 250 words." },
  "an empty entry": { sentences: [null, ...tail] },
};

for (const [name, json] of Object.entries(BAD_REPLIES)) {
  test(`no rule from a reply with ${name}`, async () => {
    assert.deepEqual(candidatesFromLabels(sentences, json), {
      candidates: [],
      warnings: [GUIDELINE_WARNING],
    });
    const r = await extractGuidelineRules(text, { oneshot: answering(json), splitSentences });
    assert.deepEqual([r.candidates, r.warnings], [[], [GUIDELINE_WARNING]]);
  });
}

test("a failed request gives no rule and says why", async () => {
  const r = await extractGuidelineRules(text, {
    oneshot: async () => ({ ok: false, error: "timed out after 120000 ms" }),
    splitSentences,
  });
  assert.equal(r.candidates.length, 0);
  assert.match(r.warnings[0], /timed out/);
});

test("no sentences: the model is not asked", async () => {
  let calls = 0;
  const r = await extractGuidelineRules("   ", {
    oneshot: async () => (calls++, { ok: true, json: { sentences: [] } }),
    splitSentences,
  });
  assert.deepEqual([calls, r.candidates, r.warnings], [0, [], []]);
});

test("the splitter and the one-shot function must be passed in", async () => {
  await assert.rejects(extractGuidelineRules(text, { oneshot: answering(reply) }), TypeError);
  await assert.rejects(extractGuidelineRules(text, { splitSentences }), TypeError);
});

test("a long guideline is asked in batches; one bad batch leaves no rule at all", async () => {
  const many = Array.from({ length: 160 }, (_, i) => `Rule sentence number ${i + 1} stands here.`);
  const sizes = [];
  const good = async (req) => {
    const n = (req.prompt.match(/^\[s\d+\] /gm) ?? []).length;
    sizes.push(n);
    const labels = Array.from({ length: n }, (_, i) => ({ s: i + 1, type: "style" }));
    return { ok: true, json: { sentences: labels } };
  };
  const r = await extractGuidelineRules(many.join(" "), { oneshot: good, splitSentences });
  assert.deepEqual(sizes, [100, 60]);
  assert.equal(r.candidates.length, 160);
  assert.equal(r.candidates[159].text, many[159]);
  let call = 0;
  const second = async (req) => (++call === 2 ? { ok: true, json: { sentences: [] } } : good(req));
  const bad = await extractGuidelineRules(many.join(" "), { oneshot: second, splitSentences });
  assert.deepEqual([bad.candidates, bad.warnings], [[], [GUIDELINE_WARNING]]);
});

test("the invented guideline: twelve sentences, five counted rules, each one a sentence of the file", async () => {
  const guide = await readFile(new URL("./fixtures/fake_guideline.txt", import.meta.url), "utf8");
  assert.equal(splitSentences(guide).length, 12);
  const order = ["introduction", "methods", "results", "discussion"];
  const labels = {
    sentences: [
      { s: 1, type: "other" },
      { s: 2, type: "count", rule: { target: "section_words", section: "abstract", max: 250 } },
      { s: 3, type: "count", rule: { target: "keywords", min: 4, max: 6 } },
      { s: 4, type: "count", rule: { target: "section_order", order } },
      { s: 5, type: "count", rule: { target: "figures", max: 8 } },
      { s: 6, type: "count", rule: { target: "tables", max: 3 } },
      { s: 7, type: "count", rule: { target: "section_words", section: "introduction", max: 600 } },
      { s: 8, type: "style" },
      { s: 9, type: "style" },
      { s: 10, type: "other" },
      // The sentence says "six" in letters, and it is not about figures.
      { s: 11, type: "count", rule: { target: "figures", max: 6 } },
      { s: 12, type: "other" },
    ],
  };
  const r = await extractGuidelineRules(guide, { oneshot: answering(labels), splitSentences });
  assert.equal(r.candidates.length, 9);
  // s7 says "under 600 words": a strict limit is not turned into "at most 600".
  assert.equal(r.candidates.filter((c) => c.check.type === "count").length, 5);
  assert.equal(r.warnings.length, 2);
  assert.match(r.warnings[0], /s7\b.*upper limit/);
  assert.match(r.warnings[1], /s11\b/);
  for (const c of r.candidates) assert.ok(guide.includes(c.text));
});

test("htmlToText drops tags, scripts and styles and keeps block breaks", () => {
  assert.equal(htmlToText("<p>A <b>b</b>.</p><script>x</script>"), "A b.");
  assert.equal(
    htmlToText(
      "<style>p{color:red}</style><h1>Guide</h1><p>One&nbsp;&amp; two.</p><ul><li>Item &lt;1&gt;</li></ul><!-- c -->",
    ),
    "Guide\nOne & two.\nItem <1>",
  );
  assert.equal(htmlToText("<SCRIPT type='x'>var a = '<p>no</p>';</SCRIPT>Yes"), "Yes");
});

test("a guideline URL is fetched once with the injected fetch; html becomes text", async () => {
  const calls = [];
  const page =
    (body, type = "text/html; charset=utf-8", status = 200) =>
    async (url) => {
      calls.push(url);
      return {
        ok: status === 200,
        status,
        headers: { get: (h) => (h.toLowerCase() === "content-type" ? type : null) },
        text: async () => body,
      };
    };
  const get = (url, fetchFn, more = {}) => fetchGuidelineText(url, { fetchFn, ...more });
  const base = "https://journal.example";
  assert.deepEqual(await get(`${base}/guide`, page("<p>Use SI units.</p>")), {
    ok: true,
    text: "Use SI units.",
  });
  assert.deepEqual(calls, [`${base}/guide`]);
  const plain = await get(`${base}/guide.txt`, page("<Use> SI units.", "text/plain"));
  assert.equal(plain.text, "<Use> SI units.");
  assert.equal((await get(`${base}/x`, page("", "text/html", 404))).ok, false);
  const big = await get(`${base}/x`, page("a".repeat(11), "text/plain"), { maxBytes: 10 });
  assert.equal(big.ok, false);
  const n = calls.length;
  for (const url of ["file:///etc/hostname", "ftp://journal.example/x", "not a url"]) {
    assert.equal((await get(url, page("x"))).ok, false);
  }
  assert.equal(calls.length, n);
  const down = await get(`${base}/x`, async () => {
    throw new Error("getaddrinfo ENOTFOUND");
  });
  assert.deepEqual([down.ok, /ENOTFOUND/.test(down.error)], [false, true]);
});
