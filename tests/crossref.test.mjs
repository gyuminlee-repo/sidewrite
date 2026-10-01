import { test } from "node:test";
import assert from "node:assert/strict";
import {
  referenceEntries,
  extractDoi,
  titleShare,
  checkReferences,
  CROSSREF_NOTICE,
  CROSSREF_OFF,
  CROSSREF_NEEDS_MAILTO,
} from "../daemon/crossref.mjs";

const P = (index, text) => ({ index, text });
// Invented manuscript tail: body text, the reference list, and a section after it.
const DOC = [
  P(0, "The secret body sentence must never leave this PC."),
  P(1, "References"),
  P(2, "Kim A, Lee B. 2019. Growth of invented cells on methane. J Invent Results 4:10-20. https://doi.org/10.1000/abc."),
  P(3, "Nam F. 2016. A paper that does not exist. J Invent Results 1:1-2. doi:10.1000/gone"),
  P(4, "Park C. 2020. Carbon flux in a made-up strain. J Invent Results 5:30-41. https://doi.org/10.1000/title"),
  P(5, "Choi D. 2018. Enzymes of an imaginary pathway. J Invent Results 3:1-9."),
  P(6, "Yoon E. 2021. Yields of a fictional reactor. J Invent Results 6:50-58. https://doi.org/10.1000/year"),
  P(7, ""),
  P(8, "Supplementary material"),
  P(9, "Another secret sentence after the list."),
];
const profile = { ref_section_head: "^references$", ref_section_end_heads: ["^supplementary"] };

const work = (title, family, year, DOI) => ({
  title: [title],
  author: [{ family, given: "X." }],
  issued: { "date-parts": [[year, 1, 1]] },
  DOI,
});
const WORKS = {
  "10.1000/abc": work("Growth of invented cells on methane", "Kim", 2019, "10.1000/abc"),
  "10.1000/title": work("Migration of fish in cold rivers", "Park", 2020, "10.1000/title"),
  "10.1000/year": work("Yields of a fictional reactor", "Yoon", 2017, "10.1000/year"),
  "10.1000/who": work("Growth of invented cells on methane", "Smith", 2019, "10.1000/who"),
};
const SEARCH = [work("Enzymes of an imaginary pathway", "Choi", 2018, "10.1000/found")];

// A stand-in for api.crossref.org: /works/<doi> and /works?query.bibliographic=.
function fakeCrossref({ search = SEARCH, script = {} } = {}) {
  const calls = [];
  const fetchFn = async (url) => {
    calls.push(url);
    const u = new URL(url);
    const scripted = script[calls.length];
    if (scripted) return scripted;
    const json = (status, body) => ({
      ok: status === 200,
      status,
      headers: { get: () => null },
      json: async () => body,
    });
    if (u.pathname === "/works") return json(200, { message: { items: search } });
    const doi = decodeURIComponent(u.pathname.replace(/^\/works\//, ""));
    return WORKS[doi] ? json(200, { message: WORKS[doi] }) : json(404, null);
  };
  return { fetchFn, calls };
}
function clock() {
  let t = 0;
  const waits = [];
  return {
    now: () => t,
    sleep: async (ms) => {
      waits.push(ms);
      t += ms;
    },
    waits,
    tick: (ms) => (t += ms),
  };
}
const run = (more = {}) => {
  const c = clock();
  return checkReferences({
    entries: referenceEntries(DOC, profile),
    mailto: "a@example.org",
    enabled: true,
    now: c.now,
    sleep: c.sleep,
    ...more,
  });
};

test("the entries are the paragraphs between the references head and the next end head", () => {
  assert.deepEqual(
    referenceEntries(DOC, profile).map((p) => p.index),
    [2, 3, 4, 5, 6],
  );
  // Without an end head the list runs to the end; without a profile the usual head is found.
  assert.deepEqual(
    referenceEntries(DOC).map((p) => p.index),
    [2, 3, 4, 5, 6, 8, 9],
  );
  assert.deepEqual(referenceEntries([P(0, "No list here.")], profile), []);
});

test("a DOI is read without the link prefix and without the closing punctuation", () => {
  assert.equal(extractDoi("... https://doi.org/10.1000/abc."), "10.1000/abc");
  assert.equal(extractDoi("doi:10.1000/gone"), "10.1000/gone");
  assert.equal(extractDoi("(DOI 10.12345/J.X-1(2)3);"), "10.12345/J.X-1(2)3");
  assert.equal(extractDoi("10.1000/a.b.c, 2020"), "10.1000/a.b.c");
  for (const none of ["pp. 10.1000 of the book", "10.10/short", "version 10.1.2", ""]) {
    assert.equal(extractDoi(none), null);
  }
});

test("titles are compared by their words, without case, accents or punctuation", () => {
  const entry = "Kim A. 2019. Growth of invented cells on methane. J Invent 4:1.";
  assert.equal(titleShare("Growth of Invented Cells on Methane", entry), 1);
  assert.equal(titleShare("Growth of <i>invented</i> cells: on methane!", entry), 1);
  assert.ok(titleShare("Migration of fish in cold rivers", entry) < 0.9);
  assert.equal(titleShare("Étude des cellules", "Etude des cellules. 2019."), 1);
  assert.equal(titleShare("", entry), 0);
});

test("five entries give match, not_found, title_mismatch, candidate and year_mismatch, in order", async () => {
  const { fetchFn, calls } = fakeCrossref();
  const r = await run({ fetchFn });
  assert.equal(r.ok, true);
  assert.deepEqual(
    r.rows.map((x) => [x.entry_index, x.verdict]),
    [
      [2, "match"],
      [3, "not_found"],
      [4, "title_mismatch"],
      [5, "candidate"],
      [6, "year_mismatch"],
    ],
  );
  assert.equal(r.rows[0].doi, "10.1000/abc");
  assert.match(r.rows[2].detail, /Migration of fish/);
  assert.equal(r.rows[3].doi, "10.1000/found");
  assert.match(r.rows[4].detail, /2017/);
  assert.equal(calls.length, 5);
});

test("every request goes to the polite pool, and only reference entries leave the PC", async () => {
  const { fetchFn, calls } = fakeCrossref();
  await run({ fetchFn });
  for (const url of calls) {
    assert.ok(url.startsWith("https://api.crossref.org/works"));
    assert.ok(url.includes("mailto=a%40example.org"), url);
    assert.ok(!/secret|Supplementary/i.test(decodeURIComponent(url)));
  }
  const search = new URL(calls[3]);
  assert.equal(search.searchParams.get("query.bibliographic"), DOC[5].text);
  assert.equal(search.searchParams.get("rows"), "1");
  assert.equal(new URL(calls[0]).pathname, "/works/10.1000%2Fabc");
});

test("without a contact e-mail, or switched off, nothing is requested", async () => {
  const { fetchFn, calls } = fakeCrossref();
  for (const mailto of ["", "   ", undefined, "not-an-address"]) {
    assert.deepEqual(await run({ fetchFn, mailto }), { ok: false, error: CROSSREF_NEEDS_MAILTO });
  }
  assert.equal(CROSSREF_NEEDS_MAILTO, "Enter a contact e-mail for the Crossref polite pool.");
  for (const enabled of [false, undefined, "true"]) {
    assert.deepEqual(await run({ fetchFn, enabled }), { ok: false, error: CROSSREF_OFF });
  }
  assert.equal(calls.length, 0);
});

test("requests are sequential and at most five a second", async () => {
  const c = clock();
  const starts = [];
  let open = 0;
  const { fetchFn } = fakeCrossref();
  const slow = async (url) => {
    assert.equal(open, 0);
    open++;
    starts.push(c.now());
    c.tick(30);
    const res = await fetchFn(url);
    open--;
    return res;
  };
  await run({ fetchFn: slow, now: c.now, sleep: c.sleep });
  assert.equal(starts.length, 5);
  for (let i = 1; i < starts.length; i++) assert.ok(starts[i] - starts[i - 1] >= 200);
});

test("a 429 waits as long as Retry-After says and asks again", async () => {
  const c = clock();
  const busy = { ok: false, status: 429, headers: { get: (h) => (h.toLowerCase() === "retry-after" ? "7" : null) }, json: async () => null };
  const { fetchFn, calls } = fakeCrossref({ script: { 1: busy } });
  const r = await run({ fetchFn, now: c.now, sleep: c.sleep });
  assert.equal(r.ok, true);
  assert.equal(r.rows[0].verdict, "match");
  assert.equal(calls.length, 6);
  assert.equal(calls[0], calls[1]);
  assert.ok(c.waits.includes(7000));
});

test("a service that stays busy stops the check and keeps the rows it has", async () => {
  const busy = { ok: false, status: 429, headers: { get: () => null }, json: async () => null };
  const { fetchFn, calls } = fakeCrossref({ script: { 2: busy, 3: busy, 4: busy } });
  const r = await run({ fetchFn });
  assert.equal(r.ok, false);
  assert.match(r.error, /too many requests/);
  assert.deepEqual(
    r.rows.map((x) => x.verdict),
    ["match"],
  );
  assert.equal(calls.length, 4);
});

test("a wrong first author, a failing server and an empty search each get their own verdict", async () => {
  const entries = [
    P(1, "Kim A. 2019. Growth of invented cells on methane. J Invent 4:10. doi:10.1000/who"),
    P(2, "Lim G. 2015. A title nobody indexed. J Invent 9:9."),
    P(3, "Seo H. 2014. Another title. doi:10.1000/down"),
  ];
  const down = { ok: false, status: 503, headers: { get: () => null }, json: async () => null };
  const { fetchFn } = fakeCrossref({ search: [], script: { 3: down } });
  const r = await run({ fetchFn, entries });
  assert.deepEqual(
    r.rows.map((x) => x.verdict),
    ["author_mismatch", "not_found", "error"],
  );
  assert.match(r.rows[0].detail, /Smith/);
  assert.match(r.rows[2].detail, /503/);
  const broken = await run({
    entries: entries.slice(0, 1),
    fetchFn: async () => {
      throw new Error("ENOTFOUND api.crossref.org");
    },
  });
  assert.deepEqual([broken.rows[0].verdict, /ENOTFOUND/.test(broken.rows[0].detail)], ["error", true]);
});

test("a search hit with another title is not offered as a candidate", async () => {
  const entries = [P(2, "Lim G. 2015. A title nobody indexed. J Invent 9:9.")];
  const r = await run({ fetchFn: fakeCrossref().fetchFn, entries });
  assert.equal(r.rows[0].verdict, "not_found");
});

test("the log has the status only: no address, no entry text", async () => {
  const lines = [];
  await run({ fetchFn: fakeCrossref().fetchFn, log: (l) => lines.push(String(l)) });
  assert.equal(lines.length, 5);
  for (const line of lines) {
    assert.match(line, /^\[crossref\] (?:work|search) status=\d+$/);
    assert.ok(!line.includes("example.org") && !line.includes("Kim"));
  }
});

test("the notice says what leaves this PC", () => {
  assert.match(CROSSREF_NOTICE, /sends each reference entry and your contact e-mail from this PC to Crossref/);
});
