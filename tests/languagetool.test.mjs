import { test } from "node:test";
import assert from "node:assert/strict";
import {
  chunkParagraphs,
  createLimiter,
  ltCheck,
  matchToEdit,
  LT_LINK,
  LT_LIMITS,
  LT_NOTICE,
  LT_OFF,
} from "../daemon/languagetool.mjs";

const para = (index, text) => ({ index, id: `p${index}`, text });
const bytes = (s) => Buffer.byteLength(s);

test("the public limits are the documented ones", () => {
  assert.deepEqual(LT_LIMITS, {
    requestsPerMinute: 20,
    bytesPerMinute: 75_000,
    bytesPerRequest: 20_000,
  });
});

test("chunks stay under 18 kB and keep every paragraph whole and in order", () => {
  const ps = Array.from({ length: 40 }, (_, i) => para(i, `${"word ".repeat(199)}end${i}.`));
  const { chunks, refused } = chunkParagraphs(ps);
  assert.deepEqual(refused, []);
  assert.ok(chunks.length > 1);
  for (const c of chunks) {
    assert.ok(bytes(c.text) <= 18_000);
    for (const part of c.parts) {
      assert.equal(c.text.slice(part.start, part.start + part.length), part.paragraph.text);
    }
  }
  assert.deepEqual(
    chunks.flatMap((c) => c.parts.map((p) => p.paragraph.index)),
    ps.map((p) => p.index),
  );
});

test("the size is counted in bytes, not characters", () => {
  const hangul = para(0, "가".repeat(5000)); // 15,000 bytes
  const { chunks } = chunkParagraphs([hangul, para(1, "나".repeat(2000))]);
  assert.equal(chunks.length, 2);
});

test("one paragraph over the limit is refused by index and not sent; empty paragraphs are left out", () => {
  const { chunks, refused } = chunkParagraphs([
    para(0, "Short one."),
    para(1, "x".repeat(18_001)),
    para(2, "   "),
    para(3, "Short two."),
  ]);
  assert.deepEqual(refused, [{ index: 1, bytes: 18_001 }]);
  assert.equal(chunks.length, 1);
  assert.equal(chunks[0].text, "Short one.\n\nShort two.");
});

test("limiter: 20 requests a minute, then it waits for the window to move", async () => {
  let t = 0;
  const waits = [];
  const lim = createLimiter({
    now: () => t,
    sleep: async (ms) => {
      waits.push(ms);
      t += ms;
    },
  });
  for (let i = 0; i < 20; i++) await lim.take(100);
  assert.deepEqual(waits, []);
  await lim.take(100);
  assert.deepEqual(waits, [60_000]);
});

test("limiter: 75 kB a minute, then it waits; a later request passes once old ones have left the window", async () => {
  let t = 0;
  const waits = [];
  const lim = createLimiter({
    now: () => t,
    sleep: async (ms) => {
      waits.push(ms);
      t += ms;
    },
  });
  for (let i = 0; i < 4; i++) {
    await lim.take(18_000);
    t += 10_000;
  }
  assert.deepEqual(waits, []);
  await lim.take(18_000); // 90 kB in the window: wait until the first request is a minute old
  assert.deepEqual(waits, [20_000]);
  t += 60_000;
  await lim.take(18_000);
  assert.equal(waits.length, 1);
});

// ---- ltCheck ----
function fakeFetch(answer) {
  const calls = [];
  const fetchFn = async (url, init) => {
    calls.push({ url, init });
    const a = typeof answer === "function" ? answer(calls.length, init) : answer;
    if (a instanceof Error) throw a;
    return { ok: a.status === 200, status: a.status, json: async () => a.body };
  };
  return { fetchFn, calls };
}
const noWait = () => createLimiter({ now: () => 0, sleep: async () => {} });
const two = [para(4, "This is an eror here."), para(7, "A apple fell.")];
const reply = {
  status: 200,
  body: {
    matches: [
      {
        message: "Possible spelling mistake found.",
        offset: 11,
        length: 4,
        replacements: [
          { value: "error" },
          { value: "errors" },
          { value: "eros" },
          { value: "euro" },
          { value: "era" },
          { value: "ear" },
        ],
        rule: { id: "MORFOLOGIK_RULE_EN_US" },
      },
      {
        message: "Use 'An' before a vowel sound.",
        offset: 23,
        length: 1,
        replacements: [{ value: "An" }],
        rule: { id: "EN_A_VS_AN" },
      },
      // A match across the paragraph break belongs to no paragraph.
      { message: "x", offset: 20, length: 4, replacements: [], rule: { id: "X" } },
    ],
  },
};

test("check posts form data to <url>/v2/check and maps offsets back to paragraphs", async () => {
  const { fetchFn, calls } = fakeFetch(reply);
  const r = await ltCheck({ paragraphs: two, enabled: true, fetchFn, limiter: noWait() });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, "https://api.languagetool.org/v2/check");
  assert.equal(calls[0].init.method, "POST");
  assert.equal(calls[0].init.headers["Content-Type"], "application/x-www-form-urlencoded");
  const form = new URLSearchParams(calls[0].init.body);
  assert.deepEqual([...form.keys()].sort(), ["language", "text"]);
  assert.equal(form.get("text"), "This is an eror here.\n\nA apple fell.");
  assert.equal(form.get("language"), "auto");
  assert.equal(r.ok, true);
  assert.equal(r.link, LT_LINK);
  assert.deepEqual(r.matches, [
    {
      paragraph: { index: 4, id: "p4" },
      offset: 11,
      length: 4,
      message: "Possible spelling mistake found.",
      replacements: ["error", "errors", "eros", "euro", "era"],
      rule_id: "MORFOLOGIK_RULE_EN_US",
    },
    {
      paragraph: { index: 7, id: "p7" },
      offset: 0,
      length: 1,
      message: "Use 'An' before a vowel sound.",
      replacements: ["An"],
      rule_id: "EN_A_VS_AN",
    },
  ]);
});

test("the language and the server come from the settings", async () => {
  const { fetchFn, calls } = fakeFetch({ status: 200, body: { matches: [] } });
  await ltCheck({
    paragraphs: two,
    enabled: true,
    fetchFn,
    limiter: noWait(),
    url: "http://127.0.0.1:8081/",
    language: "en-GB",
  });
  assert.equal(calls[0].url, "http://127.0.0.1:8081/v2/check");
  assert.equal(new URLSearchParams(calls[0].init.body).get("language"), "en-GB");
});

test("off: no request, and the message says the text would leave this PC", async () => {
  const { fetchFn, calls } = fakeFetch(reply);
  for (const enabled of [false, undefined, "true", 1]) {
    const r = await ltCheck({ paragraphs: two, enabled, fetchFn, limiter: noWait() });
    assert.deepEqual(r, { ok: false, error: LT_OFF });
  }
  assert.equal(calls.length, 0);
  assert.match(LT_OFF, /LanguageTool is off.*leaves this PC/);
});

test("a long document goes out chunk by chunk through the limiter", async () => {
  const ps = Array.from({ length: 40 }, (_, i) => para(i, `${"word ".repeat(199)}end${i}.`));
  const taken = [];
  const limiter = { take: async (n) => taken.push(n) };
  const { fetchFn, calls } = fakeFetch({ status: 200, body: { matches: [] } });
  const r = await ltCheck({ paragraphs: ps, enabled: true, fetchFn, limiter });
  assert.equal(r.ok, true);
  assert.equal(calls.length, chunkParagraphs(ps).chunks.length);
  assert.equal(taken.length, calls.length);
  for (const n of taken) assert.ok(n > 0 && n <= LT_LIMITS.bytesPerRequest);
  // What the limiter counts is the text that is sent.
  assert.deepEqual(
    taken,
    calls.map((c) => bytes(new URLSearchParams(c.init.body).get("text"))),
  );
});

test("a refused paragraph is reported and the rest is still checked", async () => {
  const { fetchFn, calls } = fakeFetch({ status: 200, body: { matches: [] } });
  const r = await ltCheck({
    paragraphs: [para(0, "Fine."), para(1, "x".repeat(19_000))],
    enabled: true,
    fetchFn,
    limiter: noWait(),
  });
  assert.equal(calls.length, 1);
  assert.deepEqual([r.ok, r.refused], [true, [{ index: 1, bytes: 19_000 }]]);
});

test("429 and other failures stop the check with a plain message", async () => {
  const check = (answer) =>
    ltCheck({
      paragraphs: two,
      enabled: true,
      fetchFn: fakeFetch(answer).fetchFn,
      limiter: noWait(),
    });
  assert.match((await check({ status: 429, body: {} })).error, /too many requests/);
  assert.match((await check({ status: 500, body: {} })).error, /answered 500/);
  assert.match((await check(new Error("ECONNRESET"))).error, /could not be reached: ECONNRESET/);
  assert.match((await check({ status: 200, body: { nothing: 1 } })).error, /unexpected answer/);
});

test("the log has counts only, never the text", async () => {
  const lines = [];
  await ltCheck({
    paragraphs: two,
    enabled: true,
    fetchFn: fakeFetch(reply).fetchFn,
    limiter: noWait(),
    log: (l) => lines.push(String(l)),
  });
  assert.equal(lines.length, 1);
  assert.match(lines[0], /^\[languagetool\] check paragraphs=2 bytes=\d+ status=200 matches=2$/);
  assert.ok(!lines.join("\n").includes("eror") && !lines.join("\n").includes("apple"));
});

// ---- matchToEdit ----
test("a match becomes a find/replace inside its paragraph", () => {
  const e = matchToEdit({
    paragraph: { index: 4, id: "p4", text: "This is an eror here." },
    offset: 11,
    length: 4,
    replacements: ["error"],
  });
  assert.deepEqual(e, { find: "eror", replace: "error", paragraph_hint: "p4", paragraph_index: 4 });
});

test("when the words occur twice the find grows by neighbouring words until it is unique", () => {
  const paragraph = { index: 2, id: "p2", text: "A apple fell and A apple rolled." };
  const first = matchToEdit({ paragraph, offset: 0, length: 1, replacements: ["An"] });
  assert.deepEqual([first.find, first.replace], ["A apple fell", "An apple fell"]);
  const second = matchToEdit({ paragraph, offset: 17, length: 1, replacements: ["An"] });
  assert.deepEqual([second.find, second.replace], ["A apple rolled.", "An apple rolled."]);
  for (const e of [first, second]) assert.equal(paragraph.text.split(e.find).length, 2);
});

test("the picked replacement is used; no replacement or a stale offset gives no edit", () => {
  const paragraph = { index: 4, id: "p4", text: "This is an eror here." };
  const m = { paragraph, offset: 11, length: 4, replacements: ["error", "errors"] };
  assert.equal(matchToEdit(m, 1).replace, "errors");
  assert.equal(matchToEdit({ ...m, replacements: [] }), null);
  assert.equal(matchToEdit(m, 5), null);
  assert.equal(matchToEdit({ ...m, offset: 30 }), null);
  // Repeating text grows up to the whole paragraph, which occurs once.
  const same = { index: 1, id: "p1", text: "aa aa aa" };
  const e = matchToEdit({ paragraph: same, offset: 3, length: 2, replacements: ["b"] });
  assert.deepEqual([e.find, e.replace], ["aa aa aa", "aa b aa"]);
});

test("the link and the notice name the service", () => {
  assert.equal(LT_LINK, "https://languagetool.org");
  assert.match(LT_NOTICE, /sends the checked text from this PC to LanguageTool/);
});
