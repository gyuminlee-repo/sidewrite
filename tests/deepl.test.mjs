import { test } from "node:test";
import assert from "node:assert/strict";
import { deeplTranslate, deeplBase, translationRoute, DEEPL_NOTICE } from "../daemon/deepl.mjs";

const KEY = "0000-test-key:fx";
const SECRET_TEXT = "The culture reached a plateau after four days.";

// A fetch that records every call and answers from a list.
function fakeFetch(...answers) {
  const calls = [];
  const fetchFn = async (url, init) => {
    calls.push({ url, init });
    const a = answers[Math.min(calls.length - 1, answers.length - 1)];
    if (a instanceof Error) throw a;
    return {
      ok: a.status >= 200 && a.status < 300,
      status: a.status,
      json: async () => a.body,
      text: async () => JSON.stringify(a.body ?? ""),
    };
  };
  return { fetchFn, calls };
}
const ok = {
  status: 200,
  body: {
    translations: [{ detected_source_language: "EN", text: "배양은 나흘 뒤 정체기에 이르렀다." }],
  },
};
const ask = (more = {}) =>
  deeplTranslate({ text: SECRET_TEXT, target: "KO", key: KEY, enabled: true, ...more });

test("a free key posts to the free endpoint with the auth header and a JSON body", async () => {
  const { fetchFn, calls } = fakeFetch(ok);
  const r = await ask({ fetchFn });
  assert.deepEqual(r, { ok: true, text: "배양은 나흘 뒤 정체기에 이르렀다.", route: "deepl" });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, "https://api-free.deepl.com/v2/translate");
  assert.equal(calls[0].init.method, "POST");
  assert.deepEqual(calls[0].init.headers, {
    Authorization: `DeepL-Auth-Key ${KEY}`,
    "Content-Type": "application/json",
  });
  assert.deepEqual(JSON.parse(calls[0].init.body), { text: [SECRET_TEXT], target_lang: "KO" });
});

test("the endpoint follows the key: ':fx' is the free plan, anything else the other host", () => {
  assert.equal(deeplBase("abc:fx"), "https://api-free.deepl.com");
  assert.equal(deeplBase("abc"), "https://api.deepl.com");
  assert.equal(deeplBase("abc:fx", "http://127.0.0.1:9"), "http://127.0.0.1:9");
});

test("the target language is sent in upper case; a bad one is refused before any request", async () => {
  const { fetchFn, calls } = fakeFetch(ok);
  await ask({ fetchFn, target: "pt-br" });
  assert.equal(JSON.parse(calls[0].init.body).target_lang, "PT-BR");
  for (const target of ["", "korean!", null, "K"]) {
    assert.equal((await ask({ fetchFn, target })).ok, false);
  }
  assert.equal(calls.length, 1);
});

test("off, without a key or without text: no request", async () => {
  const { fetchFn, calls } = fakeFetch(ok);
  const off = await ask({ fetchFn, enabled: false });
  assert.deepEqual([off.ok, /DeepL is off/.test(off.error)], [false, true]);
  assert.equal(
    (await deeplTranslate({ text: SECRET_TEXT, target: "KO", key: KEY, fetchFn })).ok,
    false,
  );
  assert.equal((await ask({ fetchFn, key: "" })).ok, false);
  assert.equal((await ask({ fetchFn, text: "   " })).error, "Select some text first.");
  assert.equal(calls.length, 0);
});

test("403, 429 and 456 get plain words; another status shows what the service said", async () => {
  const say = async (status, body) =>
    (await ask({ fetchFn: fakeFetch({ status, body }).fetchFn })).error;
  assert.match(await say(403, {}), /key rejected/);
  assert.match(await say(429, {}), /too many requests/);
  assert.match(await say(456, {}), /quota exceeded/);
  assert.match(
    await say(400, { message: "Value for 'target_lang' not supported." }),
    /400.*not supported/,
  );
  assert.match(await say(403, { message: "Wrong endpoint." }), /key rejected.*Wrong endpoint/);
  assert.match(await say(200, { translations: [] }), /no translation/);
  const down = await ask({ fetchFn: fakeFetch(new Error("connect ECONNREFUSED")).fetchFn });
  assert.deepEqual([down.ok, /ECONNREFUSED/.test(down.error)], [false, true]);
});

test("a failure is an answer, not a second request and not another route", async () => {
  const { fetchFn, calls } = fakeFetch({ status: 429, body: {} });
  const r = await ask({ fetchFn });
  assert.equal(r.ok, false);
  assert.equal(r.route, "deepl");
  assert.equal(calls.length, 1);
});

test("the route is DeepL only when it is switched on and a key is stored", () => {
  const on = { external: { deepl: { enabled: true } } };
  assert.equal(translationRoute({ settings: on, key: KEY }), "deepl");
  assert.equal(translationRoute({ settings: on, key: "" }), "oneshot");
  assert.equal(
    translationRoute({ settings: { external: { deepl: { enabled: false } } }, key: KEY }),
    "oneshot",
  );
  assert.equal(translationRoute({ settings: {}, key: KEY }), "oneshot");
  assert.equal(translationRoute({ key: KEY }), "oneshot");
  assert.equal(
    translationRoute({ settings: { external: { deepl: { enabled: "yes" } } }, key: KEY }),
    "oneshot",
  );
});

test("the log line has counts only: no key, no text, no translation", async () => {
  const lines = [];
  const log = (line) => lines.push(String(line));
  await ask({ fetchFn: fakeFetch(ok).fetchFn, log });
  await ask({ fetchFn: fakeFetch({ status: 403, body: { message: "no" } }).fetchFn, log });
  await ask({ fetchFn: fakeFetch(new Error(`bad header DeepL-Auth-Key ${KEY}`)).fetchFn, log });
  assert.equal(lines.length, 3);
  assert.match(lines[0], /^\[deepl\] translate chars=46 target=KO status=200$/);
  for (const line of lines) {
    assert.ok(!line.includes(KEY) && !line.includes("0000-test-key"));
    assert.ok(!line.includes("culture") && !line.includes("배양"));
  }
});

test("an error message never carries the key", async () => {
  const r = await ask({
    fetchFn: fakeFetch(new Error(`bad header DeepL-Auth-Key ${KEY}`)).fetchFn,
  });
  assert.ok(!r.error.includes(KEY));
});

test("the notice says the text leaves this PC and names the service", () => {
  assert.match(DEEPL_NOTICE, /sends the selected text from this PC to DeepL/);
  assert.match(DEEPL_NOTICE, /terms/);
});
