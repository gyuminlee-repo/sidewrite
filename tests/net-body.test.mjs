// Every module that asks a server: a connection that breaks after the
// headers, or a body that never comes, ends in an error result. Nothing is
// thrown and nothing waits for ever.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fetchStyle } from "../daemon/csl-style.mjs";
import { deeplTranslate } from "../daemon/deepl.mjs";
import { ltCheck } from "../daemon/languagetool.mjs";
import { checkReferences } from "../daemon/crossref.mjs";
import { fetchGuidelineText } from "../daemon/guideline-extract.mjs";
import { FETCH_TIMEOUT_MS } from "../daemon/fetch-timeout.mjs";

// A response whose headers arrived and whose body is read with `body()`.
const response = (body) => ({
  ok: true,
  status: 200,
  headers: { get: () => "text/html" },
  json: body,
  text: body,
  arrayBuffer: body,
});
const broken = async () => response(async () => Promise.reject(new Error("socket hang up")));
// The body never comes; it ends only when the caller's signal fires.
const signals = [];
const silent = async (url, init) => {
  signals.push(init?.signal instanceof AbortSignal);
  return response(
    () =>
      new Promise((_, reject) => {
        if (!init?.signal) return reject(new Error("no signal was passed to fetch"));
        // An open socket keeps the process alive while a real body is awaited; here a timer does.
        const socket = setInterval(() => {}, 1000);
        init.signal.addEventListener("abort", () => {
          clearInterval(socket);
          reject(init.signal.reason);
        });
      }),
  );
};

const cacheDir = () => mkdtemp(join(tmpdir(), "net-body-"));
const CALLS = {
  "csl-style": async (fetchFn, timeoutMs) =>
    fetchStyle({
      slug: "an-invented-style",
      enabled: true,
      fetchFn,
      cacheDir: await cacheDir(),
      timeoutMs,
    }),
  deepl: (fetchFn, timeoutMs) =>
    deeplTranslate({
      text: "An invented sentence.",
      target: "ko",
      key: "k:fx",
      enabled: true,
      fetchFn,
      timeoutMs,
    }),
  languagetool: (fetchFn, timeoutMs) =>
    ltCheck({
      paragraphs: [{ index: 0, text: "An invented sentence." }],
      enabled: true,
      fetchFn,
      limiter: { take: async () => {} },
      timeoutMs,
    }),
  crossref: (fetchFn, timeoutMs) =>
    checkReferences({
      entries: [{ index: 0, text: "Somebody A. An invented title. 2031. doi:10.1000/abc" }],
      mailto: "someone@example.org",
      enabled: true,
      fetchFn,
      sleep: async () => {},
      timeoutMs,
    }),
  guideline: (fetchFn, timeoutMs) =>
    fetchGuidelineText("https://example.org/guide", { fetchFn, timeoutMs }),
};
// Where the error text of a failed call is.
const errorOf = (r) => r.error ?? r.rows?.[0]?.detail ?? "";
const failed = (r) => r.ok === false || r.rows?.[0]?.verdict === "error";

for (const [name, call] of Object.entries(CALLS)) {
  test(`${name}: a connection that breaks after the headers is an error result, not an exception`, async () => {
    const r = await call(broken);
    assert.equal(failed(r), true, JSON.stringify(r));
  });

  test(`${name}: a body that never comes ends when the time is up`, async () => {
    const started = Date.now();
    signals.length = 0;
    const r = await call(silent, 40);
    assert.equal(failed(r), true, JSON.stringify(r));
    assert.deepEqual([...new Set(signals)], [true], "fetch was called without a time limit");
    assert.doesNotMatch(errorOf(r), /no signal was passed/);
    assert.ok(Date.now() - started < 5000);
  });
}

test("the time limit covers the whole exchange and has one default", () => {
  assert.ok(
    Number.isInteger(FETCH_TIMEOUT_MS) && FETCH_TIMEOUT_MS >= 10_000 && FETCH_TIMEOUT_MS <= 120_000,
  );
});
