import { test } from "node:test";
import assert from "node:assert/strict";
import { chatOneshot } from "../daemon/providers/openai-compat.mjs";

const reply = (status, body) => async (url, init) => {
  reply.seen = { url, init, body: JSON.parse(init.body) };
  return {
    ok: status < 300,
    status,
    json: async () => body,
    text: async () => JSON.stringify(body),
  };
};
const REQ = { baseUrl: "https://x/v1", apiKey: "k", model: "m", system: "", prompt: "" };

test("posts system and user, parses JSON, strips think blocks", async () => {
  const f = reply(200, { choices: [{ message: { content: '<think>hm</think>{"labels":[]}' } }] });
  const r = await chatOneshot(
    { baseUrl: "http://localhost:11434/v1", model: "qwen3.5:9b", system: "S", prompt: "P" },
    { fetchFn: f },
  );
  assert.deepEqual([r.ok, r.json], [true, { labels: [] }]);
  assert.equal(reply.seen.url, "http://localhost:11434/v1/chat/completions");
  assert.equal(reply.seen.init.method, "POST");
  assert.deepEqual(reply.seen.body.messages, [
    { role: "system", content: "S" },
    { role: "user", content: "P" },
  ]);
  assert.equal(reply.seen.body.model, "qwen3.5:9b");
  assert.equal(reply.seen.body.temperature, 0);
  assert.equal(reply.seen.body.stream, false);
  assert.equal("authorization" in (reply.seen.init.headers ?? {}), false);
  assert.ok(reply.seen.init.signal, "the request has a timeout signal");
});

test("bearer key, 429 is a limit, other errors keep status", async () => {
  const lim = await chatOneshot(REQ, {
    fetchFn: reply(429, { error: { message: "slow down" } }),
  });
  assert.equal(reply.seen.init.headers.authorization, "Bearer k");
  assert.equal(lim.ok, false);
  assert.equal(lim.limited.kind, "rate_limit");
  assert.match(lim.error, /HTTP 429.*slow down/);
  const bad = await chatOneshot(REQ, { fetchFn: reply(500, { error: "x" }) });
  assert.match(bad.error, /HTTP 500/);
  assert.equal(bad.limited ?? null, null);
});

test("non-JSON reply keeps the text, text mode skips parsing", async () => {
  const f = reply(200, { choices: [{ message: { content: "sorry" } }] });
  const a = await chatOneshot({ baseUrl: "u", model: "m", system: "", prompt: "" }, { fetchFn: f });
  assert.deepEqual([a.ok, a.error, a.text], [false, "reply was not JSON", "sorry"]);
  const b = await chatOneshot(
    { baseUrl: "u", model: "m", system: "", prompt: "", expectJson: false },
    { fetchFn: f },
  );
  assert.deepEqual([b.ok, b.text], [true, "sorry"]);
});

test("never throws: a network failure, a timeout and an empty reply are error results", async () => {
  const net = await chatOneshot(REQ, {
    fetchFn: async () => {
      throw Object.assign(new TypeError("fetch failed"), { cause: { code: "ECONNRESET" } });
    },
  });
  assert.deepEqual([net.ok, net.error], [false, "Connection failed (ECONNRESET)"]);
  const slow = await chatOneshot(
    { ...REQ, timeoutMs: 20 },
    {
      fetchFn: (url, init) =>
        new Promise((_, rej) =>
          init.signal.addEventListener("abort", () => rej(init.signal.reason)),
        ),
    },
  );
  assert.deepEqual([slow.ok, slow.error], [false, "timed out after 20 ms"]);
  const empty = await chatOneshot(REQ, { fetchFn: reply(200, { choices: [] }) });
  assert.deepEqual([empty.ok, empty.error], [false, "the reply had no message"]);
  const noModel = await chatOneshot({ ...REQ, model: "" }, { fetchFn: reply(200, {}) });
  assert.match(noModel.error, /Choose a model/);
});

test("the key is in the Authorization header only: not in the URL, the body or an error", async () => {
  const SECRET = "sk-secret-0123456789";
  const r = await chatOneshot(
    { ...REQ, apiKey: SECRET },
    { fetchFn: reply(401, { error: { message: `Incorrect API key provided: ${SECRET}` } }) },
  );
  assert.equal(r.ok, false);
  assert.equal(JSON.stringify(r).includes(SECRET), false, "an echoed key is cut from the error");
  assert.equal(reply.seen.url.includes(SECRET), false);
  assert.equal(reply.seen.init.body.includes(SECRET), false);
});

test("a trailing slash of the base URL and a think block left open are handled", async () => {
  const f = reply(200, { choices: [{ message: { content: '{"a":1}<think>tail' } }] });
  const r = await chatOneshot({ ...REQ, baseUrl: "https://x/v1/" }, { fetchFn: f });
  assert.equal(reply.seen.url, "https://x/v1/chat/completions");
  assert.deepEqual(r.json, { a: 1 });
});
