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

// A server that sends its headers and then stalls the body (Codex review of
// stage 2.5): the timeout covers the body too, for a good and an error status.
test("the timeout runs until the body is read: a stalled body is a timeout, not a success", async () => {
  const r = await stalledBody("chat");
  for (const name of ["good", "bad"]) {
    assert.deepEqual(r[name].result, { ok: false, error: "timed out after 70 ms" }, name);
    assert.ok(r[name].ms < 350, `${name}: returned after ${r[name].ms} ms, before the body came`);
  }
  // The same server answers in time when the limit allows it.
  assert.deepEqual([r.patient.result.ok, r.patient.result.json], [true, { ok: true }]);
  assert.ok(r.patient.ms >= 350, "that answer waited for the late body");
});

// The real requests run in a process of their own (tests/helpers/
// stalled-body.mjs says why) and print what they got.
async function stalledBody(scenario) {
  const { execFile } = await import("node:child_process");
  const { fileURLToPath } = await import("node:url");
  const helper = fileURLToPath(new URL("./helpers/stalled-body.mjs", import.meta.url));
  const stdout = await new Promise((resolve, reject) =>
    execFile(process.execPath, [helper, scenario], { timeout: 30_000 }, (err, out, errText) =>
      err ? reject(new Error(`${err.message}\n${errText}`)) : resolve(out),
    ),
  );
  return JSON.parse(stdout);
}

// A thinking model on Ollama (qwen3.5:9b, first run against Word): the
// thinking went to message.reasoning, used up the output and left content
// empty with finish_reason "length". That is not "reply was not JSON".
test("an answer cut off before any content names the cause", async () => {
  const cut = {
    choices: [{ finish_reason: "length", message: { content: "", reasoning: "..." } }],
  };
  const r = await chatOneshot(REQ, { fetchFn: reply(200, cut) });
  assert.deepEqual(r, {
    ok: false,
    error:
      "the model stopped before it wrote an answer (finish_reason length; a thinking model can use its whole output on thinking)",
  });
  // An answer that is there is judged as before, whatever the finish reason.
  const full = { choices: [{ finish_reason: "length", message: { content: '{"a":1}' } }] };
  assert.deepEqual((await chatOneshot(REQ, { fetchFn: reply(200, full) })).json, { a: 1 });
});

// What a provider wants in every request (body) and in a request for JSON
// (jsonBody) comes from its preset; without them the request is as before.
test("body fields of the preset are sent; those for JSON only when JSON is expected", async () => {
  const f = reply(200, { choices: [{ message: { content: "{}" } }] });
  const extra = {
    body: { reasoning_effort: "none" },
    jsonBody: { response_format: { type: "json_object" } },
  };
  await chatOneshot({ ...REQ, ...extra }, { fetchFn: f });
  assert.equal(reply.seen.body.reasoning_effort, "none");
  assert.deepEqual(reply.seen.body.response_format, { type: "json_object" });
  await chatOneshot({ ...REQ, ...extra, expectJson: false }, { fetchFn: f });
  assert.equal(reply.seen.body.reasoning_effort, "none");
  assert.equal("response_format" in reply.seen.body, false);
  await chatOneshot(REQ, { fetchFn: f });
  assert.deepEqual(Object.keys(reply.seen.body).sort(), [
    "messages",
    "model",
    "stream",
    "temperature",
  ]);
});

// A reasoning model refuses a temperature (HTTP 400 naming the parameter).
// The request is then sent once more without it; any other 400 stays a
// failure. A local server stands in for the provider (in the helper process):
// nothing is sent out.
test("a 400 about temperature is retried once without it; another 400 is not", async () => {
  const r = await stalledBody("temperature");
  assert.deepEqual([r.picky.result.ok, r.picky.result.json], [true, { a: 1 }]);
  assert.deepEqual(r.picky.temperatures, [0, "none"]);
  assert.equal(r.picky.same_messages, true, "the same request otherwise");

  assert.equal(r.never.result.ok, false);
  assert.match(r.never.result.error, /^HTTP 400: Unsupported value: 'temperature'/);
  assert.deepEqual(r.never.temperatures, [0, "none"], "once more, not again and again");

  assert.deepEqual(r.other.result, { ok: false, error: "HTTP 400: model not found" });
  assert.deepEqual(r.other.temperatures, [0], "no retry for another reason");
});

const ANSWER = { choices: [{ message: { content: '{"a":1}' } }] };

test("the temperature is the preset's (body): 0 unless it says otherwise, none when it says null", async () => {
  const f = reply(200, ANSWER);
  await chatOneshot(REQ, { fetchFn: f });
  assert.equal(reply.seen.body.temperature, 0);
  await chatOneshot({ ...REQ, body: { temperature: 0.2 } }, { fetchFn: f });
  assert.equal(reply.seen.body.temperature, 0.2);
  await chatOneshot({ ...REQ, body: { temperature: null } }, { fetchFn: f });
  assert.equal("temperature" in reply.seen.body, false);
});
