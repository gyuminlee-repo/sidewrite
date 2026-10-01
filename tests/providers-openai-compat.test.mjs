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
// failure. A local server stands in for the provider: nothing is sent out.
async function localProvider(answer) {
  const http = await import("node:http");
  const bodies = [];
  const server = http.createServer((req, res) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => {
      const body = JSON.parse(Buffer.concat(chunks).toString("utf8"));
      bodies.push(body);
      const [status, json] = answer(body, bodies.length);
      res.writeHead(status, { "content-type": "application/json", connection: "close" });
      res.end(JSON.stringify(json));
    });
  });
  await new Promise((r) => server.listen(0, "127.0.0.1", r));
  return {
    bodies,
    baseUrl: `http://127.0.0.1:${server.address().port}/v1`,
    close: () =>
      new Promise((r) => {
        server.close(r);
        server.closeAllConnections();
      }),
  };
}
const NO_TEMPERATURE = {
  error: {
    message:
      "Unsupported value: 'temperature' does not support 0 with this model. Only the default (1) value is supported.",
    param: "temperature",
  },
};
const ANSWER = { choices: [{ message: { content: '{"a":1}' } }] };

test("a 400 about temperature is retried once without it; another 400 is not", async () => {
  const picky = await localProvider((body) =>
    "temperature" in body ? [400, NO_TEMPERATURE] : [200, ANSWER],
  );
  const never = await localProvider(() => [400, NO_TEMPERATURE]);
  const other = await localProvider(() => [400, { error: { message: "model not found" } }]);
  try {
    const req = { model: "o-reasoner", system: "S", prompt: "P" };
    const r = await chatOneshot({ ...req, baseUrl: picky.baseUrl });
    assert.deepEqual([r.ok, r.json], [true, { a: 1 }]);
    assert.deepEqual(
      picky.bodies.map((b) => b.temperature),
      [0, undefined],
    );
    assert.deepEqual(
      picky.bodies[1].messages,
      picky.bodies[0].messages,
      "the same request otherwise",
    );

    const still = await chatOneshot({ ...req, baseUrl: never.baseUrl });
    assert.equal(still.ok, false);
    assert.match(still.error, /^HTTP 400: Unsupported value: 'temperature'/);
    assert.equal(never.bodies.length, 2, "once more, not again and again");

    const bad = await chatOneshot({ ...req, baseUrl: other.baseUrl });
    assert.deepEqual(bad, { ok: false, error: "HTTP 400: model not found" });
    assert.equal(other.bodies.length, 1, "no retry for another reason");
  } finally {
    await Promise.all([picky.close(), never.close(), other.close()]);
  }
});

test("the temperature is the preset's (body): 0 unless it says otherwise, none when it says null", async () => {
  const f = reply(200, ANSWER);
  await chatOneshot(REQ, { fetchFn: f });
  assert.equal(reply.seen.body.temperature, 0);
  await chatOneshot({ ...REQ, body: { temperature: 0.2 } }, { fetchFn: f });
  assert.equal(reply.seen.body.temperature, 0.2);
  await chatOneshot({ ...REQ, body: { temperature: null } }, { fetchFn: f });
  assert.equal("temperature" in reply.seen.body, false);
});
