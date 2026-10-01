import { test } from "node:test";
import assert from "node:assert/strict";
import { buildOneshotOptions, parseJsonReply, runOneshot } from "../daemon/oneshot.mjs";

function fakeQuery(messages, seen = {}) {
  return ({ prompt, options }) => {
    seen.prompt = prompt;
    seen.options = options;
    return (async function* () {
      for (const m of messages) {
        if (m instanceof Error) throw m;
        yield m;
      }
    })();
  };
}
const ok = (text) => ({ type: "result", subtype: "success", is_error: false, result: text });

test("options: no tools, no MCP, no settings, not persisted, no resume, haiku", () => {
  const o = buildOneshotOptions({ system: "S" });
  assert.deepEqual(o.tools, []);
  assert.deepEqual(o.mcpServers, {});
  assert.deepEqual(o.settingSources, []);
  assert.equal(o.strictMcpConfig, true);
  assert.equal(o.persistSession, false);
  assert.equal(o.maxTurns, 1);
  assert.equal(o.model, "haiku");
  assert.equal(o.systemPrompt, "S");
  assert.equal("resume" in o, false);
  assert.equal("canUseTool" in o, false);
});

test("parseJsonReply: bare, fenced, prose-wrapped, invalid", () => {
  assert.deepEqual(parseJsonReply('{"a":1}'), { a: 1 });
  assert.deepEqual(parseJsonReply('```json\n{"a":2}\n```'), { a: 2 });
  assert.deepEqual(parseJsonReply('Here you go: {"a":3} done'), { a: 3 });
  assert.equal(parseJsonReply("no json here"), undefined);
  assert.equal(parseJsonReply(""), undefined);
});

test("runOneshot returns parsed JSON and passes the prompt and options through", async () => {
  const seen = {};
  const r = await runOneshot(
    { system: "sys", prompt: "split this" },
    { queryFn: fakeQuery([{ type: "system", subtype: "init" }, ok('{"comments":[]}')], seen) },
  );
  assert.deepEqual([r.ok, r.json], [true, { comments: [] }]);
  assert.equal(seen.prompt, "split this");
  assert.deepEqual(seen.options.tools, []);
  assert.ok(seen.options.abortController instanceof AbortController);
});

test("runOneshot: text mode skips JSON parsing", async () => {
  const r = await runOneshot(
    { system: "s", prompt: "p", expectJson: false },
    { queryFn: fakeQuery([ok("plain")]) },
  );
  assert.deepEqual([r.ok, r.text], [true, "plain"]);
});

test("runOneshot: non-JSON reply is an error that keeps the text", async () => {
  const r = await runOneshot({ system: "s", prompt: "p" }, { queryFn: fakeQuery([ok("sorry")]) });
  assert.deepEqual([r.ok, r.error, r.text], [false, "reply was not JSON", "sorry"]);
});

test("runOneshot: usage limit wins over a success result", async () => {
  const lim = {
    type: "assistant",
    error: "rate_limit",
    message: {
      model: "<synthetic>",
      content: [{ type: "text", text: "You've hit your session limit" }],
    },
  };
  const r = await runOneshot(
    { system: "s", prompt: "p" },
    { queryFn: fakeQuery([lim, ok("You've hit your session limit")]) },
  );
  assert.equal(r.ok, false);
  assert.match(r.error, /usage limit/);
  assert.equal(r.limited.limited, true);
});

test("runOneshot: error subtype, missing result, thrown error", async () => {
  let r = await runOneshot(
    { system: "s", prompt: "p" },
    {
      queryFn: fakeQuery([
        { type: "result", subtype: "error_max_turns", is_error: true, errors: ["x"] },
      ]),
    },
  );
  assert.match(r.error, /error_max_turns/);
  r = await runOneshot({ system: "s", prompt: "p" }, { queryFn: fakeQuery([]) });
  assert.equal(r.error, "no result message");
  r = await runOneshot(
    { system: "s", prompt: "p" },
    { queryFn: fakeQuery([new Error("spawn failed")]) },
  );
  assert.deepEqual([r.ok, r.error], [false, "spawn failed"]);
});

test("runOneshot: timeout aborts and reports", async () => {
  const hang = ({ options }) =>
    (async function* () {
      await new Promise((resolve, reject) =>
        options.abortController.signal.addEventListener("abort", () => {
          const e = new Error("aborted");
          e.name = "AbortError";
          reject(e);
        }),
      );
    })();
  const r = await runOneshot({ system: "s", prompt: "p", timeoutMs: 20 }, { queryFn: hang });
  assert.deepEqual([r.ok, r.error], [false, "timed out after 20 ms"]);
});

// Class audit (Codex review 13): replies that must never read as an answer.
test("runOneshot: a refusal, an empty reply and cut-off JSON are errors", async () => {
  const run = (messages) =>
    runOneshot({ system: "s", prompt: "p" }, { queryFn: fakeQuery(messages) });
  const refused = await run([
    { type: "system", subtype: "model_refusal_no_fallback", api_refusal_category: "x" },
    ok('{"violations":[]}'),
  ]);
  assert.equal(refused.ok, false);
  assert.match(refused.error, /declined/);
  assert.doesNotMatch(refused.error, /usage limit/);
  const stopped = await run([{ ...ok('{"violations":[]}'), stop_reason: "refusal" }]);
  assert.equal(stopped.ok, false);
  for (const text of [
    "",
    "   ",
    '{"violations":[{"paragraph_index":3,"quote":"a","reason":"r"}',
    '{"violations":[],"unsure":[{"paragraph_index":5,"reason":"x"}',
    '```json\n{"violations":[\n```',
    'Note {see below}: {"violations":[]}',
  ]) {
    const r = await run([ok(text)]);
    assert.deepEqual([r.ok, r.error], [false, "reply was not JSON"], JSON.stringify(text));
  }
});
