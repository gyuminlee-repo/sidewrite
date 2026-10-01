// P3: translate the selected text or ask about it (daemon/aux-service.mjs).
// The model is a fake; the point is what reaches it and that nothing of the
// document is read or written.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createAuxService, AUX_JOBS, ASK_SYSTEM, MAX_SELECTION } from "../daemon/aux-service.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
const TEXT = "The cells were grown at 30 °C for two days.";

async function setup({ reply = { ok: true, text: " answer " }, model = "sonnet", run } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "sel-"));
  await writeFile(join(dir, "paper.docx"), "docx");
  const asked = [];
  const pins = [];
  const contexts = [];
  const events = [];
  const service = createAuxService({
    getContext: (key) => (contexts.push(key), { activeDoc: join(dir, "paper.docx") }),
    oneshotFor: async (key, options) => {
      pins.push({ key, options });
      if (run === null) return { run: null, model: null };
      return { model, run: async (req) => (asked.push(req), reply) };
    },
    sendEvent: (key, event) => events.push(event),
    log: () => {},
    env: ENV,
  });
  async function job(type, msg) {
    let first;
    await service.handlers[type]({ request_id: "q", ...msg }, (r) => (first = r), "k", "word");
    assert.equal(first.accepted, true);
    for (let i = 0; i < 50 && !events.length; i++) await new Promise((r) => setTimeout(r, 2));
    assert.equal(events[0].type, "tools_event");
    assert.equal(events[0].job_id, first.job_id);
    return events.shift().result;
  }
  return { dir, asked, pins, contexts, job, done: () => rm(dir, { recursive: true }) };
}

test("both selection requests are jobs of the tools service", () => {
  assert.ok(AUX_JOBS.has("selection_translate"));
  assert.ok(AUX_JOBS.has("selection_ask"));
});

test("translate: the selection goes to the model with the language, and the answer comes back", async () => {
  const t = await setup({ reply: { ok: true, text: " 번역문 " } });
  try {
    const r = await t.job("selection_translate", { text: `  ${TEXT} `, lang: "Korean" });
    assert.deepEqual(r, { ok: true, text: "번역문", lang: "Korean", model: "sonnet" });
    assert.equal(t.asked.length, 1);
    assert.match(t.asked[0].system, /Translate into Korean/);
    assert.equal(t.asked[0].prompt, TEXT);
    assert.equal(t.asked[0].expectJson, false);
    // The pane's provider is pinned once for the job, with the prose model.
    assert.deepEqual(t.pins, [{ key: "k", options: { claudeModel: "sonnet" } }]);
  } finally {
    await t.done();
  }
});

test("ask: the selection and the question go to the model in one prompt", async () => {
  const t = await setup({ reply: { ok: true, text: "At 30 °C." } });
  try {
    const r = await t.job("selection_ask", { text: TEXT, question: "At what temperature?" });
    assert.deepEqual(r, { ok: true, text: "At 30 °C.", model: "sonnet" });
    assert.equal(t.asked[0].system, ASK_SYSTEM);
    assert.equal(t.asked[0].prompt, `Text:\n${TEXT}\n\nQuestion: At what temperature?`);
    assert.equal(t.asked[0].expectJson, false);
  } finally {
    await t.done();
  }
});

test("neither request reads the document or writes a file next to it", async () => {
  const t = await setup();
  try {
    await t.job("selection_translate", { text: TEXT, lang: "Korean" });
    await t.job("selection_ask", { text: TEXT, question: "Why?" });
    assert.deepEqual(t.contexts, []);
    assert.deepEqual(await readdir(t.dir), ["paper.docx"]);
  } finally {
    await t.done();
  }
});

test("what is missing is an error and the model is not asked", async () => {
  const t = await setup();
  try {
    const cases = [
      ["selection_translate", { text: "  ", lang: "Korean" }, /Select some text/],
      ["selection_translate", { text: TEXT, lang: " " }, /Name the language/],
      ["selection_ask", { text: TEXT, question: "" }, /Type a question/],
      ["selection_ask", { text: "x".repeat(MAX_SELECTION + 1), question: "Why?" }, /Select less/],
    ];
    for (const [type, msg, error] of cases) {
      const r = await t.job(type, msg);
      assert.equal(r.ok, false);
      assert.match(r.error, error);
    }
    assert.equal(t.asked.length, 0);
  } finally {
    await t.done();
  }
});

test("a model failure, an empty answer and no model at all are errors with their reason", async () => {
  let t = await setup({ reply: { ok: false, error: "usage limit" } });
  try {
    const r = await t.job("selection_translate", { text: TEXT, lang: "Korean" });
    assert.deepEqual(r, { ok: false, error: "translation failed: usage limit" });
  } finally {
    await t.done();
  }
  t = await setup({ reply: { ok: true, text: "  " } });
  try {
    const r = await t.job("selection_ask", { text: TEXT, question: "Why?" });
    assert.deepEqual(r, { ok: false, error: "the answer failed: empty answer" });
  } finally {
    await t.done();
  }
  t = await setup({ run: null });
  try {
    const r = await t.job("selection_ask", { text: TEXT, question: "Why?" });
    assert.match(r.error, /No model is available for answering/);
  } finally {
    await t.done();
  }
});
