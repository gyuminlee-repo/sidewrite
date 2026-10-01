// The requests of the Tools tab (daemon/aux-service.mjs), with a fake model
// and a fake Python runner. No Word, no port.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createAuxService, AUX_JOBS } from "../daemon/aux-service.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };

// A service on a manuscript in a folder of its own. call(type, msg) answers
// what the pane gets: the reply of a request, or the result of a job.
async function setup({ saved = true, ...deps } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "aux-"));
  const doc = join(dir, "paper.docx");
  await writeFile(doc, "docx");
  const events = [];
  const waiting = new Map();
  const service = createAuxService({
    getContext: () => ({ activeDoc: saved ? doc : null }),
    sendEvent: (key, event) => {
      events.push(event);
      waiting.get(event.job_id)?.(event.result);
    },
    log: () => {},
    env: ENV,
    ...deps,
  });
  async function call(type, msg = {}) {
    let reply;
    await service.handlers[type]({ request_id: "q", ...msg }, (r) => (reply = r), "k", "word");
    if (!reply.accepted) return reply;
    return new Promise((resolve) => {
      const done = events.find((e) => e.job_id === reply.job_id);
      if (done) resolve(done.result);
      else waiting.set(reply.job_id, resolve);
    });
  }
  return { dir, doc, service, call, events, done: () => rm(dir, { recursive: true }) };
}

test("the tools service is on the shared request wrapper: replies are <type>_result", async () => {
  const t = await setup();
  try {
    let reply;
    await t.service.handlers.tools_state({ request_id: "q7" }, (r) => (reply = r), "k", "word");
    assert.equal(reply.type, "tools_state_result");
    assert.equal(reply.request_id, "q7");
    assert.equal(reply.ok, true);
    assert.equal(reply.doc_name, "paper.docx");
    assert.equal(reply.profile_name, "generic");
    assert.deepEqual(reply.policy_packs, { review_scope: true });
    for (const type of AUX_JOBS) assert.equal(typeof t.service.handlers[type], "function");
  } finally {
    await t.done();
  }
});

test("a document that is not saved gets an error reply that says to save it", async () => {
  const t = await setup({ saved: false });
  try {
    const r = await t.call("tools_state");
    assert.equal(r.ok, false);
    assert.match(r.error, /Save the document to a local folder first/);
  } finally {
    await t.done();
  }
});
