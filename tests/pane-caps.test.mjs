// A task pane says what it can do in its hello (caps). A tracked edit is
// sent only to a pane that compares the paragraph before the edit
// ("target_hash", tools-word.js toolTrackedEdit): a pane loaded before the
// daemon was updated would apply the edit unchecked, so it gets no tool call
// and the model is told to have the pane reloaded.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { WebSocket } from "ws";
import { createBridge } from "../daemon/bridge.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit, TARGET_HASH_CAP } from "../daemon/review-tools.mjs";

async function until(fn, what, ms = 3000) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await new Promise((r) => setTimeout(r, 10));
  }
}

test("the task pane names the capability the daemon asks for, and sends it in its hello", async () => {
  const { PANE_CAPS } = await import("../taskpane/shared/tools-word.js");
  assert.equal(TARGET_HASH_CAP, "target_hash");
  assert.ok(PANE_CAPS.includes(TARGET_HASH_CAP));
  const src = await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");
  const hello = /function sendHello\(\) \{[\s\S]*?\n\}/.exec(src)[0];
  assert.match(hello, /caps: PANE_CAPS,/);
});

test("bridge: the caps of a hello are in the pane's context; a hello without them has none", async () => {
  const keys = [];
  const bridge = createBridge({
    port: 0,
    token: "t",
    allowedOrigins: [],
    onHello: (key) => keys.push(key),
  });
  const socks = [];
  const hello = async (extra) => {
    const ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    socks.push(ws);
    await new Promise((r, j) => (ws.once("open", r), ws.once("error", j)));
    const n = keys.length;
    const msg = { type: "hello", token: "t", host: "word", active_doc: "/d/a.docx", ...extra };
    ws.send(JSON.stringify(msg));
    await until(() => keys.length > n, "hello");
    return { ws, key: keys.at(-1) };
  };
  try {
    await until(() => bridge.address(), "listening");
    const a = await hello({ pane_id: "a", caps: ["target_hash", 7, "other"] });
    assert.deepEqual(bridge.getContext(a.key).caps, ["target_hash", "other"]);
    // A later message of the pane does not take them away.
    a.ws.send(JSON.stringify({ type: "context_update", active_doc: "/d/b.docx" }));
    await until(() => bridge.getContext(a.key).activeDoc === "/d/b.docx", "context_update");
    assert.deepEqual(bridge.getContext(a.key).caps, ["target_hash", "other"]);
    // An older task pane: no caps. Not a list: none either.
    const b = await hello({ pane_id: "b" });
    assert.deepEqual(bridge.getContext(b.key).caps, []);
    const c = await hello({ pane_id: "c", caps: "target_hash" });
    assert.deepEqual(bridge.getContext(c.key).caps, []);
    // No pane at all.
    assert.deepEqual(bridge.getContext("word|nowhere").caps, []);
  } finally {
    for (const ws of socks) ws.terminate();
    await bridge.close();
  }
});

const R = { what: "w", why: "y", addresses: "a" };
const EDIT = { comment_ref: "R1.1", find: "the substrate", replace: "the medium", rationale: R };

test("applyTrackedEdit: a pane that does not compare the paragraph gets no call and no attempt", async () => {
  const seen = [];
  const deps = (can) => ({
    paneChecksTarget: () => can,
    getLedger: async () => [],
    onAttempt: async () => (seen.push("attempt"), { id: "R1.1-b1" }),
    call: async () => (seen.push("call"), { ok: true, target_checked: true }),
    onEdit: async () => (seen.push("edit"), { id: "R1.1-1" }),
  });
  const r = await applyTrackedEdit(deps(false), EDIT, { ref: "R1.1" });
  assert.deepEqual([r.ok, r.code], [false, "TaskPaneOutdated"]);
  assert.match(r.message, /not applied/i);
  assert.match(r.message, /reload the task pane/i);
  assert.deepEqual(seen, []);
  assert.equal((await applyTrackedEdit(deps(true), EDIT, { ref: "R1.1" })).ok, true);
  assert.deepEqual(seen, ["attempt", "call", "edit"]);
});

async function world(caps) {
  const dir = await mkdtemp(join(tmpdir(), "pane-caps-"));
  const doc = join(dir, "paper.docx");
  const calls = [];
  const callTool = async (name, args) => {
    calls.push(args);
    return {
      ok: true,
      word_comment_id: 1,
      paragraph_index: 1,
      before: "Cells grew on the substrate.",
      after: "Cells grew on the medium.",
      target_checked: true,
    };
  };
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, ...(caps ? { caps } : {}) }),
    callTool,
    pushTurn: () => {},
    log: () => {},
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
    );
  const edit = (args = EDIT) =>
    applyTrackedEdit({ ...svc.trackedEditDeps("k"), call: callTool }, args, { ref: "R1.1" });
  await ask("review_save_comments", { comments: [{ ref: "R1.1", reviewer: 1, text: "Rename." }] });
  await ask("ledger_add", {
    item: { kind: "forbid", text: "no medium", check: { type: "regex", pattern: "medium" } },
  });
  return { ask, edit, calls, cleanup: () => rm(dir, { recursive: true, force: true }) };
}

test("service: an older task pane (no caps in its hello) is sent no tracked edit, as a normal edit or an exception", async () => {
  for (const caps of [undefined, [], ["something_else"]]) {
    const w = await world(caps);
    try {
      const r = await w.edit({ ...EDIT, replace: "the agar" });
      assert.deepEqual([r.ok, r.code, w.calls.length], [false, "TaskPaneOutdated", 0]);
      let c = (await w.ask("review_load")).revision.comments[0];
      assert.deepEqual([c.changes.length, c.blocked.length], [0, 0], "no attempt left behind");
      // A forbidden text is still refused and stored for the author.
      const blocked = await w.edit();
      assert.deepEqual(blocked.blocked_by, ["L1"]);
      const o = await w.ask("review_override", { ref: "R1.1", blocked_id: blocked.blocked_id });
      assert.deepEqual([o.ok, o.applied, o.result.code], [true, false, "TaskPaneOutdated"]);
      assert.match(o.result.message, /reload the task pane/i);
      c = o.revision.comments[0];
      assert.deepEqual(
        [c.blocked.map((b) => [b.state, b.attempt]), w.calls.length],
        [[["blocked", null]], 0],
      );
    } finally {
      await w.cleanup();
    }
  }
});

test("service: a task pane with the capability gets the edit", async () => {
  const w = await world(["target_hash"]);
  try {
    const r = await w.edit({ ...EDIT, replace: "the agar" });
    assert.deepEqual([r.ok, w.calls.length], [true, 1]);
  } finally {
    await w.cleanup();
  }
});
