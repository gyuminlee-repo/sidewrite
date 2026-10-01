// A review turn must leave the same record whoever runs it. The same edit
// script goes through (a) the in-process MCP server a Claude turn uses and
// (b) the HTTP MCP endpoint a CLI turn uses, each over its own review
// service and revision file; the recorded changes, blocked edits, id
// counters, decision receipts and Word calls are then compared.
import { test } from "node:test";
import assert from "node:assert/strict";
import { createServer } from "node:http";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { createReviewService } from "../daemon/review-service.mjs";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";
import { createMcpHttp } from "../daemon/mcp-http.mjs";

const KEY = "pane";
const REF = "R1.1";
const TOKEN = "p".repeat(48);
const PARAS = [
  "2. Materials and Methods",
  "Cultures were grown in sealed flasks and held at 30 degrees for 5 days.",
  "The carbon source was methanol.",
];
const R = { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] };

// What the model does in the turn, in order. `expect` is checked on both paths.
const SCRIPT = [
  {
    tool: "office_tracked_edit",
    args: { comment_ref: REF, find: "30 degrees", replace: "30 °C", rationale: R },
    expect: "applied",
  },
  {
    // The forbid list (L1) stops this one: it becomes a blocked edit.
    tool: "office_tracked_edit",
    args: {
      comment_ref: REF,
      find: "methanol",
      replace: "the novel substrate methanol",
      rationale: R,
    },
    expect: "refused",
  },
  {
    // A writer other than office_tracked_edit.
    tool: "office_replace_text",
    args: { find: "methanol", replace: "ethanol" },
    expect: "refused",
  },
  {
    // The comment of another turn.
    tool: "office_tracked_edit",
    args: { comment_ref: "R1.2", find: "flasks", replace: "bottles", rationale: R },
    expect: "refused",
  },
  {
    // A requirement id the comment does not list.
    tool: "office_tracked_edit",
    args: {
      comment_ref: REF,
      find: "flasks",
      replace: "bottles",
      rationale: { ...R, requirement_ids: ["q9"] },
    },
    expect: "refused",
  },
  {
    // Drops two numbers of the paragraph: applied, with the lost facts noted.
    tool: "office_tracked_edit",
    args: {
      comment_ref: REF,
      find: "held at 30 degrees for 5 days",
      replace: "kept warm",
      rationale: R,
    },
    expect: "applied",
  },
  { tool: "office_read_paragraphs", args: {}, expect: "applied" },
];

async function world() {
  const dir = await mkdtemp(join(tmpdir(), "parity-"));
  const doc = join(dir, "paper.docx");
  const wordCalls = [];
  let commentId = 10;
  const callTaskpaneTool = async (name, args) => {
    wordCalls.push([name, args]);
    if (name !== "office_tracked_edit") return { ok: true, paragraphs: [] };
    return {
      ok: true,
      word_comment_id: ++commentId,
      paragraph_index: 1,
      before: PARAS[1],
      after: PARAS[1].replace(args.find, args.replace ?? ""),
    };
  };
  const events = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc }),
    callTool: (name, args) => callTaskpaneTool(name, args),
    pushTurn: () => {},
    getModel: () => "model-under-test",
    sendEvent: (key, e) => events.push(e),
    log: () => {},
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, KEY, "word"),
    );
  await ask("review_save_comments", {
    comments: [{ ref: REF, reviewer: 1, text: "Use Celsius.", requirements: ["Use Celsius"] }],
  });
  await ask("ledger_add", {
    item: {
      kind: "forbid",
      text: "no novel",
      check: { type: "regex", pattern: String.raw`\bnovel\b`, flags: "i" },
    },
  });
  await ask("review_refresh", {
    snapshot: {
      paragraphs: PARAS.map((text, index) => ({
        index,
        id: `p${index}`,
        text,
        reviewed_text: text,
        w_ins: 0,
        w_del: 0,
      })),
      comments: [],
    },
  });
  wordCalls.length = 0;
  return { dir, doc, svc, ask, events, wordCalls, bridge: { callTaskpaneTool } };
}

const outcome = (r) => {
  const text = r.content?.[0]?.text ?? "";
  let body = null;
  try {
    body = JSON.parse(text);
  } catch {
    body = null;
  }
  return r.isError || body?.ok === false ? "refused" : "applied";
};

// (a) the server of a Claude turn, its handlers called as the SDK calls them.
async function claudePath(w) {
  const s = createOfficeBridgeMcp(w.bridge, "word", KEY, {
    getReviewTurn: () => ({ ref: REF }),
    trackedEdit: w.svc.trackedEditDeps(KEY),
  });
  const tools = s.instance?._registeredTools;
  if (!tools) return null;
  const got = [];
  for (const step of SCRIPT) got.push(outcome(await tools[step.tool].handler(step.args, {})));
  return got;
}

// (b) the endpoint of a CLI turn, called by an MCP client over HTTP.
async function cliPath(w) {
  const mcp = createMcpHttp({
    bridge: w.bridge,
    getReviewTurn: () => ({ ref: REF }),
    trackedEditDeps: (key) => w.svc.trackedEditDeps(key),
    log: () => {},
  });
  const http = createServer(async (req, res) => {
    if (!(await mcp.handle(req, res, req.url.split("?")[0]))) res.writeHead(404).end();
  });
  await new Promise((r) => http.listen(0, "127.0.0.1", r));
  try {
    mcp.bind("turn", { token: TOKEN, key: KEY, ref: REF });
    const client = new Client({ name: "parity", version: "0" });
    await client.connect(
      new StreamableHTTPClientTransport(
        new URL(`http://127.0.0.1:${http.address().port}/mcp/turn`),
        { requestInit: { headers: { authorization: `Bearer ${TOKEN}` } } },
      ),
    );
    const got = [];
    for (const step of SCRIPT) {
      const r = await client.callTool({ name: step.tool, arguments: step.args }).then(
        (x) => x,
        (e) => ({ isError: true, content: [{ text: String(e?.message ?? e) }] }),
      );
      got.push(outcome(r));
    }
    await client.close();
    return got;
  } finally {
    http.close();
  }
}

// Times differ between two runs; everything else must not.
const ISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/;
function normalize(x, dir) {
  if (Array.isArray(x)) return x.map((v) => normalize(v, dir));
  if (x && typeof x === "object") {
    return Object.fromEntries(Object.entries(x).map(([k, v]) => [k, normalize(v, dir)]));
  }
  if (typeof x === "string") return ISO.test(x) ? "<time>" : x.split(dir).join("<dir>");
  return x;
}

async function record(w, got) {
  // The author then approves the blocked edit: the override flow of the
  // service, the same for both.
  const rev0 = JSON.parse(await readFile(join(w.dir, "paper.revision.json"), "utf8"));
  const blockedId = rev0.comments[0].blocked[0]?.id;
  const override = blockedId
    ? await w.ask("review_override", { ref: REF, blocked_id: blockedId })
    : null;
  const rev = JSON.parse(await readFile(join(w.dir, "paper.revision.json"), "utf8"));
  const decisions = (await readFile(join(w.dir, "paper.decisions.jsonl"), "utf8"))
    .trim()
    .split("\n")
    .map((l) => JSON.parse(l));
  return normalize(
    {
      outcomes: got,
      blocked_before_override: rev0.comments[0].blocked,
      override: override && { ok: override.ok, applied: override.applied },
      changes: rev.comments[0].changes,
      blocked: rev.comments[0].blocked,
      counters: {
        ledger_seq: rev.ledger_seq,
        requirement_seq: rev.comments[0].requirement_seq,
        blocked_seq: rev.comments[0].blocked_seq,
      },
      status: rev.comments[0].status,
      decisions,
      events: w.events.map((e) => [e.event, e.ref ?? null, e.change?.id ?? e.blocked?.id ?? null]),
      word: w.wordCalls,
    },
    w.dir,
  );
}

test("the same edit script leaves the same record through the Claude path and the CLI path", async (t) => {
  const a = await world();
  const b = await world();
  try {
    const viaClaude = await claudePath(a);
    if (!viaClaude) return t.skip("McpServer._registeredTools not available in this SDK build");
    const viaCli = await cliPath(b);
    assert.deepEqual(
      viaClaude,
      SCRIPT.map((s) => s.expect),
      "Claude path outcomes",
    );
    assert.deepEqual(
      viaCli,
      SCRIPT.map((s) => s.expect),
      "CLI path outcomes",
    );
    const ra = await record(a, viaClaude);
    const rb = await record(b, viaCli);
    assert.deepEqual(rb, ra);

    // The record is not trivially equal because it is empty.
    assert.deepEqual(
      ra.changes.map((c) => c.id),
      ["R1.1-1", "R1.1-2", "R1.1-3"],
    );
    assert.equal(ra.blocked_before_override.length, 1);
    assert.deepEqual(ra.blocked_before_override[0].blocked_by, ["L1"]);
    assert.deepEqual(ra.override, { ok: true, applied: true });
    assert.equal(ra.blocked.length, 0);
    assert.ok(ra.changes[1].lost_facts.length >= 1, "L1 lost facts are recorded on the change");
    assert.deepEqual(ra.counters, { ledger_seq: 1, requirement_seq: 1, blocked_seq: 1 });
    assert.ok(ra.decisions.length >= 3);
    assert.ok(ra.decisions.every((d) => d.kind === "requirement_link"));
    assert.deepEqual(
      ra.word.map((c) => c[0]),
      [
        "office_tracked_edit",
        "office_tracked_edit",
        "office_read_paragraphs",
        "office_tracked_edit",
      ],
    );
  } finally {
    await rm(a.dir, { recursive: true, force: true });
    await rm(b.dir, { recursive: true, force: true });
  }
});
