// What a review turn leaves on record. An edit script goes through the
// in-process MCP server a Claude turn uses, over a review service and a
// revision file of its own; the recorded changes, blocked edits, id counters,
// decision receipts and Word calls are then checked.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";

const KEY = "pane";
const REF = "R1.1";
const PARAS = [
  "2. Materials and Methods",
  "Cultures were grown in sealed flasks and held at 30 degrees for 5 days.",
  "The carbon source was methanol.",
];
const R = { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] };

// What the model does in the turn, in order. `expect` is what each step must end as.
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
    // A task pane of this build: it compares the paragraph before an edit.
    getContext: () => ({ activeDoc: doc, caps: ["target_hash", "author_insertion_guard"] }),
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

// The server of a Claude turn, its handlers called as the SDK calls them.
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

// Times and the folder differ from run to run.
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
  // service.
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

test("the edit script of a review turn leaves its record: changes, the blocked edit, counters, receipts, Word calls", async (t) => {
  const a = await world();
  try {
    const viaClaude = await claudePath(a);
    if (!viaClaude) return t.skip("McpServer._registeredTools not available in this SDK build");
    assert.deepEqual(
      viaClaude,
      SCRIPT.map((s) => s.expect),
    );
    const ra = await record(a, viaClaude);
    assert.deepEqual(
      ra.changes.map((c) => c.id),
      ["R1.1-1", "R1.1-2", "R1.1-3"],
    );
    assert.equal(ra.blocked_before_override.length, 1);
    assert.deepEqual(ra.blocked_before_override[0].blocked_by, ["L1"]);
    assert.deepEqual(ra.override, { ok: true, applied: true });
    assert.equal(ra.blocked.length, 0);
    assert.ok(ra.changes[1].lost_facts.length >= 1, "L1 lost facts are recorded on the change");
    // blocked_seq: the two edits sent to Word each spent an attempt id (they
    // are on file before the call), and the blocked edit one.
    assert.deepEqual(ra.counters, { ledger_seq: 1, requirement_seq: 1, blocked_seq: 3 });
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
  }
});
