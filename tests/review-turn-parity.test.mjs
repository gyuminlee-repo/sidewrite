// A review turn must leave the same record whoever runs it. The same edits
// go through (a) the in-process MCP server a Claude turn uses, where the
// model calls office_tracked_edit itself, and (b) the proposal runner of a
// provider without tools (Codex), where the model returns the edits as a
// list and the daemon applies them. Each has its own review service and
// revision file; the recorded changes, blocked edits, id counters, decision
// receipts, events and Word calls are then compared.
//
// What differs by design is asserted on its own: a proposal turn stops at
// the first edit a forbid rule blocks (the Claude turn goes on), so the
// edits after it come in a second turn; a proposal turn writes one
// proposal_batch receipt; and the guards a model with tools needs (another
// writer, the comment of another turn) have no counterpart, because a
// proposal has no tool name and no comment_ref.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";
import { createProposalTurns } from "../daemon/proposal-turns.mjs";

const KEY = "pane";
const REF = "R1.1";
const PARAS = [
  "2. Materials and Methods",
  "Cultures were grown in sealed flasks and held at 30 degrees for 5 days.",
  "The carbon source was methanol.",
];
const R = { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] };

// The edits, in order: paragraph, find, replace, and what happens to each.
const EDITS = [
  { paragraph: "p1", find: "30 degrees", replace: "30 °C", expect: "applied" },
  // The forbid list (L1) stops this one: it becomes a blocked edit.
  {
    paragraph: "p2",
    find: "methanol",
    replace: "the novel substrate methanol",
    expect: "refused",
  },
  // A requirement id the comment does not list.
  {
    paragraph: "p1",
    find: "flasks",
    replace: "bottles",
    requirement_ids: ["q9"],
    expect: "refused",
  },
  // Drops two numbers of the paragraph: applied, with the lost facts noted.
  { paragraph: "p1", find: "held at 30 °C for 5 days", replace: "kept warm", expect: "applied" },
];
const rationaleOf = (e) => ({
  ...R,
  ...(e.requirement_ids ? { requirement_ids: e.requirement_ids } : {}),
});

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
      paragraph_index: args.expected_paragraph?.index ?? 1,
      before: PARAS[1],
      after: PARAS[1].replace(args.find, args.replace ?? ""),
      target_checked: true,
    };
  };
  const events = [];
  const svc = createReviewService({
    // A task pane of this build: it compares the paragraph before an edit.
    getContext: () => ({ activeDoc: doc, caps: ["target_hash", "proposal_turns"] }),
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
  // A change counts as made after the snapshot by its time: not the same ms.
  await new Promise((r) => setTimeout(r, 5));
  wordCalls.length = 0;
  return { dir, doc, svc, ask, events, wordCalls, bridge: { callTaskpaneTool } };
}

// (a) the server of a Claude turn, its handlers called as the SDK calls
// them. Returns the outcome of each edit, or null when the SDK build does
// not show its tools.
async function claudePath(w) {
  const s = createOfficeBridgeMcp(w.bridge, "word", KEY, {
    getReviewTurn: () => ({ ref: REF }),
    trackedEdit: w.svc.trackedEditDeps(KEY),
  });
  const tools = s.instance?._registeredTools;
  if (!tools) return null;
  const got = [];
  for (const e of EDITS) {
    const r = await tools.office_tracked_edit.handler(
      {
        comment_ref: REF,
        find: e.find,
        replace: e.replace,
        paragraph_hint: e.paragraph,
        rationale: rationaleOf(e),
      },
      {},
    );
    let body = null;
    try {
      body = JSON.parse(r.content?.[0]?.text ?? "");
    } catch {
      body = null;
    }
    got.push(r.isError || body?.ok === false ? "refused" : "applied");
  }
  return got;
}

// (b) the proposal runner. One turn per list: the first ends at the blocked
// edit, the second brings what came after it.
async function proposalPath(w) {
  const proposal = (e) => ({
    paragraph_id: e.paragraph,
    find: e.find,
    op: "replace",
    text: e.replace,
    ...rationaleOf(e),
  });
  const turnsOf = [EDITS.slice(0, 3), EDITS.slice(3)];
  const outcomes = [];
  const ended = [];
  for (const list of turnsOf) {
    let done;
    const end = new Promise((r) => (done = r));
    const turns = createProposalTurns({
      handleFor: () => ({
        model: "model-under-test",
        run: async () => {
          const json = { proposals: list.map(proposal), note: "" };
          return { ok: true, text: JSON.stringify(json), json };
        },
      }),
      trackedEditDeps: w.svc.trackedEditDeps,
      call: (name, args) => w.bridge.callTaskpaneTool(name, args),
      sendText: () => {},
      sendEvent: () => {},
      endTurn: (key, outcome) => (ended.push(outcome), done()),
      log: () => {},
    });
    await turns.start(
      KEY,
      "word",
      { ref: REF },
      { at: w.svc.at(KEY), selected: [], choice: { id: "codex", model: null } },
    );
    await end;
  }
  const decisions = (await readFile(join(w.dir, "paper.decisions.jsonl"), "utf8"))
    .trim()
    .split("\n")
    .map((l) => JSON.parse(l));
  for (const d of decisions.filter((x) => x.kind === "proposal_batch")) {
    for (const p of d.output.proposals)
      outcomes.push(p.outcome === "applied" ? "applied" : "refused");
  }
  return { outcomes, ended, batches: decisions.filter((x) => x.kind === "proposal_batch") };
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
    .map((l) => JSON.parse(l))
    .filter((d) => d.kind !== "proposal_batch"); // asserted on its own
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

test("the same edits leave the same record through a Claude turn and through proposal turns", async (t) => {
  const a = await world();
  const b = await world();
  try {
    const viaClaude = await claudePath(a);
    if (!viaClaude) return t.skip("McpServer._registeredTools not available in this SDK build");
    const viaProposal = await proposalPath(b);
    const expected = EDITS.map((e) => e.expect);
    assert.deepEqual(viaClaude, expected, "Claude path outcomes");
    assert.deepEqual(viaProposal.outcomes, expected, "proposal path outcomes");
    const ra = await record(a, viaClaude);
    const rb = await record(b, viaProposal.outcomes);
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
    // blocked_seq: the two edits sent to Word each spent an attempt id (they
    // are on file before the call), and the blocked edit one.
    assert.deepEqual(ra.counters, { ledger_seq: 1, requirement_seq: 1, blocked_seq: 3 });
    assert.ok(ra.decisions.length >= 3);
    assert.ok(ra.decisions.every((d) => d.kind === "requirement_link"));
    assert.deepEqual(
      ra.word.map((c) => c[0]),
      ["office_tracked_edit", "office_tracked_edit", "office_tracked_edit"],
    );
    // Every edit that reached Word named the paragraph it was checked on.
    for (const w of [ra.word, rb.word]) {
      for (const [, args] of w) {
        assert.match(args.expected_before_hash, /^[0-9a-f]+$/);
        assert.equal(typeof args.expected_paragraph.index, "number");
      }
    }

    // What differs by design. The first proposal turn stopped at the blocked
    // edit and did not look at the one after it; the second applied the rest.
    assert.deepEqual(
      viaProposal.ended.map((o) => o.ok),
      [false, true],
    );
    assert.match(viaProposal.ended[0].error, /^The turn stopped at P2 \(Blocked\)\. .*L1/);
    assert.deepEqual(
      viaProposal.batches.map((d) => d.output.proposals.map((p) => [p.outcome, p.code])),
      [
        [
          ["applied", null],
          ["failed", "Blocked"],
          ["rejected", "UnknownRequirement"],
        ],
        [["applied", null]],
      ],
    );
  } finally {
    await rm(a.dir, { recursive: true, force: true });
    await rm(b.dir, { recursive: true, force: true });
  }
});
