// daemon/proposal-fixture.mjs: a stand-in for the provider of a proposal
// turn, for the acceptance run only. With CC_OFFICE_E2E_PROPOSAL_FIXTURE set
// to a JSON file, a proposal turn takes its reply from that file and asks no
// provider: the stretch from the runner to Word can be measured without a
// model. Without the variable nothing of this is in use.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fixtureHandleFor } from "../daemon/proposal-fixture.mjs";
import { proposalTurnsFor } from "../daemon/proposal-turns.mjs";
import { readProposalReply } from "../daemon/proposal.mjs";

const PROMPT = [
  "[Proposal turn R1.2]",
  "Reviewer 1, comment R1.2 (verbatim):",
  "<<<",
  "Call it a clarified lysate.",
  ">>>",
  "",
  "Paragraphs of the manuscript (all of it), one per line as [id] text:",
  "[p0] 2. Materials and Methods",
  "[A1B2] A cell-free extract was prepared by sonication.",
  "[p2] The cell-free extract showed activity.",
].join("\n");
const P = (over) => ({
  find: "cell-free extract",
  op: "replace",
  text: "clarified lysate",
  what: "w",
  why: "y",
  addresses: "a",
  requirement_ids: [],
  ...over,
});

async function withFile(data, fn) {
  const dir = await mkdtemp(join(tmpdir(), "proposal-fixture-"));
  const file = join(dir, "fixture.json");
  await writeFile(file, typeof data === "string" ? data : JSON.stringify(data));
  try {
    return await fn(file);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

test("the reply of a turn comes from the file, by the comment of the prompt; a paragraph is named by id or by its text", async () => {
  const data = {
    "R1.2": {
      proposals: [
        P({ paragraph_contains: "was prepared by sonication" }),
        P({ paragraph_id: "p2", find: "showed activity", text: "was active" }),
      ],
      note: "from a file",
    },
  };
  await withFile(data, async (file) => {
    const lines = [];
    const h = fixtureHandleFor(file, { log: (l) => lines.push(l) })({ id: "codex", model: "x" });
    assert.deepEqual([h.provider, h.model], ["fixture", "fixture:fixture.json"]);
    const r = await h.run({ system: "S", prompt: PROMPT });
    assert.equal(r.ok, true);
    assert.deepEqual(
      r.json.proposals.map((p) => [p.paragraph_id, p.find]),
      [
        ["A1B2", "cell-free extract"],
        ["p2", "showed activity"],
      ],
    );
    assert.equal("paragraph_contains" in r.json.proposals[0], false);
    assert.equal(r.text, JSON.stringify(r.json));
    // The reply is in the form a provider's reply is read in.
    const read = readProposalReply(r);
    assert.deepEqual([read.ok, read.proposals.length, read.note], [true, 2, "from a file"]);
    // One log line per run, with the comment and no text.
    assert.deepEqual(lines, ["[proposal] fixture reply ref=R1.2 proposals=2"]);
  });
});

test("a comment the file does not have, a paragraph that is not shown, a broken file: a failed reply, never a throw", async () => {
  await withFile({ "R1.1": { proposals: [], note: "" } }, async (file) => {
    const h = fixtureHandleFor(file, { log: () => {} })(null);
    assert.deepEqual(await h.run({ prompt: PROMPT }), {
      ok: false,
      error: "the proposal fixture has no reply for R1.2",
    });
    assert.deepEqual(await h.run({ prompt: "no first line of a turn" }), {
      ok: false,
      error: "the proposal fixture could not tell the comment of the prompt",
    });
  });
  await withFile(
    { "R1.2": { proposals: [P({ paragraph_contains: "not in the prompt" })] } },
    async (file) => {
      const r = await fixtureHandleFor(file, { log: () => {} })(null).run({ prompt: PROMPT });
      assert.deepEqual(r, {
        ok: false,
        error: "the proposal fixture names a paragraph that is not shown (proposal 1 of R1.2)",
      });
    },
  );
  await withFile("{ not json", async (file) => {
    const r = await fixtureHandleFor(file, { log: () => {} })(null).run({ prompt: PROMPT });
    assert.equal(r.ok, false);
    assert.match(r.error, /^the proposal fixture could not be read: /);
  });
  // A stopped turn is a stopped reply.
  await withFile({ "R1.2": { proposals: [] } }, async (file) => {
    const r = await fixtureHandleFor(file, { log: () => {} })(null).run({
      prompt: PROMPT,
      signal: AbortSignal.abort(),
    });
    assert.deepEqual(r, { ok: false, aborted: true, error: "stopped" });
  });
});

test("proposalTurnsFor: the fixture stands in only when the e2e variable is set, and says so once", async () => {
  const routed = [];
  const router = {
    handleFor: (choice) => (routed.push(choice), { model: "codex:x", run: async () => ({}) }),
  };
  const base = {
    bridge: {},
    review: { trackedEditDeps: () => ({}), includeForbid: () => true },
    router,
    endTurn: () => {},
  };
  const plain = [];
  proposalTurnsFor({ ...base, env: {}, log: (l) => plain.push(l) });
  assert.deepEqual(plain, [], "nothing is said and nothing is replaced without the variable");
  const lines = [];
  proposalTurnsFor({
    ...base,
    env: { CC_OFFICE_E2E_PROPOSAL_FIXTURE: "/tmp/f.json" },
    log: (l) => lines.push(l),
  });
  assert.deepEqual(lines, [
    "[proposal] E2E FIXTURE: the replies of proposal turns come from /tmp/f.json (CC_OFFICE_E2E_PROPOSAL_FIXTURE). No provider is asked for a review turn.",
  ]);
  // No other environment variable opens it.
  const src = (await import("node:fs")).readFileSync(
    new URL("../daemon/proposal-turns.mjs", import.meta.url),
    "utf8",
  );
  assert.equal((src.match(/fixtureHandleFor\(/g) ?? []).length, 1);
  assert.match(src, /env\.CC_OFFICE_E2E_PROPOSAL_FIXTURE\s*\?\s*fixtureHandleFor\(/);
});
