// End-to-end through the in-process MCP server: the review-turn guard and
// office_tracked_edit as the model would call them. Reads the registered
// handlers from McpServer._registeredTools (a private field of
// @modelcontextprotocol/sdk 1.29); the test skips if that field moves.
import { test } from "node:test";
import assert from "node:assert/strict";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";

function server(turn, recorded = []) {
  const calls = [];
  const bridge = {
    callTaskpaneTool: async (name, args) => {
      calls.push([name, args]);
      return { ok: true, word_comment_id: 7, paragraph_index: 3, before: "old", after: "new" };
    },
  };
  const s = createOfficeBridgeMcp(bridge, "word", "k", {
    getReviewTurn: () => turn.current,
    trackedEdit: {
      getLedger: async () => [],
      onEdit: async (x) => (recorded.push(x), { id: `${x.ref}-1` }),
    },
  });
  return { tools: s.instance?._registeredTools, calls };
}

const text = (r) => r.content[0].text;

test("review turn: office_replace_text refused, office_tracked_edit applied and recorded", async (t) => {
  const turn = { current: { ref: "R1.1" } };
  const recorded = [];
  const { tools, calls } = server(turn, recorded);
  if (!tools) return t.skip("McpServer._registeredTools not available in this SDK build");
  const refused = await tools.office_replace_text.handler(
    { paragraph_ids: ["p1"], find: "a", replace: "b" },
    {},
  );
  assert.equal(refused.isError, true);
  assert.match(text(refused), /disabled during review turn R1\.1/);
  const ok = await tools.office_tracked_edit.handler(
    {
      comment_ref: "R1.1",
      find: "old",
      replace: "new",
      rationale: { what: "w", why: "y", addresses: "a" },
    },
    {},
  );
  assert.equal(JSON.parse(text(ok)).change_id, "R1.1-1");
  assert.deepEqual(
    calls.map((c) => c[0]),
    ["office_tracked_edit"],
  );
  assert.equal(recorded.length, 1);
});

test("normal turn: upstream tools run, office_tracked_edit refuses", async (t) => {
  const turn = { current: null };
  const { tools, calls } = server(turn);
  if (!tools) return t.skip("McpServer._registeredTools not available in this SDK build");
  await tools.office_replace_text.handler({ paragraph_ids: ["p1"], find: "a", replace: "b" }, {});
  const r = await tools.office_tracked_edit.handler(
    {
      comment_ref: "R1.1",
      find: "old",
      replace: "new",
      rationale: { what: "w", why: "y", addresses: "a" },
    },
    {},
  );
  assert.match(text(r), /only in a review turn/);
  assert.deepEqual(
    calls.map((c) => c[0]),
    ["office_replace_text"],
  );
});
