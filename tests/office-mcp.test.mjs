// Exercise the public MCP protocol after SDK updates: initialization, discovery,
// schema validation, document routing and tool results. No private SDK fields,
// Claude credentials, external MCP servers or running Office host are required.
import { test } from "node:test";
import assert from "node:assert/strict";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";

async function connect(t, host, bridge, opts = {}) {
  const { instance } = createOfficeBridgeMcp(bridge, host, "test-pane", opts);
  const client = new Client({ name: "sidewrite-mcp-test", version: "1.0.0" });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  t.after(async () => {
    await client.close();
    await instance.close();
  });
  await instance.connect(serverTransport);
  await client.connect(clientTransport);
  return client;
}

for (const [host, name, args, result] of [
  [
    "word",
    "office_read_paragraphs",
    { ids: ["p1"] },
    { paragraphs: [{ id: "p1", text: "Hello" }] },
  ],
  [
    "excel",
    "excel_read_range",
    { address: "A1:B2" },
    {
      values: [
        [1, 2],
        [3, 4],
      ],
    },
  ],
]) {
  test(`MCP ${host}: initialize, discover and call a document-scoped tool`, async (t) => {
    const calls = [];
    const client = await connect(t, host, {
      callTaskpaneTool: async (...call) => {
        calls.push(call);
        return result;
      },
    });
    assert.equal(client.getServerVersion().name, "office");
    const { tools } = await client.listTools();
    const prefix = host === "word" ? "office_" : "excel_";
    assert.ok(tools.length > 0);
    assert.ok(tools.every((tool) => tool.name.startsWith(prefix)));
    assert.ok(tools.some((tool) => tool.name === name));
    const response = await client.callTool({ name, arguments: args });
    assert.notEqual(response.isError, true);
    assert.deepEqual(JSON.parse(response.content[0].text), result);
    assert.deepEqual(calls, [[name, args, "test-pane"]]);
  });
}

test("MCP rejects invalid arguments and preserves review-turn write guards", async (t) => {
  const calls = [];
  const client = await connect(
    t,
    "word",
    { callTaskpaneTool: async (...call) => calls.push(call) },
    { getReviewTurn: () => ({ ref: "R1.1" }) },
  );
  const invalid = await client.callTool({
    name: "office_read_paragraphs",
    arguments: { ids: "p1" },
  });
  assert.equal(invalid.isError, true);
  const refused = await client.callTool({
    name: "office_replace_text",
    arguments: { paragraph_ids: ["p1"], find: "old", replace: "new" },
  });
  assert.equal(refused.isError, true);
  assert.match(refused.content[0].text, /disabled during review turn R1\.1/);
  assert.deepEqual(calls, []);
});

test("MCP returns bridge failures as tool errors", async (t) => {
  const client = await connect(t, "excel", {
    callTaskpaneTool: async () => {
      throw new Error("Taskpane disconnected");
    },
  });
  const response = await client.callTool({ name: "excel_list_sheets", arguments: {} });
  assert.equal(response.isError, true);
  assert.match(response.content[0].text, /Taskpane disconnected/);
});
