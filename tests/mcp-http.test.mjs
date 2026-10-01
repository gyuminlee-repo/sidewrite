// The office tools over local streamable HTTP, for the CLI providers' review
// turns. A real HTTP server on a free port and the MCP SDK client.
import { test } from "node:test";
import assert from "node:assert/strict";
import { createServer, request } from "node:http";
import { EventEmitter } from "node:events";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { createMcpHttp, sameToken, REVIEW_TURN_TOOLS } from "../daemon/mcp-http.mjs";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";
import {
  REVIEW_READ_TOOLS,
  REVIEW_WRITE_TOOL,
  REVIEW_BLOCKED_TOOLS,
} from "../daemon/review-tools.mjs";

const TOKEN = "t".repeat(48);
const OTHER = "u".repeat(48);
const INIT = JSON.stringify({
  jsonrpc: "2.0",
  id: 1,
  method: "tools/list",
  params: {},
});

function setup({ turns = { pane1: { ref: "R1.1" } } } = {}) {
  const calls = [];
  const lines = [];
  const bridge = {
    callTaskpaneTool: async (name, args, key) => {
      calls.push([name, key, args]);
      return { ok: true, word_comment_id: 7, paragraph_index: 3, before: "old", after: "new" };
    },
  };
  const recorded = [];
  const mcp = createMcpHttp({
    bridge,
    getReviewTurn: (key) => turns[key] ?? null,
    trackedEditDeps: (key) => ({
      getLedger: async () => [],
      onEdit: async (x) => (recorded.push([key, x.ref]), { id: `${x.ref}-1` }),
    }),
    log: (s) => lines.push(s),
  });
  return { mcp, calls, lines, recorded, turns };
}

async function serve(mcp) {
  const http = createServer(async (req, res) => {
    if (!(await mcp.handle(req, res, req.url.split("?")[0]))) res.writeHead(404).end();
  });
  await new Promise((r) => http.listen(0, "127.0.0.1", r));
  return { http, base: `http://127.0.0.1:${http.address().port}`, port: http.address().port };
}

// A POST with any headers (fetch cannot set Host). Resolves { status, body }.
function post(port, path, headers = {}, body = INIT) {
  return new Promise((resolve, reject) => {
    const req = request({ host: "127.0.0.1", port, path, method: "POST", headers }, (res) => {
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () =>
        resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString("utf8") }),
      );
    });
    req.on("error", reject);
    req.end(body);
  });
}
const auth = (t) => ({
  authorization: `Bearer ${t}`,
  "content-type": "application/json",
  accept: "application/json, text/event-stream",
});

async function connect(base, route, token) {
  const client = new Client({ name: "t", version: "0" });
  await client.connect(
    new StreamableHTTPClientTransport(new URL(`${base}/mcp/${route}`), {
      requestInit: { headers: { authorization: `Bearer ${token}` } },
    }),
  );
  return client;
}

test("a bound route with its token lists only the review-turn tools and routes a call to its pane", async () => {
  const t = setup();
  const { http, base } = await serve(t.mcp);
  try {
    const unbind = t.mcp.bind("r1", { token: TOKEN, key: "pane1", ref: "R1.1" });
    const client = await connect(base, "r1", TOKEN);
    const names = (await client.listTools()).tools.map((x) => x.name).sort();
    assert.deepEqual(names, [...REVIEW_READ_TOOLS, REVIEW_WRITE_TOOL].sort());
    assert.deepEqual([...REVIEW_TURN_TOOLS].sort(), names);
    for (const n of names) assert.equal(REVIEW_BLOCKED_TOOLS.has(n), false, n);
    await client.callTool({ name: "office_find", arguments: { query: "x" } });
    assert.deepEqual(t.calls[0].slice(0, 2), ["office_find", "pane1"]);
    assert.ok(t.lines.some((l) => l === "[mcp-http] POST route=r1 key=pane1 ref=R1.1"));
    assert.equal(t.lines.join("\n").includes(TOKEN), false, "the token is never logged");
    await client.close();
    unbind();
    const gone = await post(http.address().port, "/mcp/r1", auth(TOKEN));
    assert.equal(gone.status, 401);
  } finally {
    http.close();
  }
});

test("no token, a wrong token, another route's token: 401 with an empty body and no tool name", async () => {
  const t = setup({ turns: { pane1: { ref: "R1.1" }, pane2: { ref: "R2.1" } } });
  const { http, port } = await serve(t.mcp);
  try {
    t.mcp.bind("r1", { token: TOKEN, key: "pane1", ref: "R1.1" });
    t.mcp.bind("r2", { token: OTHER, key: "pane2", ref: "R2.1" });
    const tries = [
      await post(port, "/mcp/r1", { "content-type": "application/json" }),
      await post(port, "/mcp/r1", auth("x".repeat(48))),
      await post(port, "/mcp/r1", auth(OTHER)), // the token of r2
      await post(port, "/mcp/r1", { ...auth(TOKEN), authorization: TOKEN }), // no "Bearer "
      await post(port, "/mcp/nope", auth(TOKEN)), // an unbound route
    ];
    for (const r of tries) {
      assert.equal(r.status, 401);
      assert.equal(r.body, "");
      assert.equal(/office_/.test(r.body), false);
    }
    assert.equal(t.calls.length, 0);
    const good = await post(port, "/mcp/r1", auth(TOKEN));
    assert.equal(good.status, 200);
    assert.match(good.body, /office_tracked_edit/);
  } finally {
    http.close();
  }
});

test("browser and rebinding requests are refused before the token is looked at", async () => {
  const t = setup();
  const { http, port } = await serve(t.mcp);
  try {
    t.mcp.bind("r1", { token: TOKEN, key: "pane1", ref: "R1.1" });
    // Any Origin header: a page in a browser, also one served from loopback.
    for (const origin of ["http://evil.example", `http://127.0.0.1:${port}`, "null"]) {
      const r = await post(port, "/mcp/r1", { ...auth(TOKEN), origin });
      assert.deepEqual([r.status, r.body], [403, ""], origin);
    }
    // DNS rebinding: the browser resolves the attacker's name to 127.0.0.1,
    // so the Host header still carries that name.
    for (const host of ["evil.example", `evil.example:${port}`, `127.0.0.1.evil.example:${port}`]) {
      const r = await post(port, "/mcp/r1", { ...auth(TOKEN), host });
      assert.deepEqual([r.status, r.body], [403, ""], host);
    }
    for (const host of [`localhost:${port}`, `127.0.0.1:${port}`, `[::1]:${port}`]) {
      assert.equal((await post(port, "/mcp/r1", { ...auth(TOKEN), host })).status, 200, host);
    }
    // A cross-site fetch marks itself; a CLI never sends this header.
    const cross = await post(port, "/mcp/r1", { ...auth(TOKEN), "sec-fetch-site": "cross-site" });
    assert.equal(cross.status, 403);
    assert.equal(t.calls.length, 0);
  } finally {
    http.close();
  }
});

test("a peer that is not loopback is refused; only POST is served; other paths are not ours", async () => {
  const t = setup();
  t.mcp.bind("r1", { token: TOKEN, key: "pane1", ref: "R1.1" });
  const fake = (over = {}) => {
    const req = new EventEmitter();
    Object.assign(req, {
      method: "POST",
      headers: { authorization: `Bearer ${TOKEN}`, host: "127.0.0.1:47834" },
      socket: { remoteAddress: "127.0.0.1" },
      ...over,
    });
    const res = { status: null, writeHead: (s) => ((res.status = s), res), end: () => res };
    return { req, res };
  };
  const lan = fake({ socket: { remoteAddress: "192.168.0.7" } });
  assert.equal(await t.mcp.handle(lan.req, lan.res, "/mcp/r1"), true);
  assert.equal(lan.res.status, 403);
  const mapped = fake({ socket: { remoteAddress: "::ffff:10.0.0.5" } });
  await t.mcp.handle(mapped.req, mapped.res, "/mcp/r1");
  assert.equal(mapped.res.status, 403);
  const get = fake({ method: "GET" });
  await t.mcp.handle(get.req, get.res, "/mcp/r1");
  assert.equal(get.res.status, 405);
  const other = fake();
  assert.equal(await t.mcp.handle(other.req, other.res, "/bridge-token"), false);
  assert.equal(await t.mcp.handle(other.req, other.res, "/mcp/r1/extra"), false);
  assert.equal(other.res.status, null);
});

test("a binding cannot widen the tool set: extra fields are ignored, Excel and writer tools are never listed", async () => {
  const t = setup();
  const { http, base } = await serve(t.mcp);
  try {
    t.mcp.bind("r1", {
      token: TOKEN,
      key: "pane1",
      ref: "R1.1",
      // What a careless or hostile caller might pass.
      only: new Set(["office_replace_text", "excel_write_range"]),
      host: "excel",
      getReviewTurn: () => null,
    });
    const client = await connect(base, "r1", TOKEN);
    const names = (await client.listTools()).tools.map((x) => x.name);
    assert.deepEqual(names.sort(), [...REVIEW_TURN_TOOLS].sort());
    const r = await client
      .callTool({
        name: "office_replace_text",
        arguments: { find: "a", replace: "b" },
      })
      .then(
        (x) => x,
        (e) => ({ isError: true, content: [{ text: String(e.message) }] }),
      );
    assert.equal(r.isError, true, "an unlisted tool cannot be called");
    assert.equal(t.calls.length, 0);
    await client.close();
    assert.equal(Object.isFrozen(REVIEW_TURN_TOOLS), true);
  } finally {
    http.close();
  }
});

test("office_tracked_edit goes through the same recorder, bound to the comment of the turn", async () => {
  const t = setup();
  const { http, base } = await serve(t.mcp);
  try {
    t.mcp.bind("r1", { token: TOKEN, key: "pane1", ref: "R1.1" });
    const client = await connect(base, "r1", TOKEN);
    const edit = (comment_ref) =>
      client.callTool({
        name: "office_tracked_edit",
        arguments: {
          comment_ref,
          find: "old",
          replace: "new",
          rationale: { what: "w", why: "y", addresses: "a" },
        },
      });
    const ok = await edit("R1.1");
    assert.equal(JSON.parse(ok.content[0].text).change_id, "R1.1-1");
    assert.deepEqual(t.recorded, [["pane1", "R1.1"]]);
    const other = await edit("R1.2");
    assert.match(other.content[0].text, /does not match the comment of this turn R1\.1/);
    assert.deepEqual(
      t.calls.map((c) => c[0]),
      ["office_tracked_edit"],
      "only the matching edit reached Word",
    );
    await client.close();
  } finally {
    http.close();
  }
});

test("outside its review turn a route serves nothing: turn ended, or another comment's turn on that pane", async () => {
  const t = setup();
  const { http, port } = await serve(t.mcp);
  try {
    t.mcp.bind("r1", { token: TOKEN, key: "pane1", ref: "R1.1" });
    assert.equal((await post(port, "/mcp/r1", auth(TOKEN))).status, 200);
    t.turns.pane1 = { ref: "R1.2" }; // a later turn of the pane, the old route still bound
    const stale = await post(port, "/mcp/r1", auth(TOKEN));
    assert.deepEqual([stale.status, stale.body], [409, ""]);
    delete t.turns.pane1; // no turn at all
    assert.equal((await post(port, "/mcp/r1", auth(TOKEN))).status, 409);
    assert.equal(t.calls.length, 0);
  } finally {
    http.close();
  }
});

test("bad requests: a body that is not JSON, a body over the limit, a bad route name or a short token at bind", async () => {
  const t = setup();
  const { http, port } = await serve(t.mcp);
  try {
    t.mcp.bind("r1", { token: TOKEN, key: "pane1", ref: "R1.1" });
    assert.equal((await post(port, "/mcp/r1", auth(TOKEN), "{ nope")).status, 400);
    const big = JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "x",
      params: { p: "a".repeat(5e6) },
    });
    assert.equal((await post(port, "/mcp/r1", auth(TOKEN), big)).status, 413);
    assert.throws(() => t.mcp.bind("../x", { token: TOKEN, key: "k", ref: "R1.1" }), /route/);
    assert.throws(() => t.mcp.bind("r9", { token: "short", key: "k", ref: "R1.1" }), /token/);
    assert.throws(() => t.mcp.bind("r9", { token: TOKEN, key: "k" }), /ref/);
    assert.throws(() => t.mcp.bind("r1", { token: OTHER, key: "k", ref: "R1.1" }), /already bound/);
  } finally {
    http.close();
  }
});

test("token comparison: equal only for the same string, whatever the lengths", () => {
  assert.equal(sameToken("Bearer abc", "Bearer abc"), true);
  assert.equal(sameToken("Bearer abc", "Bearer abd"), false);
  assert.equal(sameToken("Bearer abc", "Bearer abcd"), false);
  assert.equal(sameToken("", "Bearer abc"), false);
  assert.equal(sameToken(undefined, "Bearer abc"), false);
});

// office-tools.mjs: opts.only filters what is registered; the review guard
// stays under it (a writer that slipped into `only` is still refused).
test("createOfficeBridgeMcp: only filters the registered tools and the guard still holds", async () => {
  const bridge = { callTaskpaneTool: async () => ({ ok: true }) };
  const s = createOfficeBridgeMcp(bridge, "word", "k", {
    getReviewTurn: () => ({ ref: "R1.2" }),
    only: new Set(["office_find", "office_replace_text"]),
  });
  const tools = s.instance?._registeredTools;
  if (!tools) return;
  assert.deepEqual(Object.keys(tools).sort(), ["office_find", "office_replace_text"]);
  const r = await tools.office_replace_text.handler({ find: "a", replace: "b" }, {});
  assert.equal(r.isError, true);
  assert.match(r.content[0].text, /disabled during review turn R1\.2/);
});
