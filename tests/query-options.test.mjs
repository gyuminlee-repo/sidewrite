import { test } from "node:test";
import assert from "node:assert/strict";
import {
  resolveSettingSources,
  userMcpEnabled,
  buildQueryOptions,
  formatUsageLine,
} from "../daemon/query-options.mjs";

const base = {
  cwd: "/w",
  append: "A",
  officeMcp: { name: "office" },
  userMcpServers: { visio: { type: "http", url: "http://x" }, office: { bogus: true } },
  disallowedTools: ["t"],
  canUseTool: () => {},
  model: "sonnet",
  abortController: new AbortController(),
  stderr: () => {},
  resume: null,
};

test("settingSources defaults to project only", () => {
  assert.deepEqual(resolveSettingSources({}), ["project"]);
  assert.deepEqual(buildQueryOptions(base, {}).settingSources, ["project"]);
});

test("settingSources=all omits the key (SDK loads every source, upstream behavior)", () => {
  assert.equal(resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: "all" }), undefined);
  const opts = buildQueryOptions(base, { CC_OFFICE_SETTING_SOURCES: "all" });
  assert.equal("settingSources" in opts, false);
});

test("settingSources empty string means isolation, list is filtered", () => {
  assert.deepEqual(resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: "" }), []);
  assert.deepEqual(
    resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: "user, project,bogus" }),
    ["user", "project"],
  );
});

test("user MCP from ~/.claude.json is off by default", () => {
  assert.equal(userMcpEnabled({}), false);
  assert.deepEqual(Object.keys(buildQueryOptions(base, {}).mcpServers), ["office"]);
});

test("CC_OFFICE_LOAD_USER_MCP=1 merges user servers, office wins", () => {
  const env = { CC_OFFICE_LOAD_USER_MCP: "1" };
  const opts = buildQueryOptions(base, env);
  assert.deepEqual(Object.keys(opts.mcpServers).sort(), ["office", "visio"]);
  assert.equal(opts.mcpServers.office, base.officeMcp);
});

test("resume only when a session id is given", () => {
  assert.equal("resume" in buildQueryOptions(base, {}), false);
  assert.equal(buildQueryOptions({ ...base, resume: "abc" }, {}).resume, "abc");
});

test("upstream option shape is preserved", () => {
  const o = buildQueryOptions(base, {});
  assert.deepEqual(o.systemPrompt, { type: "preset", preset: "claude_code", append: "A" });
  assert.equal(o.includePartialMessages, true);
  assert.equal(o.model, "sonnet");
  assert.equal(o.cwd, "/w");
  assert.equal(o.canUseTool, base.canUseTool);
  assert.equal(o.abortController, base.abortController);
  assert.equal(o.stderr, base.stderr);
  assert.deepEqual(o.disallowedTools, ["t"]);
});

test("formatUsageLine prints the four usage fields", () => {
  assert.equal(
    formatUsageLine({
      input_tokens: 3,
      cache_creation_input_tokens: 50128,
      cache_read_input_tokens: 0,
      output_tokens: 45,
    }),
    "[agent] usage input=3 cache_creation=50128 cache_read=0 output=45 (turn total)",
  );
  assert.equal(
    formatUsageLine(undefined),
    "[agent] usage input=0 cache_creation=0 cache_read=0 output=0 (turn total)",
  );
});
