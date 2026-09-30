import { test, mock } from "node:test";
import assert from "node:assert/strict";
import {
  resolveSettingSources,
  userMcpEnabled,
  buildQueryOptions,
  claudeExecutablePath,
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
  const warn = mock.method(console, "warn", () => {});
  try {
    assert.deepEqual(resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: "" }), []);
    assert.equal(warn.mock.callCount(), 0);
    assert.deepEqual(resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: "user, project,bogus" }), [
      "user",
      "project",
    ]);
  } finally {
    warn.mock.restore();
  }
});

test("settingSources values are case-insensitive, unknown names warn", () => {
  const warn = mock.method(console, "warn", () => {});
  try {
    assert.equal(resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: "ALL" }), undefined);
    assert.equal(resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: " All " }), undefined);
    assert.deepEqual(resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: "Project" }), ["project"]);
    assert.equal(warn.mock.callCount(), 0);
    assert.deepEqual(resolveSettingSources({ CC_OFFICE_SETTING_SOURCES: "projects" }), []);
    assert.equal(warn.mock.callCount(), 1);
    assert.match(warn.mock.calls[0].arguments[0], /projects/);
  } finally {
    warn.mock.restore();
  }
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

test("CC_OFFICE_CLAUDE_PATH sets pathToClaudeCodeExecutable, unset keeps SDK built-in", () => {
  assert.equal(claudeExecutablePath({}), undefined);
  assert.equal("pathToClaudeCodeExecutable" in buildQueryOptions(base, {}), false);
  const env = { CC_OFFICE_CLAUDE_PATH: " /home/u/.local/bin/claude " };
  assert.equal(
    buildQueryOptions(base, env).pathToClaudeCodeExecutable,
    "/home/u/.local/bin/claude",
  );
});
