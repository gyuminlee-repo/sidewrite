// Agent SDK query() options for one pane's session, assembled from plain
// arguments plus an env object so node --test can exercise it without a
// running daemon. Environment variables (all optional):
//
//   CC_OFFICE_SETTING_SOURCES  unset: ["project"] (workspace CLAUDE.md and
//                              .claude/settings.json only). "all": key
//                              omitted, the SDK loads user+project+local
//                              (upstream behavior). "": [] (isolation).
//                              Otherwise a comma list of user,project,local.
//   CC_OFFICE_LOAD_USER_MCP    "1": merge ~/.claude.json mcpServers into the
//                              session (upstream behavior). Default off.
//   CC_OFFICE_CLAUDE_PATH      absolute path to a local claude executable
//                              (pathToClaudeCodeExecutable). Default: the
//                              SDK's bundled CLI.

const VALID_SOURCES = new Set(["user", "project", "local"]);

export function resolveSettingSources(env = process.env) {
  const raw = env.CC_OFFICE_SETTING_SOURCES;
  if (raw === undefined) return ["project"];
  if (raw.trim() === "all") return undefined;
  return raw
    .split(",")
    .map((s) => s.trim())
    .filter((s) => VALID_SOURCES.has(s));
}

export function userMcpEnabled(env = process.env) {
  return env.CC_OFFICE_LOAD_USER_MCP === "1";
}

export function claudeExecutablePath(env = process.env) {
  const p = env.CC_OFFICE_CLAUDE_PATH?.trim();
  return p ? p : undefined;
}

export function buildQueryOptions(
  {
    cwd,
    append,
    officeMcp,
    userMcpServers = {},
    disallowedTools,
    canUseTool,
    model,
    abortController,
    stderr,
    resume,
  },
  env = process.env,
) {
  const settingSources = resolveSettingSources(env);
  const executable = claudeExecutablePath(env);
  return {
    cwd,
    systemPrompt: { type: "preset", preset: "claude_code", append },
    mcpServers: userMcpEnabled(env)
      ? { ...userMcpServers, office: officeMcp }
      : { office: officeMcp },
    disallowedTools,
    canUseTool,
    includePartialMessages: true,
    model,
    abortController,
    stderr,
    ...(settingSources ? { settingSources } : {}),
    ...(executable ? { pathToClaudeCodeExecutable: executable } : {}),
    ...(resume ? { resume } : {}),
  };
}

// result.usage is the whole turn (every API call in it). The P2 pass metric
// is different: the first assistant call's usage in the session .jsonl
// (scripts/probe-hi.mjs).
export function formatUsageLine(usage) {
  const u = usage ?? {};
  return (
    `[agent] usage input=${u.input_tokens ?? 0}` +
    ` cache_creation=${u.cache_creation_input_tokens ?? 0}` +
    ` cache_read=${u.cache_read_input_tokens ?? 0}` +
    ` output=${u.output_tokens ?? 0} (turn total)`
  );
}
