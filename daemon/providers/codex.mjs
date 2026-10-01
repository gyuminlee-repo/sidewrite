// Codex CLI adapter (P14): argument lists, the JSONL events of
// `codex exec --json`, and one-shot judgments.
//
// Event shapes are those of codex-cli 0.159.2, captured on 2026-10-01
// (tests/fixtures/codex_exec_oneshot.jsonl, codex_exec_failed.jsonl):
//   {"type":"thread.started","thread_id":"..."}
//   {"type":"turn.started"}
//   {"type":"item.completed","item":{"id":"item_0","type":"error","message":"..."}}
//   {"type":"item.completed","item":{"id":"item_2","type":"agent_message","text":"..."}}
//   {"type":"turn.completed","usage":{"input_tokens":..,"output_tokens":..}}
//   {"type":"error","message":"..."}  then  {"type":"turn.failed","error":{"message":"..."}}
// An item of type "error" is a notice (a config warning came that way in a
// run that answered); only turn.failed and a top-level error end the turn.
// The mcp_tool_call and command_execution items were not seen in the capture
// (no tool ran); their fields here follow the plan and wait for Task 12.
//
// Every run is `-s read-only`, `--ephemeral` and never carries a flag that
// turns approvals or the sandbox off. The user's ~/.codex/config.toml is
// kept: without it the sign-in has no model provider and the run fails.

import { parseJsonReply } from "../oneshot.mjs";
import { runCli as defaultRunCli, emptyWorkDir } from "./cli-exec.mjs";

const LIMIT_RE = /(usage limit|rate limit|quota|too many requests|out of credits|\b429\b)/i;
const toml = (s) => JSON.stringify(String(s));

function baseArgs({ model, cwd }) {
  return [
    "exec",
    "--json",
    "--ephemeral",
    "--skip-git-repo-check",
    "-s",
    "read-only",
    "-C",
    cwd,
    ...(model ? ["-m", model] : []),
  ];
}

// A review turn: the office tools come from the daemon's MCP endpoint of
// this turn (mcpUrl); its token is read from the environment of the child
// (CC_OFFICE_MCP_TOKEN), so it is in no argument and no config file. The
// prompt comes on stdin ("-").
export function codexTurnArgs({ model, cwd, mcpUrl }) {
  return [
    ...baseArgs({ model, cwd }),
    "-c",
    `mcp_servers.office.url=${toml(mcpUrl)}`,
    "-c",
    'mcp_servers.office.bearer_token_env_var="CC_OFFICE_MCP_TOKEN"',
    "-",
  ];
}

export const codexOneshotArgs = ({ model, cwd }) => [...baseArgs({ model, cwd }), "-"];

// One stdout line as { kind: "session" | "text" | "tool" | "notice" | "done"
// | "error", ... }, or null for a line that says nothing we use.
export function parseCodexLine(line) {
  let o;
  try {
    o = JSON.parse(line);
  } catch {
    return null;
  }
  if (!o || typeof o !== "object") return null;
  const it = o.item ?? {};
  switch (o.type) {
    case "thread.started":
      return { kind: "session", id: o.thread_id ?? null };
    case "item.started":
      if (it.type === "mcp_tool_call") {
        return {
          kind: "tool",
          name: `mcp__${it.server}__${it.tool}`,
          input: it.arguments ?? {},
        };
      }
      if (it.type === "command_execution") {
        return { kind: "tool", name: "codex_shell", input: { command: it.command } };
      }
      return null;
    case "item.completed":
      if (it.type === "agent_message") return { kind: "text", text: String(it.text ?? "") };
      if (it.type === "error") return { kind: "notice", message: String(it.message ?? "") };
      return null;
    case "turn.completed":
      return { kind: "done", usage: o.usage ?? null };
    case "turn.failed":
    case "error": {
      const message = String(o.error?.message ?? o.message ?? "codex failed");
      return { kind: "error", message, limited: LIMIT_RE.test(message) };
    }
    default:
      return null;
  }
}

// The contract of oneshot.mjs runOneshot: { ok, text, json?, error?,
// limited? }; never throws. System and prompt go as one stdin text (the CLI
// has no separate system prompt).
export async function codexOneshot(
  { system, prompt, model, expectJson = true, timeoutMs = 120_000 },
  {
    runCli = defaultRunCli,
    bin = process.env.CC_OFFICE_CODEX_BIN || "codex",
    emptyDir = emptyWorkDir,
  } = {},
) {
  let r;
  try {
    r = await runCli(bin, codexOneshotArgs({ model, cwd: await emptyDir() }), {
      stdin: `${system}\n\n${prompt}`,
      timeoutMs,
    });
  } catch (e) {
    return { ok: false, error: `codex could not be run: ${String(e?.message ?? e)}` };
  }
  let text = null;
  let failed = null;
  for (const line of String(r.stdout ?? "").split("\n")) {
    const ev = parseCodexLine(line);
    if (ev?.kind === "text") text = ev.text;
    if (ev?.kind === "error") failed = ev;
  }
  if (failed) {
    return {
      ok: false,
      error: failed.message,
      limited: failed.limited ? { limited: true, kind: "rate_limit", text: failed.message } : null,
    };
  }
  if (r.timedOut) return { ok: false, error: `timed out after ${timeoutMs} ms` };
  if (r.code !== 0) {
    const detail = String(r.stderr ?? "").slice(0, 300);
    return detail === "not installed"
      ? { ok: false, error: "Codex CLI is not installed" }
      : { ok: false, error: `codex exited ${r.code}: ${detail}` };
  }
  if (text === null) return { ok: false, error: "codex gave no answer" };
  if (!expectJson) return { ok: true, text };
  const json = parseJsonReply(text);
  return json === undefined
    ? { ok: false, error: "reply was not JSON", text }
    : { ok: true, text, json };
}
