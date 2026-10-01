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
// A tool call (tests/fixtures/codex_exec_turn_approval.jsonl) is an
// item.started and an item.completed of type "mcp_tool_call":
//   {"id":"item_3","type":"mcp_tool_call","server":"office","tool":"...",
//    "arguments":{...},"result":null,"error":null,"status":"in_progress"}
//   ... "error":{"message":"..."},"status":"failed"
// In that run Codex refused the call itself: "MCP tool call requires
// approval, but approval policy is never". The command_execution item was
// not seen in a capture.
//
// Every run is `-s read-only`, `--ephemeral` and never carries a flag that
// turns approvals or the sandbox off. The user's ~/.codex/config.toml is
// kept: without it the sign-in has no model provider and the run fails.
//
// What a run can reach, measured on 2026-10-01 (codex-cli 0.159.2, made-up
// text only; fixtures codex_exec_isolated.jsonl, codex_exec_tool_probe.jsonl,
// codex_request_tools.json):
//   With the arguments this file used until then (exec --json --ephemeral
//   --skip-git-repo-check -s read-only -C <dir>):
//   - the model ran shell commands (item type "command_execution": `echo
//     hi`, `ls -a`, a python one-liner that tried to open a URL; the
//     read-only sandbox refused the network, nothing else);
//   - the user's hooks ran (SessionStart, UserPromptSubmit, PreToolUse,
//     PostToolUse, Stop), one of them with the prompt text as an argument of
//     a process, and they wrote files into the run directory;
//   - the request to the model carried exec_command, write_stdin, the MCP
//     resource tools, request_plugin_install, view_image, apply_patch and
//     the goal tools; with model gpt-5.5 also web_search.
//   With the switches of ISOLATION below:
//   - no hook process started, and the same prompt got "NO_TOOL" for the
//     shell, the file listing and the web page (two real runs: one with the
//     switches the plan named, one with the whole list). With shell_snapshot
//     on, Codex itself still started the user's shell to take a snapshot of
//     it, which reads the user's shell startup file (the prompt is not part
//     of that); with the whole list no shell process started at all.
//   - a judgment through codexOneshot (two made-up sentences) answered in
//     8 s with the items error, error, agent_message and nothing else.
//   - the request to the model carries no shell, no MCP, no web search, no
//     image and no plugin tool. What Codex has no switch for stays: for the
//     default model `exec` (a JavaScript isolate; its own description says
//     "no Node, no file system, no network access"), wait,
//     request_user_input, apply_patch and clock__curr_time; for gpt-5.5
//     request_user_input, apply_patch and tool_search. Not measured, taken
//     from what the tools are: apply_patch writes a local file and the
//     sandbox is read-only; tool_search has nothing to find with plugins
//     and apps off. None of them is a way to the network.
//   - On this machine a local proxy of the user stands between Codex and
//     the model (the provider's base_url is 127.0.0.1). The request was read
//     where Codex hands it to that proxy. In the real run the model named
//     multi_agent tools the request does not carry; whether it made them up
//     or the proxy adds tools is not known. What the result rests on is the
//     behaviour: no tool item in the run, NO_TOOL for all three steps.
//   - `--disable unified_exec` does not change what `codex features list`
//     prints for a user config that sets it, but exec_command and
//     write_stdin are gone from the request all the same.
//   Not usable: --ignore-user-config (the run goes to the default provider
//   instead of the one the user set up and failed here); so the user's
//   AGENTS.md is still part of the prompt.
//   The environment PATH and HOME alone was enough (with that proxy
//   already running; Windows was not measured).
//   --output-schema <file> was accepted and the answer had the shape.
//   An item of type "reasoning" comes when the model prints a summary of
//   its thinking (seen with reasoning effort high); it is output, not a
//   tool (ALLOWED_ITEMS).
// The request body was read by pointing the provider's base_url (a -c
// override on the command line) at a local server that answers 400: no
// model call and no file of the user is touched. It was read for two models
// of one Codex version; a later version or another model may send more.
// That is why a run is also watched: the first item that is not an answer
// or a notice ends it and its reply is discarded. The watch is an alarm
// after the fact (the text is already with the CLI when an item shows); the
// wall is the list of switches.

import { parseJsonReply } from "../oneshot.mjs";
import { runCli as defaultRunCli, childEnv, resolveCliBin, withRunDir } from "./cli-exec.mjs";

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

// The switches of a judgment (see the measurements at the top): every
// feature of Codex that gives the model a tool or runs something of the
// user's, its exec rules, and the web search of the provider.
const ISOLATION_OFF = [
  "plugins", // the user's MCP servers come from plugins
  "apps",
  "hooks",
  "shell_tool",
  "unified_exec",
  "browser_use",
  "computer_use",
  "multi_agent",
  "image_generation",
  "skill_search",
  "tool_suggest",
  "view_image",
  "memories",
  "goals",
  "sleep_tool",
  "shell_snapshot",
];
const ISOLATION = [
  ...ISOLATION_OFF.flatMap((f) => ["--disable", f]),
  "--ignore-rules",
  "-c",
  'web_search="disabled"',
];

export const codexOneshotArgs = ({ model, cwd }) => [
  ...baseArgs({ model, cwd }),
  ...ISOLATION,
  "-",
];

// The kinds of item `codex exec --json` has, read on 2026-10-01 from the
// installed binary (codex-cli 0.159.3; the names stand together in its data,
// after item.started, item.updated, item.completed):
//   agent_message, reasoning, command_execution, file_change, mcp_tool_call,
//   web_search, todo_list   (and "error", the notice seen in every run)
// A judgment may have the ones that are output of the model itself:
//   agent_message  the answer
//   error          a notice (a config warning comes that way)
//   reasoning      a summary of its thinking (fixture
//                  codex_exec_reasoning.jsonl, a real run)
//   todo_list      the list of steps it keeps for itself; nothing is run and
//                  nothing leaves the process (not seen in a run)
// Everything else ends the run: the tools (command_execution, file_change,
// mcp_tool_call, web_search) and any kind this list does not name, such as
// one a later version adds.
const ALLOWED_ITEMS = new Set(["agent_message", "error", "reasoning", "todo_list"]);

// The name of the tool a stdout line shows, or null. Only names Codex
// itself gives: no argument and no text of the run.
export function toolEventOf(line) {
  let o;
  try {
    o = JSON.parse(line);
  } catch {
    return null;
  }
  if (!o || typeof o !== "object" || !String(o.type ?? "").startsWith("item.")) return null;
  const it = o.item ?? {};
  if (ALLOWED_ITEMS.has(it.type)) return null;
  if (it.type === "mcp_tool_call") return `mcp__${it.server}__${it.tool}`;
  if (it.type === "command_execution") return "codex_shell";
  return /^[a-z0-9_]{1,40}$/i.test(String(it.type ?? "")) ? String(it.type) : "unknown item";
}

// One stdout line as { kind: "session" | "text" | "tool" | "tool_done" |
// "tool_error" | "notice" | "done" | "error", ... }, or null for a line that
// says nothing we use.
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
      if (it.type === "mcp_tool_call") {
        const name = `mcp__${it.server}__${it.tool}`;
        if (it.status !== "failed" && !it.error) return { kind: "tool_done", name };
        const message = String(it.error?.message ?? "tool call failed");
        // approval: Codex would not run the tool without the user's approval.
        return { kind: "tool_error", name, message, approval: /requires approval/i.test(message) };
      }
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
//
// The run has a directory of its own, an environment without keys, and the
// tools of Codex switched off. It is watched as it runs: the first item that
// is a tool ends it, and the reply is discarded ({ ok: false, tool_event }).
// signal ends it from outside ({ ok: false, aborted: true }).
export async function codexOneshot(
  { system, prompt, model, expectJson = true, timeoutMs = 120_000, signal = null },
  {
    runCli = defaultRunCli,
    bin = resolveCliBin("codex", { envVar: "CC_OFFICE_CODEX_BIN" }),
    runDir = withRunDir,
    env = process.env,
    log = console.log,
  } = {},
) {
  let r;
  let tool = null;
  try {
    r = await runDir((cwd) =>
      runCli(bin, codexOneshotArgs({ model, cwd }), {
        stdin: `${system}\n\n${prompt}`,
        timeoutMs,
        cwd,
        env: childEnv(env),
        signal,
        onLine(line, kill) {
          tool ??= toolEventOf(line);
          if (tool) kill();
        },
      }),
    );
  } catch (e) {
    return { ok: false, error: `codex could not be run: ${String(e?.message ?? e)}` };
  }
  let text = null;
  let failed = null;
  for (const line of String(r.stdout ?? "").split("\n")) {
    tool ??= toolEventOf(line); // a runner that did not hand the lines over
    const ev = parseCodexLine(line);
    if (ev?.kind === "text") text = ev.text;
    if (ev?.kind === "error") failed = ev;
  }
  if (tool) {
    log(`[providers] tool event provider=codex name=${tool}`);
    return {
      ok: false,
      tool_event: tool,
      error: `Codex CLI used a tool (${tool}) in a run that must not use any. Its reply was discarded.`,
    };
  }
  if (r.aborted) return { ok: false, aborted: true, error: "stopped" };
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
