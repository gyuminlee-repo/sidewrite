// Review turns run by a CLI provider instead of the Agent SDK session
// (P14): Codex (`codex exec --json`) or Antigravity (`agy -p`). The CLI gets
// the office tools from the daemon's own MCP endpoint (mcp-http.mjs), opened
// for this one turn, and the task pane gets the same events as in a Claude
// turn: assistant text, tool_use_announce, error, turn_complete; index.mjs
// then sends the review turn_done through endTurn(key, outcome).
//
// How a CLI is run, always:
//   Codex  `-s read-only`, `--ephemeral`, an empty temporary directory, the
//          prompt on stdin. The MCP token is in the child's environment
//          (CC_OFFICE_MCP_TOKEN, read through bearer_token_env_var): it is
//          in no argument and no file.
//   agy    `--sandbox`, an empty temporary directory, the prompt as one
//          argument. agy takes an MCP server only from its own settings and
//          does not expand a variable in a header (checked 2026-10-01), so
//          `agy mcp add` writes the route and its token to
//          ~/.gemini/config/mcp_config.json when the turn starts, and
//          `agy mcp remove office` takes them out when the turn ends and at
//          shutdown. Route and token are random per turn and dead once the
//          turn ends. One agy turn at a time: there is one such entry.
//   No flag that approves tool calls or turns a sandbox off, for either.
//
// Both CLIs refuse the office tools without an approval (probe.mjs
// TURN_BLOCKED), so today a turn here ends as failed with the CLI's own
// message and nothing is changed. The task pane does not offer them for
// review turns unless CC_OFFICE_CLI_TURNS lists them.

import { spawn } from "node:child_process";
import { randomBytes } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createInterface } from "node:readline";
import { runCli, resolveCliBin } from "./providers/cli-exec.mjs";
import { codexTurnArgs, parseCodexLine } from "./providers/codex.mjs";
import { agyTurnArgs, agyMcpAddArgs, parseAgyLine } from "./providers/agy.mjs";

const TURN_DIR_PREFIX = "mp-turn-";
const KILL_GRACE_MS = 5_000;

const CLIS = {
  codex: { label: "Codex CLI", bin: "codex", envVar: "CC_OFFICE_CODEX_BIN" },
  antigravity: { label: "agy", bin: "agy", envVar: "CC_OFFICE_AGY_BIN" },
};

async function defaultMkTurnDir() {
  return mkdtemp(join(tmpdir(), TURN_DIR_PREFIX));
}
// Only a directory this module made.
async function defaultRmTurnDir(dir) {
  if (!String(dir).startsWith(join(tmpdir(), TURN_DIR_PREFIX))) return;
  await rm(dir, { recursive: true, force: true });
}

export function createCliTurns({
  mcpHttp,
  baseUrl,
  sendText,
  sendEvent,
  endTurn,
  log = console.log,
  spawnFn = spawn,
  runCliFn = runCli,
  mkTurnDir = defaultMkTurnDir,
  rmTurnDir = defaultRmTurnDir,
  env = process.env,
  timeoutMs = Number(env.CC_OFFICE_CLI_TURN_TIMEOUT_MS) || 900_000,
}) {
  const live = new Map(); // pane key -> { id, child, stop(), done }
  const binOf = (id) => resolveCliBin(CLIS[id].bin, { env, envVar: CLIS[id].envVar });

  // One event of the CLI's stream, in the terms of the task pane. `st` keeps
  // what decides the outcome of the turn.
  function onEvent(ev, key, st) {
    if (!ev) return;
    if (ev.kind === "text") {
      sendText(st.id === "codex" ? `${ev.text}\n` : ev.text, key);
    } else if (ev.kind === "tool") {
      st.tools += 1;
      sendEvent({ event: "tool_use_announce", tool: ev.name, input: ev.input ?? {} }, key);
    } else if (ev.kind === "tool_error" && ev.approval) {
      st.failure ??= `${CLIS[st.id].label} refused the office tool call: ${ev.message}. No change was made by that call. Codex can be used for judgments only.`;
    } else if (ev.kind === "denied") {
      st.failure ??= `agy denied the office tool call (${ev.actions.join(", ")}). No change was made by that call. Antigravity can be used for judgments only.`;
    } else if (ev.kind === "error") {
      st.failure ??= ev.message;
    }
  }

  async function run(key, prompt, { id, model }, turn) {
    const cli = CLIS[id];
    const route = randomBytes(16).toString("hex");
    const token = randomBytes(32).toString("hex");
    const st = { id, tools: 0, failure: null, stopped: false, timedOut: false, code: null };
    const unbind = mcpHttp.bind(route, { token, key, ref: turn.ref });
    const mcpUrl = `${baseUrl}/mcp/${route}`;
    const bin = binOf(id);
    let dir = null;
    let registered = false;
    let finished = false;
    let resolveDone;
    const done = new Promise((r) => (resolveDone = r));
    const entry = { id, child: null, stop: () => {}, done };
    live.set(key, entry);

    async function finish() {
      if (finished) return;
      finished = true;
      unbind();
      if (registered) {
        const r = await runCliFn(bin, ["mcp", "remove", "office"], { timeoutMs: 30_000 });
        if (r.code !== 0) {
          log(
            `[cli-turn] agy mcp remove failed code=${r.code}: ${String(r.stderr ?? "").slice(0, 200)}`,
          );
        }
      }
      if (dir) {
        await rmTurnDir(dir).then(undefined, (e) =>
          log(`[cli-turn] could not remove ${dir}: ${e?.message ?? e}`),
        );
      }
      live.delete(key);
      log(`[cli-turn] done provider=${id} code=${st.code ?? "-"} tools=${st.tools}`);
      if (st.stopped) {
        sendEvent({ event: "turn_complete", interrupted: true }, key);
        endTurn(key, { ok: false, error: "The turn was stopped." });
      } else if (st.failure) {
        sendEvent({ event: "error", error: st.failure }, key);
        sendEvent({ event: "turn_complete", subtype: "error" }, key);
        endTurn(key, { ok: false, error: st.failure });
      } else {
        sendEvent({ event: "turn_complete", subtype: "success" }, key);
        endTurn(key, { ok: true });
      }
      resolveDone();
    }

    try {
      dir = await mkTurnDir();
      if (id === "antigravity") {
        const add = await runCliFn(bin, agyMcpAddArgs({ url: mcpUrl, token }), {
          timeoutMs: 30_000,
        });
        // The token was an argument of that command: never put its text in a log.
        if (add.code !== 0) {
          st.failure = `Could not register the office tools with agy: ${String(add.stderr ?? "")
            .split(token)
            .join("[token]")
            .slice(0, 300)}`;
          return await finish();
        }
        registered = true;
      }
      if (!bin.cmd) {
        st.failure = bin.error;
        return await finish();
      }
      const args =
        id === "codex"
          ? codexTurnArgs({ model, cwd: dir, mcpUrl })
          : agyTurnArgs({ prompt, model, timeoutS: Math.ceil(timeoutMs / 1000) });
      const child = spawnFn(bin.cmd, [...bin.prefixArgs, ...args], {
        cwd: dir,
        env: id === "codex" ? { ...env, CC_OFFICE_MCP_TOKEN: token } : env,
        shell: false,
        windowsHide: true,
        stdio: ["pipe", "pipe", "pipe"],
      });
      entry.child = child;
      log(`[cli-turn] start provider=${id} model=${model ?? "default"} ref=${turn.ref}`);

      let killTimer = null;
      const kill = () => {
        child.kill("SIGTERM");
        killTimer = setTimeout(() => child.kill("SIGKILL"), KILL_GRACE_MS);
        killTimer.unref?.();
      };
      const timer = setTimeout(() => {
        st.timedOut = true;
        st.failure ??= `${cli.label} timed out after ${timeoutMs} ms.`;
        kill();
      }, timeoutMs);
      entry.stop = () => {
        st.stopped = true;
        kill();
      };

      const stderr = [];
      child.stderr.on("data", (c) => stderr.length < 50 && stderr.push(String(c)));
      // The turn ends when the process has closed AND its last line was
      // read, in whichever order the two come.
      let exited = false;
      let exitCode = null;
      let linesDone = false;
      const maybeEnded = () => exited && linesDone && ended(exitCode);
      const rl = createInterface({ input: child.stdout });
      rl.on("line", (line) => {
        if (id === "codex") onEvent(parseCodexLine(line), key, st);
        else for (const ev of parseAgyLine(line)) onEvent(ev, key, st);
      });
      rl.on("close", () => {
        linesDone = true;
        maybeEnded();
      });
      child.stdin.on("error", () => undefined); // the CLI ended before reading its input
      child.stdin.end(id === "codex" ? prompt : undefined);

      const ended = (code) => {
        clearTimeout(timer);
        clearTimeout(killTimer);
        st.code = code;
        if (!st.stopped && !st.failure && code !== 0) {
          const detail = stderr.join("").trim().slice(0, 300);
          st.failure = `${cli.label} exited with code ${code}${detail ? `: ${detail}` : ""}`;
        }
        finish();
      };
      child.on("error", (e) => {
        st.failure ??=
          e?.code === "ENOENT"
            ? `${cli.label === "agy" ? "Antigravity CLI (agy)" : cli.label} is not installed`
            : String(e?.message ?? e);
        ended(null);
      });
      child.on("close", (code) => {
        exited = true;
        exitCode = code;
        maybeEnded();
      });
    } catch (e) {
      st.failure ??= String(e?.message ?? e);
      await finish();
    }
  }

  return {
    running: (key) => live.has(key),
    // Starts the turn; resolves once the CLI is running (or the turn has
    // already ended on a start failure, reported through the events).
    // Rejects only when nothing was started. turn: { ref }.
    async start(key, host, prompt, choice, turn) {
      if (!CLIS[choice?.id]) throw new Error(`${choice?.id} cannot run review turns.`);
      if (!turn?.ref) throw new Error("A CLI review turn needs the comment it is for.");
      if (live.has(key)) throw new Error("A CLI turn is already running in this pane.");
      if (choice.id === "antigravity" && [...live.values()].some((t) => t.id === "antigravity")) {
        throw new Error(
          "Antigravity runs one Antigravity turn at a time; another pane is using it.",
        );
      }
      await run(key, prompt, choice, turn);
    },
    // true when a turn of this pane was running and was told to stop.
    stop(key) {
      const t = live.get(key);
      if (!t?.child) return false;
      t.stop();
      return true;
    },
    // Daemon exit: stop every turn and wait until each has cleaned up (the
    // agy registration and the temporary directory).
    async shutdown() {
      const all = [...live.values()];
      for (const t of all) t.stop();
      await Promise.all(all.map((t) => t.done));
    },
  };
}
