import { test } from "node:test";
import assert from "node:assert/strict";
import {
  providerRows,
  choiceAfterProbe,
  providerStatus,
  modelRows,
  proposalNotice,
  PROPOSAL_TITLE,
} from "../taskpane/shared/providers-view.js";

const LIST = [
  {
    id: "claude",
    tier: "subscription",
    label: "Claude",
    roles: ["turn", "oneshot"],
    available: true,
    turn_ready: true,
    turn_note: null,
    reason: null,
    policy: "p",
    models: [{ value: "sonnet", displayName: "Sonnet 5.5" }],
  },
  {
    id: "antigravity",
    tier: "subscription",
    label: "agy",
    roles: ["turn", "oneshot"],
    available: true,
    turn_ready: false,
    turn_note: "Review turns are off for it in this test.",
    reason: null,
    policy: "p",
    models: [{ value: "g", displayName: "G" }],
  },
  {
    id: "ollama",
    tier: "local",
    label: "Ollama (local)",
    roles: ["oneshot"],
    available: true,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider in this stage.",
    reason: null,
    policy: "never leaves this machine",
    models: [{ value: "qwen3.5:9b", displayName: "qwen3.5:9b" }],
  },
  {
    id: "deepseek",
    tier: "key",
    label: "DeepSeek",
    roles: ["oneshot"],
    available: false,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider in this stage.",
    reason: "Connection failed (ECONNRESET)",
    policy: "China",
    has_key: true,
    key_source: "file",
    default_enabled: false,
    models: [],
  },
  {
    id: "groq",
    tier: "key",
    label: "Groq",
    roles: ["oneshot"],
    available: false,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider in this stage.",
    reason: "No API key",
    policy: "30 days",
    has_key: false,
    key_source: null,
    default_enabled: false,
    models: [],
  },
];

test("turn menu: the CLI and Claude providers, only a turn-ready one enabled; judgment menu: available only", () => {
  const t = providerRows(LIST, "turn");
  assert.deepEqual(
    t.map((r) => r.id),
    ["claude", "antigravity"],
  );
  assert.deepEqual(
    t.filter((r) => r.enabled).map((r) => r.id),
    ["claude"],
  );
  assert.equal(
    t.find((r) => r.id === "antigravity").note,
    "Review turns are off for it in this test.",
  );
  assert.equal(
    t.some((r) => r.id === "deepseek"),
    false,
  );
  const o = providerRows(LIST, "oneshot");
  assert.deepEqual(
    o.filter((r) => r.enabled).map((r) => r.id),
    ["claude", "antigravity", "ollama"],
  );
  assert.match(o.find((r) => r.id === "deepseek").note, /Connection failed/);
  assert.equal(o.find((r) => r.id === "deepseek").policy, "China");
  assert.equal(o.find((r) => r.id === "antigravity").note, "", "fine for a judgment");
  // What the option shows.
  assert.equal(
    o.find((r) => r.id === "deepseek").text,
    "DeepSeek (Connection failed (ECONNRESET))",
  );
  assert.equal(o.find((r) => r.id === "ollama").text, "Ollama (local)");
  // A provider that cannot be chosen says why, in full.
  assert.equal(
    t.find((r) => r.id === "antigravity").text,
    "agy (Review turns are off for it in this test.)",
  );
});

test("a stored choice that stopped passing falls back to Claude", () => {
  assert.deepEqual(choiceAfterProbe({ id: "deepseek", model: "x" }, LIST, "oneshot"), {
    id: "claude",
    model: null,
    dropped: "deepseek",
  });
  assert.deepEqual(choiceAfterProbe({ id: "antigravity", model: "g" }, LIST, "oneshot"), {
    id: "antigravity",
    model: "g",
  });
  assert.deepEqual(choiceAfterProbe({ id: "antigravity", model: "g" }, LIST, "turn"), {
    id: "claude",
    model: null,
    dropped: "antigravity",
  });
  assert.deepEqual(choiceAfterProbe(null, LIST, "oneshot"), { id: "claude", model: null });
  assert.deepEqual(choiceAfterProbe({ id: "gone" }, LIST, "oneshot"), {
    id: "claude",
    model: null,
    dropped: "gone",
  });
  // A stored model the provider no longer lists is dropped, the provider kept.
  assert.deepEqual(choiceAfterProbe({ id: "ollama", model: "old:1b" }, LIST, "oneshot"), {
    id: "ollama",
    model: "qwen3.5:9b",
  });
});

test("the status line of a provider row", () => {
  const by = Object.fromEntries(LIST.map((r) => [r.id, providerStatus(r)]));
  assert.deepEqual(by.claude, {
    state: "ok",
    text: "Available for review turns and judgments.",
    keyNote: null,
  });
  assert.equal(by.antigravity.state, "ok");
  assert.equal(
    by.antigravity.text,
    "Available for judgments. Review turns are off for it in this test.",
  );
  assert.equal(by.ollama.text, "Available for judgments.");
  assert.deepEqual(by.deepseek, {
    state: "off",
    text: "Not available: Connection failed (ECONNRESET)",
    keyNote: "A key is stored in the key file.",
  });
  assert.equal(by.groq.keyNote, "Off until you enter a key.");
  assert.equal(
    providerStatus({ ...LIST[4], has_key: true, key_source: "env" }).keyNote,
    "A key comes from the environment.",
  );
});

test("model rows: the provider's list, the current one selected, a first row for none", () => {
  assert.deepEqual(modelRows(LIST[2], "qwen3.5:9b"), [
    { value: "qwen3.5:9b", label: "qwen3.5:9b", selected: true },
  ]);
  assert.deepEqual(modelRows(LIST[0], null), [
    { value: "", label: "Default model", selected: true },
    { value: "sonnet", label: "Sonnet 5.5", selected: false },
  ]);
  assert.deepEqual(modelRows(LIST[3], null), [
    { value: "", label: "No model listed", selected: true },
  ]);
  assert.deepEqual(modelRows(undefined, null), []);
});

// A provider that runs review turns by proposals (Codex): the menu and the
// list say so, and the first choice of it shows what leaves the computer.
const CODEX = {
  id: "codex",
  tier: "subscription",
  label: "Codex CLI (ChatGPT sign-in)",
  roles: ["turn", "oneshot"],
  available: true,
  turn_ready: true,
  turn_mode: "proposal",
  turn_note: null,
  reason: null,
  policy: "Your ChatGPT plan terms apply.",
  models: [{ value: "gpt-6.1-sol", displayName: "GPT 6.1 Sol" }],
};

test("a proposing provider: 'proposals only' in the turn menu and the status, plain in the judgment menu", () => {
  const list = [LIST[0], CODEX];
  const turn = providerRows(list, "turn").find((r) => r.id === "codex");
  assert.deepEqual(
    [turn.text, turn.enabled, turn.note, turn.proposal],
    ["Codex CLI (ChatGPT sign-in) (proposals only)", true, PROPOSAL_TITLE, true],
  );
  assert.equal(
    PROPOSAL_TITLE,
    "Returns a list of proposed changes. This add-in checks each one and applies it as a tracked change with a rationale comment. The CLI does not touch Word.",
  );
  const claude = providerRows(list, "turn").find((r) => r.id === "claude");
  assert.deepEqual([claude.text, claude.proposal], ["Claude", false]);
  const judge = providerRows(list, "oneshot").find((r) => r.id === "codex");
  assert.deepEqual([judge.text, judge.proposal], ["Codex CLI (ChatGPT sign-in)", false]);
  assert.equal(
    providerStatus(CODEX).text,
    "Available for review turns (proposals only) and judgments.",
  );
  // Not signed in: as any provider that did not pass its check.
  const off = {
    ...CODEX,
    available: false,
    turn_ready: false,
    reason: "Not signed in (run: codex login)",
  };
  assert.equal(
    providerRows([off], "turn")[0].text,
    "Codex CLI (ChatGPT sign-in) (Not signed in (run: codex login))",
  );
});

test("the notice before the first proposal turn says what is sent, what comes along and who edits", () => {
  const n = proposalNotice(CODEX);
  assert.equal(n.use, "Use Codex CLI (ChatGPT sign-in)");
  assert.equal(
    n.text,
    "Review turns with Codex CLI (ChatGPT sign-in) send text off this computer, to the CLI and the service behind it: " +
      "the reviewer comment as written, the manuscript paragraphs shown for it (your selection, or the whole manuscript), " +
      "your author instructions and the forbidden list. The CLI adds its own global instructions (for Codex, your AGENTS.md) to every request. " +
      "Your ChatGPT plan terms apply. " +
      "It does not edit the document: it returns proposed changes, and this add-in checks each one and applies it as a tracked change with a rationale comment. " +
      "Preview prompt in the Review tab shows what is sent.",
  );
});
