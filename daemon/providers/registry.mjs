// Provider presets (P14), in three tiers:
//   subscription  a CLI the user is signed in to (Claude Agent SDK, Codex,
//                 Antigravity). Review turns and one-shot judgments.
//   local         Ollama on this machine. One-shot judgments only.
//   key           an API key the user enters (provider-keys.json or an
//                 environment variable). One-shot judgments only, and off
//                 until the user turns it on: the manuscript leaves the
//                 machine under that provider's terms.
// A review turn runs in one of two ways (turnMode): Claude calls the office
// tools itself; Codex is given no tools and proposes changes the daemon
// applies (proposal-turns.mjs). Antigravity is offered for nothing
// (probe.mjs says why). A key or local provider runs no review turn in this
// stage, so `roles` leaves "turn" out for them.
//
// policy: one line shown next to the provider in the task pane. "(unverified)"
// marks a statement taken from a search summary, not from the provider's own
// page (research note 261001, section 3).

export const PRESETS = Object.freeze([
  {
    id: "claude",
    tier: "subscription",
    label: "Claude (Agent SDK)",
    roles: ["turn", "oneshot"],
    // How a review turn runs. "agent": the model calls the office tools
    // itself. "proposal": the model is given no tools and returns a list of
    // changes the daemon checks and applies (proposal-turns.mjs).
    turnMode: "agent",
    defaultEnabled: true,
    policy: "Your Claude subscription terms apply.",
  },
  {
    id: "codex",
    tier: "subscription",
    label: "Codex CLI (ChatGPT sign-in)",
    roles: ["turn", "oneshot"],
    turnMode: "proposal",
    defaultEnabled: true,
    policy: "Your ChatGPT plan terms apply.",
    // A proposal turn (proposal.mjs): the most characters, system and prompt
    // together, that go out in one request; a longer one is refused, not
    // cut. Codex adds the user's own global instructions to every run: about
    // 25,000 input tokens for a prompt of a few lines (measured 2026-10-01).
    // 120,000 characters are at most about 120,000 tokens (Korean text, one
    // token per character at worst), which with that share and the reply
    // stays under a context of 200,000 tokens. The context of each model
    // was not measured; the number is a starting point.
    proposalBudget: 120_000,
  },
  {
    id: "antigravity",
    tier: "subscription",
    label: "Antigravity CLI (agy)",
    roles: ["turn", "oneshot"],
    defaultEnabled: true,
    policy: "Your Google account terms apply.",
  },
  {
    id: "ollama",
    tier: "local",
    label: "Ollama (local)",
    roles: ["oneshot"],
    defaultEnabled: true,
    baseUrl: "http://localhost:11434/v1",
    policy: "Runs on this computer; text never leaves this machine.",
    // Request fields (openai-compat.mjs). A thinking model (qwen3.5) spends
    // its whole output on thinking and leaves the answer empty: a judgment
    // is asked without thinking, and as a JSON object when JSON is expected.
    body: { reasoning_effort: "none" },
    jsonBody: { response_format: { type: "json_object" } },
  },
  {
    id: "openrouter",
    tier: "key",
    label: "OpenRouter (incl. DeepSeek)",
    roles: ["oneshot"],
    defaultEnabled: false,
    baseUrl: "https://openrouter.ai/api/v1",
    keyCheck: "/key",
    policy:
      "No logging by default; the final provider's policy applies. Turn off training providers in OpenRouter settings.",
  },
  {
    id: "deepseek",
    tier: "key",
    label: "DeepSeek",
    roles: ["oneshot"],
    defaultEnabled: false,
    baseUrl: "https://api.deepseek.com",
    policy: "Data stored in China; used for training unless you opt out (unverified).",
  },
  {
    id: "groq",
    tier: "key",
    label: "Groq",
    roles: ["oneshot"],
    defaultEnabled: false,
    baseUrl: "https://api.groq.com/openai/v1",
    policy: "Not used for training; kept up to 30 days for abuse checks (unverified).",
  },
  {
    id: "mistral",
    tier: "key",
    label: "Mistral",
    roles: ["oneshot"],
    defaultEnabled: false,
    baseUrl: "https://api.mistral.ai/v1",
    policy: "Free mode may train on inputs; turn it off in Admin Privacy.",
  },
  {
    id: "gemini",
    tier: "key",
    label: "Gemini API",
    roles: ["oneshot"],
    defaultEnabled: false,
    baseUrl: "https://generativelanguage.googleapis.com/v1beta/openai",
    policy: "Free tier inputs improve Google products and may be read by people.",
  },
  {
    id: "openai-compatible",
    tier: "key",
    label: "OpenAI-compatible (custom URL)",
    roles: ["oneshot"],
    defaultEnabled: false,
    baseUrl: null,
    policy: "The policy of the server you enter applies.",
  },
]);

export const presetById = (id) => PRESETS.find((p) => p.id === id) ?? null;
export const canRun = (preset, role) => !!preset?.roles.includes(role);
