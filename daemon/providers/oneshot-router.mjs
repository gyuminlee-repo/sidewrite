// One-shot judgments go to the provider the pane chose (P14).
//
// getChoice(key) is the pane's choice for the "oneshot" role: { id, model },
// or null for the default (Claude, haiku). It may be async, and it may throw
// when the pane has no confirmed choice (selection.mjs ready()).
//
// pin(key, { claudeModel }?) resolves the choice ONCE and gives { provider, model, run }: the
// handle of one job (a split with its requirement extraction, a ledger check
// with all its batches). Every run(req) of the handle goes to that provider
// and model, and `model` is the label its decision receipts carry
// ("<provider>:<model>", or the bare alias for Claude as before this stage).
// A change of the choice in the pane applies to the next pin. run has the
// contract of oneshot.mjs runOneshot ({ ok, text, json?, error?, limited? },
// never throws). forKey(key) is a single request: it pins and runs.
//
// Nothing here falls back to another provider: a choice that cannot be
// resolved, or a provider that cannot be used (no key, no server URL,
// unknown id), is an error result, so the text of a manuscript never goes
// somewhere the user did not choose.

import { presetById } from "./registry.mjs";

const CLAUDE_DEFAULT_MODEL = "haiku";

export function createOneshotRouter({
  claude,
  codex,
  antigravity,
  http,
  keys,
  getChoice,
  log = () => {},
}) {
  async function send(choice, req, claudeModel) {
    if (!choice || choice.id === "claude") {
      return claude({ ...req, model: choice?.model ?? claudeModel });
    }
    const preset = presetById(choice.id);
    if (!preset) return { ok: false, error: `Unknown provider ${choice.id}` };
    if (preset.id === "codex") return codex({ ...req, model: choice.model });
    if (preset.id === "antigravity") return antigravity({ ...req, model: choice.model });
    if (preset.tier === "local") {
      return http({
        ...req,
        baseUrl: preset.baseUrl,
        apiKey: null,
        model: choice.model,
        ...(preset.body && { body: preset.body }),
        ...(preset.jsonBody && { jsonBody: preset.jsonBody }),
      });
    }
    const stored = await keys.get(preset.id);
    if (!stored?.key) {
      return {
        ok: false,
        error: `No API key for ${preset.label}. Enter one in Setup, Providers.`,
      };
    }
    const baseUrl = preset.baseUrl ?? stored.base_url;
    if (!baseUrl) return { ok: false, error: `No server URL for ${preset.label}.` };
    return http({
      ...req,
      baseUrl,
      apiKey: stored.key,
      model: choice.model,
      ...(preset.body && { body: preset.body }),
      ...(preset.jsonBody && { jsonBody: preset.jsonBody }),
    });
  }

  // claudeModel: the Claude model of this job when the pane's choice names
  // none (a response draft asks for a stronger one than a judgment). It is a
  // default only: a model or another provider chosen in the pane stays.
  async function pin(key, { claudeModel = CLAUDE_DEFAULT_MODEL } = {}) {
    let choice;
    try {
      choice = (await getChoice(key)) ?? null;
    } catch (e) {
      // No confirmed choice: every request of this job is an error, and
      // nothing is sent to any provider.
      const error = String(e?.message ?? e);
      log(`[providers] judgment refused: ${error}`);
      return { provider: null, model: "none", run: async () => ({ ok: false, error }) };
    }
    const provider = choice?.id ?? "claude";
    const model =
      provider === "claude"
        ? (choice?.model ?? claudeModel)
        : `${provider}:${choice.model ?? "default"}`;
    return {
      provider,
      model,
      run(req) {
        // One line per request, with no text of it.
        log(`[providers] judgment provider=${provider} model=${model}`);
        return send(choice, req, claudeModel).then(
          (r) => r,
          (e) => ({ ok: false, error: String(e?.message ?? e) }),
        );
      },
    };
  }

  return {
    pin,
    forKey: (key) => (req) => pin(key).then((h) => h.run(req)),
  };
}
