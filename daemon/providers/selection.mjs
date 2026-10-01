// Which provider a pane uses for each role (P14):
//   turn     review turns. Default Claude; its model is the composer's.
//   writing  prose the author reads or sends: response drafts, translations,
//            questions about a selection. Default null: Claude with the best
//            model of the account (oneshot-router.mjs claudeDefault).
//   oneshot  judgments (split, requirements, ledger checks and suggestions).
//            Default null: Claude with its judgment model (oneshot-router.mjs).
//
// The daemon never guesses a pane's provider. Each role of each pane has a
// state:
//   restoring  nothing confirmed yet in this connection: a pane that just
//              said hello, a pane that never sends its choice (an old build),
//              or a stored choice still being checked against the probe
//   ready      the choice is confirmed: restored, chosen by hand (set), or
//              the default because the pane said it has nothing stored
//   refused    the stored choice did not pass; `error` says why
// ready(key, role) is what a judgment or a review turn calls before it sends
// anything: it waits while the state is restoring (at most waitMs), gives
// the choice when ready, and throws otherwise. A caller that gets the throw
// sends nothing: there is no fallback to another provider, so the text of a
// manuscript does not go to Claude because a stored Ollama choice was late
// (Codex review of stage 2.5).
//
// A choice is accepted only for a provider whose probe row passed
// (probe.mjs): available for a judgment, turn_ready for a review turn. A
// refusal carries the row's own reason, which the task pane shows.

const ROLES = ["turn", "writing", "oneshot"];
const WORDS = { turn: "review turns", writing: "writing", oneshot: "judgments" };
const DEFAULT = { turn: { id: "claude", model: null }, writing: null, oneshot: null };
const MODEL_NAME = /^[\w.:/@[\]-]{1,120}$/;
// Longer than the probe's own deadline (probe.mjs DEADLINE_MS, 35 s).
const WAIT_MS = 45_000;
const ADVICE = "Nothing was sent to any provider. Choose one in Setup, Providers.";
// Providers whose request cannot be made without a model name.
const needsModel = (row) => row.tier === "local" || row.tier === "key";

// Row fields the task pane gets. A row never holds a key (probe.mjs); the
// list is closed all the same, so a field added later is not sent unseen.
const ROW_FIELDS = [
  "id",
  "tier",
  "label",
  "roles",
  "policy",
  "default_enabled",
  "available",
  "turn_ready",
  "turn_note",
  "reason",
  "models",
  "has_key",
  "key_source",
  "base_url",
];

export function providersListReply(rows, selection) {
  return {
    providers: (rows ?? []).map((r) =>
      Object.fromEntries(ROW_FIELDS.filter((k) => k in r).map((k) => [k, r[k]])),
    ),
    selection,
  };
}

// { id, model } as it is stored, or a throw that says why it cannot be used.
function checked(role, { id, model = null } = {}, rows) {
  if (!ROLES.includes(role)) throw new Error(`Unknown role ${role}`);
  const row = (rows ?? []).find((r) => r.id === id);
  if (!row) throw new Error(`Unknown provider ${id}`);
  if (!row.available) {
    throw new Error(`${row.label} is not available: ${row.reason ?? "it did not pass its check"}`);
  }
  if (role === "turn" && !row.turn_ready) {
    throw new Error(
      `${row.label} cannot run review turns. ${row.turn_note ?? row.reason ?? ""}`.trim(),
    );
  }
  const name = model == null || model === "" ? null : String(model);
  const listed = (row.models ?? []).map((m) => m.value);
  if (name && listed.length && !listed.includes(name)) {
    throw new Error(`${row.label} does not list the model ${name}`);
  }
  if (name && !MODEL_NAME.test(name)) throw new Error(`${name} is not a model name`);
  if (!name && role !== "turn" && needsModel(row)) {
    throw new Error(`Choose a model for ${row.label}`);
  }
  return { id, model: name };
}

export function createProviderSelection({ waitMs = WAIT_MS } = {}) {
  const byKey = new Map(); // pane key -> { turn: slot, writing: slot, oneshot: slot }
  const fresh = () => ({ state: "restoring", choice: null, error: null, gen: 0, waiters: [] });
  function slots(key) {
    if (!byKey.has(key)) byKey.set(key, Object.fromEntries(ROLES.map((r) => [r, fresh()])));
    return byKey.get(key);
  }
  function settle(slot, state, choice, error = null) {
    slot.gen += 1;
    Object.assign(slot, { state, choice, error });
    for (const wake of slot.waiters.splice(0)) wake();
  }
  function get(key) {
    const s = slots(key);
    const of = (f) => Object.fromEntries(ROLES.map((r) => [r, f(s[r])]));
    return {
      ...of((x) => (x.state === "ready" ? x.choice : null)),
      state: of((x) => x.state),
      error: of((x) => x.error),
    };
  }

  return {
    get,
    // The confirmed choice, for a label; the role's default while none is.
    // What a request uses is ready().
    choiceFor(key, role) {
      const s = slots(key)[role];
      return s.state === "ready" ? s.choice : DEFAULT[role];
    },
    drop: (key) => byKey.delete(key),
    // A pane connected (hello): whatever an earlier connection confirmed is
    // not this one's. Requests that already wait keep waiting.
    begin(key) {
      for (const role of ROLES) {
        const s = slots(key)[role];
        s.gen += 1;
        Object.assign(s, { state: "restoring", choice: null, error: null });
      }
    },
    // The pane's stored choices, sent right after hello: { turn, writing, oneshot },
    // each { id, model } or null (nothing stored: the default applies at
    // once). A stored one is checked against the probe rows (getRows may be
    // slow); a choice made by hand meanwhile wins. Returns get(key).
    async restore(key, stored, getRows) {
      const s = slots(key);
      const pending = [];
      for (const role of ROLES) {
        const want = stored?.[role] ?? null;
        if (!want) settle(s[role], "ready", DEFAULT[role]);
        else pending.push([role, want, (s[role].gen += 1)]);
      }
      if (pending.length) {
        let rows = null;
        let failed = null;
        try {
          rows = await getRows();
        } catch (e) {
          failed = `The providers could not be checked: ${e?.message ?? e}`;
        }
        for (const [role, want, gen] of pending) {
          if (s[role].gen !== gen) continue; // set by hand, or a newer connection
          if (failed) {
            settle(s[role], "refused", null, failed);
            continue;
          }
          try {
            settle(s[role], "ready", checked(role, want, rows));
          } catch (e) {
            settle(s[role], "refused", null, String(e?.message ?? e));
          }
        }
      }
      return get(key);
    },
    // rows: the last probe (probe.mjs probeAll). Returns get(key).
    set(key, role, choice, rows) {
      const ok = checked(role, choice, rows);
      settle(slots(key)[role], "ready", ok);
      return get(key);
    },
    // The confirmed choice of a role, waiting for it at most waitMs. Throws
    // when there is none: the caller must not send anything anywhere.
    async ready(key, role) {
      const s = slots(key)[role];
      if (s.state === "restoring") {
        await new Promise((resolve) => {
          const timer = setTimeout(done, waitMs);
          function done() {
            clearTimeout(timer);
            const at = s.waiters.indexOf(done);
            if (at >= 0) s.waiters.splice(at, 1);
            resolve();
          }
          s.waiters.push(done);
        });
      }
      if (s.state === "ready") return s.choice;
      if (s.state === "refused") {
        const why = String(s.error ?? "it was refused").replace(/\.$/, "");
        throw new Error(`The stored provider for ${WORDS[role]} cannot be used: ${why}. ${ADVICE}`);
      }
      throw new Error(
        `The task pane has not confirmed its provider for ${WORDS[role]} (waited ${waitMs / 1000} s). ${ADVICE}`,
      );
    },
  };
}
