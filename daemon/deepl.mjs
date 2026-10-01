// DeepL translation of the selected text (P8). Off unless the caller passes
// enabled: true (the user's consent switch) and a key. fetchFn is injected.
// Never throws; returns { ok, text?, error?, route: "deepl" }. Neither the
// key nor any text is logged or put into an error message.
//
// Only DeepL API Free is used: a key that does not end in ":fx" is refused
// and no request is made.
//
// Checked against the DeepL docs: the free host, the "DeepL-Auth-Key" header,
// the ":fx" suffix of a free key and the JSON body. Not checked: what 403,
// 429 and 456 mean exactly, and the terms of the free plan. So any message
// the service sends is shown as it came.
//
// Not wired yet. The caller reads the switch from the settings, the key
// from the key store, and asks translationRoute() before it translates. A
// DeepL failure is shown to the user and is not retried on another route.

export const DEEPL_NOTICE =
  "This feature sends the selected text from this PC to DeepL (deepl.com). Check DeepL's free-plan terms yourself before you turn it on.";

export const DEEPL_FREE_ONLY = "Only DeepL API Free keys are supported (the key ends in ':fx').";

const FREE = "https://api-free.deepl.com";

// A DeepL API Free key ends in ":fx". No other kind of key is used: this
// tool offers no paid service.
export const isFreeKey = (key) => typeof key === "string" && /\S:fx$/.test(key);

// The free host. `override` is the address of a local stub in tests.
export function deeplBase(override) {
  return override ? String(override).replace(/\/+$/, "") : FREE;
}

export function translationRoute({ settings, key } = {}) {
  return settings?.external?.deepl?.enabled === true && key ? "deepl" : "oneshot";
}

const WORDS = { 403: "key rejected", 429: "too many requests", 456: "quota exceeded" };
const TARGET = /^[A-Za-z]{2,3}(?:-[A-Za-z]{2,4})?$/;

export async function deeplTranslate({
  text,
  target,
  key,
  enabled = false,
  fetchFn,
  base,
  log = () => {},
} = {}) {
  const fail = (error) => ({ ok: false, error, route: "deepl" });
  if (enabled !== true) return fail("DeepL is off. Turn it on in Setup; the text leaves this PC.");
  if (!key) return fail("DeepL has no key. Enter one in Setup.");
  if (!isFreeKey(key)) return fail(DEEPL_FREE_ONLY);
  const body = String(text ?? "");
  if (!body.trim()) return fail("Select some text first.");
  if (!TARGET.test(String(target ?? ""))) return fail("DeepL needs a target language code.");
  const target_lang = String(target).toUpperCase();
  const line = (status) =>
    log(`[deepl] translate chars=${body.length} target=${target_lang} status=${status}`);
  let res;
  let data = null;
  try {
    res = await fetchFn(`${deeplBase(base)}/v2/translate`, {
      method: "POST",
      headers: { Authorization: `DeepL-Auth-Key ${key}`, "Content-Type": "application/json" },
      body: JSON.stringify({ text: [body], target_lang }),
    });
    data = await res.json().catch(() => null);
  } catch (e) {
    line("failed");
    return fail(
      `DeepL could not be reached: ${String(e?.message ?? e)
        .split(key)
        .join("<key>")}`,
    );
  }
  line(res.status);
  if (!res.ok) {
    const said = typeof data?.message === "string" && data.message ? ` (${data.message})` : "";
    const words = WORDS[res.status];
    return fail(words ? `DeepL: ${words}${said}` : `DeepL answered ${res.status}${said}`);
  }
  const out = data?.translations?.[0]?.text;
  if (typeof out !== "string") return fail("DeepL sent no translation.");
  return { ok: true, text: out, route: "deepl" };
}
