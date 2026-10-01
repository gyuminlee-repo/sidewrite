// Grammar and spelling check with LanguageTool (P6). Off unless the caller
// passes enabled: true (the user's consent switch). fetchFn is injected.
// Never throws. The checked text is never logged.
//
// Public API limits (dev.languagetool.org/public-http-api, read 2026-10-01):
// 20 requests a minute per IP, 75 kB a minute, 20 kB a request, POST only,
// no guarantee; the limits can change. Text goes out in chunks under 18 kB
// and every request passes a limiter that keeps both per-minute limits.
// Not checked: whether language=auto is accepted and whether showing the
// link is required; the link is shown anyway.
//
// Not wired yet. The caller keeps one limiter for the whole daemon, passes
// the paragraphs to check, shows the matches with LT_LINK above them and
// turns a picked match into a tracked edit with matchToEdit().

export const LT_LINK = "https://languagetool.org";
export const LT_URL = "https://api.languagetool.org";
export const LT_LIMITS = { requestsPerMinute: 20, bytesPerMinute: 75_000, bytesPerRequest: 20_000 };
export const LT_NOTICE =
  "This feature sends the checked text from this PC to LanguageTool (languagetool.org), a public service with no guarantee.";
export const LT_OFF = "LanguageTool is off. Turn it on in Setup; the checked text leaves this PC.";

const CHUNK_BYTES = 18_000;
const JOIN = "\n\n";
const bytes = (s) => Buffer.byteLength(s);

// paragraphs: [{ index, id, text }]. Chunks hold whole paragraphs joined by a
// blank line; parts say where each paragraph sits in the chunk (UTF-16
// offsets, the unit LanguageTool answers in). A paragraph that is too large
// on its own is refused, not cut.
export function chunkParagraphs(paragraphs, { maxBytes = CHUNK_BYTES } = {}) {
  const chunks = [];
  const refused = [];
  let cur = null;
  let size = 0;
  for (const paragraph of paragraphs) {
    const text = String(paragraph.text ?? "");
    if (!text.trim()) continue;
    const n = bytes(text);
    if (n > maxBytes) {
      refused.push({ index: paragraph.index, bytes: n });
      continue;
    }
    if (cur && size + JOIN.length + n > maxBytes) cur = null;
    if (!cur) {
      cur = { text: "", parts: [] };
      chunks.push(cur);
      size = 0;
    } else {
      cur.text += JOIN;
      size += JOIN.length;
    }
    cur.parts.push({ paragraph, start: cur.text.length, length: text.length });
    cur.text += text;
    size += n;
  }
  return { chunks, refused };
}

// A sliding one-minute window over requests and bytes. take(n) resolves when
// a request of n bytes may be sent, and counts it.
export function createLimiter({
  now = Date.now,
  sleep = (ms) => new Promise((r) => setTimeout(r, ms)),
  requests = LT_LIMITS.requestsPerMinute,
  maxBytes = LT_LIMITS.bytesPerMinute,
  windowMs = 60_000,
} = {}) {
  let sent = [];
  let queue = Promise.resolve();
  const one = async (n) => {
    for (;;) {
      const t = now();
      sent = sent.filter((e) => t - e.t < windowMs);
      const used = sent.reduce((sum, e) => sum + e.n, 0);
      if (!sent.length || (sent.length < requests && used + n <= maxBytes)) {
        sent.push({ t, n });
        return;
      }
      await sleep(sent[0].t + windowMs - t);
    }
  };
  return {
    take(n) {
      queue = queue.then(() => one(n));
      return queue;
    },
  };
}

export async function ltCheck({
  paragraphs,
  enabled = false,
  fetchFn,
  limiter,
  url = LT_URL,
  language = "auto",
  log = () => {},
} = {}) {
  if (enabled !== true) return { ok: false, error: LT_OFF };
  const { chunks, refused } = chunkParagraphs(paragraphs ?? []);
  const endpoint = `${String(url).replace(/\/+$/, "")}/v2/check`;
  const matches = [];
  for (const chunk of chunks) {
    const size = bytes(chunk.text);
    await limiter.take(size);
    const line = (status, found) =>
      log(
        `[languagetool] check paragraphs=${chunk.parts.length} bytes=${size} status=${status}${
          found === undefined ? "" : ` matches=${found}`
        }`,
      );
    let res;
    let data;
    try {
      res = await fetchFn(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({ text: chunk.text, language }).toString(),
      });
      data = res.ok ? await res.json() : null;
    } catch (e) {
      line("failed");
      return { ok: false, error: `LanguageTool could not be reached: ${String(e?.message ?? e)}` };
    }
    if (res.status === 429) {
      line(429);
      return {
        ok: false,
        error: "LanguageTool: too many requests. Wait a minute and try again.",
      };
    }
    if (!res.ok) {
      line(res.status);
      return { ok: false, error: `LanguageTool answered ${res.status}` };
    }
    if (!Array.isArray(data?.matches)) {
      line(res.status);
      return { ok: false, error: "LanguageTool sent an unexpected answer." };
    }
    let found = 0;
    for (const m of data.matches) {
      const part = chunk.parts.find(
        (p) => m.offset >= p.start && m.offset + m.length <= p.start + p.length,
      );
      if (!part || !Number.isInteger(m.offset) || !(m.length > 0)) continue;
      found++;
      matches.push({
        paragraph: { index: part.paragraph.index, id: part.paragraph.id },
        offset: m.offset - part.start,
        length: m.length,
        message: String(m.message ?? ""),
        replacements: (m.replacements ?? [])
          .map((r) => r?.value)
          .filter((v) => typeof v === "string")
          .slice(0, 5),
        rule_id: String(m.rule?.id ?? ""),
      });
    }
    line(res.status, found);
  }
  return { ok: true, matches, refused, link: LT_LINK };
}

// The find/replace of one match, for a tracked edit. `match.paragraph` must
// carry the paragraph text the check was run on. When the matched words occur
// more than once, the find grows by whole words (to the right, then to the
// left) until it occurs once. null when there is nothing to apply.
export function matchToEdit(match, pick = 0) {
  const text = String(match?.paragraph?.text ?? "");
  const replacement = match?.replacements?.[pick];
  let a = match?.offset;
  let b = a + match?.length;
  if (typeof replacement !== "string" || !Number.isInteger(a) || !(b > a) || b > text.length) {
    return null;
  }
  const start = a;
  const end = b;
  const once = () => {
    const piece = text.slice(a, b);
    return text.indexOf(piece) === a && text.indexOf(piece, a + 1) === -1;
  };
  while (!once() && (a > 0 || b < text.length)) {
    if (b < text.length) b += /^\s*\S*/.exec(text.slice(b))[0].length || 1;
    else a -= /\S*\s*$/.exec(text.slice(0, a))[0].length || 1;
  }
  return {
    find: text.slice(a, b),
    replace: text.slice(a, start) + replacement + text.slice(end, b),
    paragraph_hint: match.paragraph.id,
    paragraph_index: match.paragraph.index,
  };
}
