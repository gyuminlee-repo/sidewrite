// Check the reference list against Crossref (P7). Report only: the document
// is never changed. Off unless the caller passes enabled: true (the user's
// consent switch) and a contact e-mail for Crossref's polite pool. fetchFn,
// now and sleep are injected. Never throws. Neither the e-mail nor any
// entry text is logged.
//
// Checked against the Crossref docs (read 2026-10-01): the polite pool takes
// mailto=<address> on every request. Not found there: a rate limit in
// numbers. So requests are sequential, at most five a second, and a 429 is
// answered by waiting as long as Retry-After says.
//
// Verdicts: match, not_found (no such DOI, or no DOI and no close record),
// title_mismatch, author_mismatch (first author), year_mismatch, candidate
// (the entry has no DOI and Crossref's first search hit has its title),
// error.
//
// Not wired yet. The caller reads the switch and the e-mail from the
// settings and passes referenceEntries(snapshot.paragraphs, profile).

import { referenceSection, paraText } from "./ref-section.mjs";

export const CROSSREF_URL = "https://api.crossref.org";
export const CROSSREF_NOTICE =
  "This feature sends each reference entry and your contact e-mail from this PC to Crossref (crossref.org). Nothing else of the manuscript is sent.";
export const CROSSREF_OFF =
  "The Crossref check is off. Turn it on in Setup; the reference entries leave this PC.";
export const CROSSREF_NEEDS_MAILTO = "Enter a contact e-mail for the Crossref polite pool.";

const GAP_MS = 200; // five requests a second at most
const RETRIES = 2;
const TITLE_MIN = 0.9;

export function referenceEntries(paragraphs, profile = {}) {
  return referenceSection(paragraphs, profile).entries;
}

const DOI = /\b10\.\d{4,9}\/[^\s"<>]+/;

export function extractDoi(text) {
  let doi = DOI.exec(String(text ?? ""))?.[0];
  if (!doi) return null;
  for (;;) {
    const cut = doi.replace(/[.,;:'"\]]+$/, "");
    const open = cut.split("(").length;
    const close = cut.split(")").length;
    const next = cut.endsWith(")") && close > open ? cut.slice(0, -1) : cut;
    if (next === doi) return doi;
    doi = next;
  }
}

const tokens = (s) =>
  String(s ?? "")
    .replace(/<[^>]*>/g, " ")
    .normalize("NFKD")
    .replace(/\p{M}/gu, "")
    .toLowerCase()
    .split(/[^\p{L}\p{N}]+/u)
    .filter(Boolean);

// The share of the title's words that stand in the entry (0..1).
export function titleShare(title, entry) {
  const want = tokens(title);
  if (!want.length) return 0;
  const have = new Map();
  for (const t of tokens(entry)) have.set(t, (have.get(t) ?? 0) + 1);
  let hit = 0;
  for (const t of want) {
    const n = have.get(t) ?? 0;
    if (n > 0) {
      hit++;
      have.set(t, n - 1);
    }
  }
  return hit / want.length;
}

const yearsOf = (work) =>
  ["issued", "published", "published-print", "published-online"]
    .map((k) => work?.[k]?.["date-parts"]?.[0]?.[0])
    .filter((y) => Number.isInteger(y));

// The verdict of one entry against one Crossref record.
function compare(entry, work) {
  const title = String(work?.title?.[0] ?? "").replace(/<[^>]*>/g, "");
  if (titleShare(title, entry) < TITLE_MIN) {
    return { verdict: "title_mismatch", detail: `Crossref title: ${title}` };
  }
  const family = work?.author?.[0]?.family;
  if (family) {
    const have = new Set(tokens(entry));
    if (!tokens(family).every((t) => have.has(t))) {
      return { verdict: "author_mismatch", detail: `Crossref first author: ${family}` };
    }
  }
  const years = yearsOf(work);
  const written = entry.replace(DOI, " ").match(/(?<!\d)(?:19|20)\d\d(?!\d)/g) ?? [];
  if (years.length && written.length && !written.some((y) => years.includes(Number(y)))) {
    return { verdict: "year_mismatch", detail: `Crossref year: ${years[0]}` };
  }
  return { verdict: "match", detail: title };
}

export async function checkReferences({
  entries,
  mailto,
  enabled = false,
  fetchFn,
  now = Date.now,
  sleep = (ms) => new Promise((r) => setTimeout(r, ms)),
  base = CROSSREF_URL,
  log = () => {},
} = {}) {
  if (enabled !== true) return { ok: false, error: CROSSREF_OFF };
  const address = String(mailto ?? "").trim();
  if (!/^[^\s@]+@[^\s@]+$/.test(address)) return { ok: false, error: CROSSREF_NEEDS_MAILTO };
  const root = String(base).replace(/\/+$/, "");
  const polite = `mailto=${encodeURIComponent(address)}`;
  let last = null;

  // One GET; { status, body } or { error }. `busy` when Crossref keeps saying 429.
  const get = async (url, kind) => {
    for (let attempt = 0; ; attempt++) {
      if (last !== null && now() - last < GAP_MS) await sleep(GAP_MS - (now() - last));
      last = now();
      let res;
      try {
        res = await fetchFn(url);
      } catch (e) {
        log(`[crossref] ${kind} status=failed`);
        return {
          error: String(e?.message ?? e)
            .split(address)
            .join("<e-mail>"),
        };
      }
      log(`[crossref] ${kind} status=${res.status}`);
      if (res.status === 429) {
        if (attempt >= RETRIES) return { busy: true };
        const after = Number(res.headers?.get?.("retry-after"));
        await sleep(Math.min(Number.isFinite(after) && after > 0 ? after : 5, 120) * 1000);
        continue;
      }
      if (!res.ok) return { status: res.status };
      return { status: res.status, body: await res.json().catch(() => null) };
    }
  };

  const rows = [];
  for (const p of entries ?? []) {
    const entry = paraText(p);
    const doi = extractDoi(entry);
    const row = (verdict, detail, found = doi) => {
      rows.push({ entry_index: p.index, verdict, detail, doi: found });
    };
    const r = doi
      ? await get(`${root}/works/${encodeURIComponent(doi)}?${polite}`, "work")
      : await get(
          `${root}/works?query.bibliographic=${encodeURIComponent(entry.slice(0, 500))}&rows=1&${polite}`,
          "search",
        );
    if (r.busy) {
      return { ok: false, error: "Crossref: too many requests. Try again later.", rows };
    }
    if (r.error) row("error", `Crossref could not be reached: ${r.error}`);
    else if (doi && r.status === 404) row("not_found", "Crossref has no record of this DOI");
    else if (r.status !== 200 || !r.body?.message) row("error", `Crossref answered ${r.status}`);
    else if (doi) {
      const c = compare(entry, r.body.message);
      row(c.verdict, c.detail);
    } else {
      const hit = r.body.message.items?.[0];
      const title = String(hit?.title?.[0] ?? "");
      if (hit?.DOI && titleShare(title, entry) >= TITLE_MIN) {
        row("candidate", `no DOI in the entry; Crossref has ${hit.DOI}: ${title}`, hit.DOI);
      } else {
        row("not_found", "no DOI in the entry and no close Crossref record");
      }
    }
  }
  return { ok: true, rows };
}
