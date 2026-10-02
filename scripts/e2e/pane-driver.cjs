// Drives the Sidewrite Word task pane over the Chrome DevTools Protocol.
// Runs on Windows node (the pane is a WebView2 inside Word). One step per
// process: the runner calls it once per step and reads one JSON line.
//
//   node pane-driver.cjs <step> <args.json>
//
// Environment: PLAYWRIGHT_CORE (path of a playwright-core package),
// E2E_CDP_PORT (the --remote-debugging-port given to Word),
// E2E_PANE_ORIGIN (default http://127.0.0.1:47834).
// E2E_SHOT (optional, the run's --screenshots): the path of a PNG. The pane
// is photographed when the step is over (also when it failed), and while it
// runs after 4 s and after 20 s (<name>-wait1.png, <name>-wait2.png) with the
// texts that tell the user what is going on (`waiting` in the answer). Each
// picture is two files: <name>.png, what the user sees (the window of the
// pane as it is scrolled), and <name>-full.png, the whole tab (the pane
// scrolls inside an element, so a full-page picture alone shows no more than
// the window).
// Element ids come from taskpane/word/index.html and the class names from
// taskpane/shared/review-pane.js, response-pane.js and tools-pane.js.

const fs = require("node:fs");

const PW = process.env.PLAYWRIGHT_CORE;
const PORT = process.env.E2E_CDP_PORT || "9333";
const ORIGIN = process.env.E2E_PANE_ORIGIN || "http://127.0.0.1:47834";
const TURN_TIMEOUT = Number(process.env.E2E_TURN_TIMEOUT_MS || 20 * 60_000);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
// The daemon keeps a model list that failed for 60 s (daemon/model-list.mjs
// fallbackTtlMs); asking again sooner gets the same fallback.
const MODELS_RETRY_AFTER_MS = 61_000;

async function connect(timeoutMs = 120_000) {
  const { chromium } = require(PW);
  const end = Date.now() + timeoutMs;
  let last;
  while (Date.now() < end) {
    try {
      const browser = await chromium.connectOverCDP(`http://127.0.0.1:${PORT}`);
      const page = browser
        .contexts()
        .flatMap((c) => c.pages())
        .find((p) => p.url().startsWith(ORIGIN + "/"));
      if (page) return { browser, page };
      await browser.close();
      last = new Error("no pane page yet");
    } catch (e) {
      last = e;
    }
    await sleep(1000);
  }
  throw new Error(`pane not reachable on port ${PORT}: ${last?.message}`);
}

const text = (page, sel) => page.$eval(sel, (n) => n.textContent.trim()).catch(() => "");

// Clear the message line, run the action, then wait for a new message (or a
// predicate), so a message from an earlier step is never read as this one.
async function clearMessage(page) {
  await page.evaluate(() => {
    const m = document.getElementById("review-message");
    if (!m) return;
    m.textContent = "";
    m.hidden = true;
    const d = document.getElementById("review-message-details");
    if (d) d.textContent = "";
  });
}
// until (a RegExp source): the message this step's own action ends with. The
// line also shows messages that are not that: "Waiting for the daemon to
// confirm the stored provider choice..." while a click is held right after a
// (re)connect, and the result of an earlier action that ended late. With
// `until`, only a message that matches it or an error (form-error) ends the
// wait.
async function waitMessage(page, timeout = 120_000, until = null) {
  await page.waitForFunction(
    (pattern) => {
      const m = document.getElementById("review-message");
      const t = m?.textContent.trim() ?? "";
      if (t === "") return false;
      return !pattern || m.classList.contains("form-error") || new RegExp(pattern).test(t);
    },
    until,
    { timeout, polling: 200 },
  );
  return text(page, "#review-message");
}

// What the Review tab shows: message, cards with badges and flags, decisions,
// outside changes, ledger, violations.
async function view(page) {
  return page.evaluate(() => {
    const t = (n) => (n ? n.textContent.replace(/\s+/g, " ").trim() : "");
    const q = (s, root = document) => [...root.querySelectorAll(s)];
    return {
      message: t(document.getElementById("review-message")),
      // The numbers behind the sentence of a refresh (under "Details").
      message_details: t(document.getElementById("review-message-details")),
      message_error:
        document.getElementById("review-message")?.classList.contains("form-error") ?? false,
      file: t(document.getElementById("review-file")),
      comments: q(".review-comment").map((c) => ({
        ref: t(c.querySelector(".review-comment-head strong")),
        // What the reviewer asks, one row per request: the id and the text of each,
        // as the old line "Requirements: q1 text; q2 text" said them.
        reqs_line: (() => {
          const rows = q(".review-req", c);
          return rows.length
            ? `Requirements: ${rows
                .map((r) => `${r.getAttribute("data-req-id")} ${r.getAttribute("data-req-text")}`)
                .join("; ")}`
            : "";
        })(),
        changes: q(".review-change-wrap", c).map((w) => ({
          badge: t(w.querySelector(".review-badge")),
          text: t(w.querySelector(".review-change")),
          flags: q(".review-flag", w).map(t),
        })),
        blocked: q(".review-blocked", c).map(t),
      })),
      decisions: q("#review-decisions > div").map(t),
      // What each waiting item is about, from its attributes (its sentence is plain words).
      decisions_raw: q("#review-decisions > div").map((d) => ({
        kind: d.getAttribute("data-kind"),
        ref: d.getAttribute("data-ref"),
        req: d.getAttribute("data-req"),
      })),
      consistency: q("#review-consistency .review-violation").map(t),
      external: q("#review-external .review-external-item .review-violation").map(t),
      // The rule id and kind of a row are in its attributes; the row shows the kind in words.
      ledger: q("#review-ledger .review-ledger-item").map(
        (n) => `${n.getAttribute("data-id")} ${n.getAttribute("data-kind")} ${t(n)}`,
      ),
      violations: q("#review-violations .review-violation").map(
        (n) => `${n.getAttribute("data-rule")}: ${t(n)}`,
      ),
      agent_status: t(document.getElementById("agent-status")),
      delete_all_button: t(document.getElementById("review-delete-all")),
    };
  });
}

// ---- Response tab (ids from taskpane/word/index.html, classes from
// taskpane/shared/response-pane.js) ----
async function clearResponse(page) {
  await page.evaluate(() => {
    const m = document.getElementById("response-message");
    if (!m) return;
    m.textContent = "";
    m.hidden = true;
  });
}
// As waitMessage, on the message line of the Response tab. `until` is not
// optional here: every step names the message its own action ends with, so
// "Drafting R1.1..." or a held click is never read as the result. An error
// (form-error) ends the wait as well; the step then reports it.
async function waitResponse(page, timeout, until) {
  await page.waitForFunction(
    (pattern) => {
      const m = document.getElementById("response-message");
      const t = m?.textContent.trim() ?? "";
      if (t === "") return false;
      return m.classList.contains("form-error") || new RegExp(pattern).test(t);
    },
    until,
    { timeout, polling: 200 },
  );
  return text(page, "#response-message");
}
// What the Response tab shows. The comment and the answer of a card are read
// as they are (no whitespace folded): the run compares them with revision.json.
async function responseView(page) {
  return page.evaluate(() => {
    const t = (n) => (n ? n.textContent.replace(/\s+/g, " ").trim() : "");
    const q = (s, root = document) => [...root.querySelectorAll(s)];
    const message = document.getElementById("response-message");
    return {
      message: t(message),
      message_error: (message?.classList.contains("form-error") ?? false) && t(message) !== "",
      next: t(document.getElementById("response-next")),
      profile: document.getElementById("response-profile")?.value ?? "",
      summary: t(document.getElementById("response-summary")),
      checks: q("#response-checks .response-check-row").map(t),
      export_label: t(document.getElementById("response-export")),
      output: t(document.getElementById("response-output")),
      proposal_rows: q("#response-proposal .response-proposal-row").map(t),
      cards: q("#response-comments .response-card").map((c) => ({
        ref: c.getAttribute("data-ref"),
        comment: c.querySelector(".response-comment")?.textContent ?? "",
        text: c.querySelector("textarea.response-text")?.value ?? "",
        declined: t(c.querySelector(".response-declined")),
        stale: t(c.querySelector(".response-stale")),
        undecided: t(c.querySelector(".response-undecided")),
        citations: q(".response-citation", c).map((x) => ({
          line: t(x.querySelector(".response-citation-line")),
          quote: t(x.querySelector(".response-citation-quote")),
        })),
        open_reqs: q(".response-req:not(.response-req-declined) > span", c).map(t),
      })),
    };
  });
}
const responseCard = (page, ref) => page.locator(`.response-card[data-ref="${ref}"]`);
const escapeRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
// Profile, languages and the letter reader are folded away in the pane.
const openProfileDetails = (page) =>
  page.evaluate(() => {
    const d = document.getElementById("response-profile-details");
    if (d) d.open = true;
  });

function card(page, ref) {
  return page.locator(".review-comment", {
    has: page.locator(".review-comment-head strong", {
      hasText: new RegExp(`^${ref.replace(".", "\\.")}$`),
    }),
  });
}

// The Review tab folds things away; a user opens them before using what is in
// them, so the steps do the same. The sections are <details>: pressing the
// summary opens one.
const openFold = (page, id) =>
  page.evaluate((i) => {
    const d = document.getElementById(i);
    if (d && !d.open) d.querySelector("summary")?.click();
  }, id);
// A comment is a row; its card (buttons, requirements, changes) opens when the
// row is pressed. One card is open at a time.
const openCard = (page, ref) =>
  page.evaluate((wanted) => {
    const c = [...document.querySelectorAll(".review-comment")].find(
      (x) => x.querySelector(".review-comment-head strong")?.textContent.trim() === wanted,
    );
    const body = c?.querySelector(".review-comment-body");
    if (body?.hidden) c.querySelector(".review-comment-head")?.click();
  }, ref);
// The paste box shrinks to one line once the review is split; the author
// shows it again before pasting another review.
const showPasteBox = (page) =>
  page.evaluate(() => {
    if (document.getElementById("review-paste-box")?.hidden) {
      document.getElementById("review-paste-toggle")?.click();
    }
  });

const turnBusy = (page) =>
  page.evaluate(() => {
    const p = document.getElementById("review-progress");
    return !!document.getElementById("input")?.disabled || p?.hidden === false;
  });

// Providers section of the Setup tab (ids from taskpane/word/index.html,
// rows from taskpane/shared/providers-pane.js). The list is asked for on
// connect; wait until it is drawn or its message says why it is not.
// ---- Tools tab (ids from taskpane/word/index.html, classes from
// taskpane/shared/tools-pane.js) ----
async function clearTools(page, box = null) {
  await page.evaluate((id) => {
    const m = document.getElementById("tools-message");
    if (m) {
      m.textContent = "";
      m.hidden = true;
    }
    if (id) document.getElementById(id)?.replaceChildren();
  }, box);
}
// The end of a Tools tab action: its message matches `until`, or the box
// `box` has lines, or the message is an error (form-error).
async function waitTools(page, timeout, { until = null, box = null }) {
  await page.waitForFunction(
    ([wanted, id]) => {
      const m = document.getElementById("tools-message");
      const t = m?.textContent.trim() ?? "";
      if (t !== "" && m.classList.contains("form-error")) return true;
      if (wanted && t !== "" && new RegExp(wanted).test(t)) return true;
      return !!id && (document.getElementById(id)?.children.length ?? 0) > 0;
    },
    [until, box],
    { timeout, polling: 200 },
  );
}
// What the Tools tab shows.
async function toolsView(page) {
  return page.evaluate(() => {
    const t = (n) => (n ? n.textContent.replace(/\s+/g, " ").trim() : "");
    const kids = (id) => [...(document.getElementById(id)?.children ?? [])].map(t);
    const rules = [...(document.querySelectorAll("#tools-rules label") ?? [])].map((l) => {
      const cite = t(l.querySelector(".review-preview"));
      return {
        kind: t(l.querySelector(".review-reqs-line")),
        cite,
        sentence: /^Guideline: "(.*)"$/.exec(cite)?.[1] ?? null,
        checked: l.querySelector("input")?.checked ?? null,
      };
    });
    return {
      doc: t(document.getElementById("tools-doc")),
      message: t(document.getElementById("tools-message")),
      message_error:
        document.getElementById("tools-message")?.classList.contains("form-error") ?? false,
      selection: kids("tools-selection"),
      rules,
      rule_warnings: [...(document.querySelectorAll("#tools-rules > div") ?? [])].map(t),
      checks: kids("tools-checks"),
      marked: kids("tools-marked"),
      anonymized: kids("tools-anonymized"),
    };
  });
}
// Select one paragraph in Word through the pane's own module, as a click on
// a finding does.
async function selectParagraph(page, index) {
  await page.evaluate(
    async ([origin, i]) => {
      const m = await import(`${origin}/shared/tools-word.js`);
      await m.selectParagraphByIndex(i);
    },
    [ORIGIN, index],
  );
}

async function openProviders(page) {
  await page.click('button.tab[data-tab="setup"]');
  await page.waitForFunction(
    () =>
      document.querySelectorAll("#providers-list .provider-row").length > 0 ||
      (document.getElementById("providers-message")?.textContent.trim() ?? "") !== "",
    null,
    { timeout: 120_000, polling: 200 },
  );
}
function providersView(page) {
  return page.evaluate(() => {
    const t = (n) => (n ? n.textContent.replace(/\s+/g, " ").trim() : "");
    const q = (s, root = document) => [...root.querySelectorAll(s)];
    const menu = (id) =>
      q(`#${id} option`).map((o) => ({
        id: o.value,
        text: t(o),
        enabled: !o.disabled,
        note: o.title,
      }));
    const raw = document.getElementById("providers-root")?.getAttribute("data-selection");
    return {
      menus: {
        turn: menu("provider-turn"),
        writing: menu("provider-writing"),
        oneshot: menu("provider-oneshot"),
      },
      models: {
        turn: menu("provider-turn-model"),
        writing: menu("provider-writing-model"),
        oneshot: menu("provider-oneshot-model"),
      },
      selection: raw ? JSON.parse(raw) : {},
      rows: q("#providers-list .provider-row").map((r) => ({
        id: r.getAttribute("data-provider"),
        state: r.getAttribute("data-state"),
        status: t(r.querySelector(".provider-status")),
        policy: t(r.querySelector(".provider-policy")),
        key_note: t(r.querySelector(".provider-key-note")),
        key_box: !!r.querySelector('input[type="password"]'),
      })),
      message: t(document.getElementById("providers-message")),
    };
  });
}

// What tells a user what the pane is doing: the lines that say it.
function paneState(page) {
  return page.evaluate(() => {
    const t = (id) => (document.getElementById(id)?.textContent ?? "").replace(/\s+/g, " ").trim();
    return {
      message: t("review-message"),
      progress: t("review-progress"),
      agent_status: t("agent-status"),
      connection: t("connection-status"),
      models_line: t("models-line"),
      input_disabled: document.getElementById("input")?.disabled ?? null,
    };
  });
}

// ---- pictures (E2E_SHOT) ----------------------------------------------------
// Is the progress line of the Review tab where the user sees it without
// scrolling: inside the window and inside every element it scrolls in?
function progressInView(page) {
  return page.evaluate(() => {
    // The work of a comment shows inside its card, the work of a run in the
    // bar at the top of the tab: either one in view is what the user needs.
    const shown = (e) => !!e && !e.hidden && !e.closest("[hidden]");
    const card = [...document.querySelectorAll(".review-card-progress")].find(shown);
    const el = card ?? document.getElementById("review-progress");
    const r = el?.getBoundingClientRect?.();
    if (!r || !shown(el) || r.height === 0) return false;
    let top = 0;
    let bottom = window.innerHeight ?? 0;
    for (let p = el.parentElement; p; p = p.parentElement) {
      const b = p.getBoundingClientRect();
      top = Math.max(top, b.top);
      bottom = Math.min(bottom, b.bottom);
    }
    return r.top >= top && r.bottom <= bottom;
  });
}
function viewportOf(page) {
  return page.evaluate(() => ({
    width: window.innerWidth ?? null,
    height: window.innerHeight ?? null,
    device_pixel_ratio: window.devicePixelRatio ?? null,
  }));
}
// Where a user who follows the step looks when it is over: the card of the
// comment the step was for, else the message line of the tab that is shown.
// Scrolled into view before the picture; "card", "message" or null.
function lookAt(page, ref) {
  return page.evaluate((wanted) => {
    const card = wanted
      ? [...document.querySelectorAll(".review-comment")].find(
          (c) => c.querySelector(".review-comment-head strong")?.textContent.trim() === wanted,
        )
      : null;
    const line = document.querySelector(
      '.tab-content:not([hidden]) [id$="-message"]:not([hidden])',
    );
    const target = card ?? (line?.textContent.trim() ? line : null);
    target?.scrollIntoView?.({ block: card ? "start" : "nearest" });
    return card ? "card" : target ? "message" : null;
  }, ref ?? null);
}
// The whole tab in one picture. The elements the pane scrolls in, and what
// holds them, are let grow to their content for the moment of the picture;
// then every style and every scroll position is put back as it was.
async function fullShot(page, file) {
  await page.evaluate(() => {
    if (!document.documentElement) return;
    const scrollers = [...document.querySelectorAll("body *")].filter((e) => {
      if (["TEXTAREA", "INPUT", "SELECT"].includes(e.tagName)) return false;
      if (e.clientHeight === 0 || e.scrollHeight <= e.clientHeight + 1) return false;
      const o = window.getComputedStyle?.(e).overflowY;
      return o === "auto" || o === "scroll";
    });
    const grown = new Set();
    for (const s of scrollers) for (let p = s; p; p = p.parentElement) grown.add(p);
    window.__shotUndo = [...grown].map((e) => ({
      e,
      style: e.getAttribute("style"),
      top: e.scrollTop,
    }));
    for (const e of grown) {
      for (const [k, v] of [
        ["height", "auto"],
        ["max-height", "none"],
        ["overflow", "visible"],
        ["flex-shrink", "0"],
        ["flex-basis", "auto"],
      ])
        e.style.setProperty(k, v, "important");
    }
    // No scroll bar of the page itself: it would take width from the tab.
    if (grown.size)
      document.documentElement.style.setProperty("scrollbar-width", "none", "important");
  });
  try {
    await page.screenshot({ path: file, fullPage: true });
  } finally {
    await page.evaluate(() => {
      const undo = window.__shotUndo ?? [];
      for (const u of undo) {
        if (u.style === null) u.e.removeAttribute("style");
        else u.e.setAttribute("style", u.style);
      }
      // Scroll positions last: they hold only once the heights are back.
      for (const u of undo) u.e.scrollTop = u.top;
      delete window.__shotUndo;
    });
  }
}
// One picture, both files. Returns { shot, shot_full } (the paths).
async function takeShot(page, file) {
  await page.screenshot({ path: file });
  const full = file.replace(/\.png$/, "-full.png");
  await fullShot(page, full);
  return { shot: file, shot_full: full };
}

const steps = {
  // Screenshot runs: one tab shown, nothing pressed in it. a.tab: chat,
  // review, tools, setup, library or the tab of the letter.
  async "show-tab"(page, a) {
    await page.waitForFunction(
      (tab) => {
        const shown = () =>
          document.querySelector(`.tab-content[data-tab="${tab}"]`)?.hidden === false;
        if (!shown()) document.querySelector(`button.tab[data-tab="${tab}"]`)?.click();
        return shown();
      },
      a.tab,
      { timeout: 60_000, polling: 500 },
    );
    // What the tab asks the daemon for when it opens has time to arrive.
    await sleep(a.settle_ms ?? 3000);
    return { tab: a.tab, ...(await paneState(page)) };
  },

  // Screenshot runs: "Refresh status" pressed while no daemon answers.
  async "refresh-offline"(page) {
    await clearMessage(page);
    await page.click("#review-refresh", { timeout: 15_000 });
    const message = await waitMessage(page, 40_000).catch(
      (e) => `no message within 40 s (${String(e.message).split("\n")[0]})`,
    );
    return { message, ...(await paneState(page)) };
  },

  // The tab buttons get their click handlers when taskpane.js has run, which
  // can be after the page is reachable over CDP: a click before that does
  // nothing and every later step would wait on a hidden tab. Click until the
  // Review tab is the one shown.
  async "open-review"(page) {
    await page.waitForFunction(
      () => {
        const shown = () =>
          document.querySelector('.tab-content[data-tab="review"]')?.hidden === false;
        if (!shown()) document.querySelector('button.tab[data-tab="review"]')?.click();
        return shown();
      },
      null,
      { timeout: 60_000, polling: 500 },
    );
    await page.waitForFunction(
      () => document.getElementById("review-file")?.textContent.startsWith("Review file:"),
      null,
      { timeout: 60_000 },
    );
    return view(page);
  },

  // a.review_file fills the paste box first; without it the text already
  // there is split (stage 2.5 loads it with load-review-file).
  async split(page, a) {
    await showPasteBox(page);
    if (a.review_file) await page.fill("#review-input", fs.readFileSync(a.review_file, "utf8"));
    await clearMessage(page);
    await page.click("#review-split");
    const message = await waitMessage(page, 10 * 60_000, "^Split into ");
    const refs = await page.$$eval(".review-draft-row .review-ref", (ns) => ns.map((n) => n.value));
    return { message, draft_refs: refs };
  },

  // Remove draft rows whose ref is not R<n>.<m>, put requirement lines on
  // the rows named in a.requirements, Save, wait for the comment cards.
  async "edit-draft"(page, a) {
    const keep = new RegExp(a.keep_ref || "^R\\d+\\.\\d+$");
    const removed = [];
    for (;;) {
      const refs = await page.$$eval(".review-draft-row .review-ref", (ns) =>
        ns.map((n) => n.value),
      );
      const i = refs.findIndex((r) => !keep.test(r));
      if (i < 0) break;
      removed.push(refs[i]);
      await page
        .locator(".review-draft-row")
        .nth(i)
        .locator("button", { hasText: "Remove" })
        .click();
    }
    for (const [ref, lines] of Object.entries(a.requirements || {})) {
      const refs = await page.$$eval(".review-draft-row .review-ref", (ns) =>
        ns.map((n) => n.value),
      );
      const i = refs.indexOf(ref);
      if (i < 0) throw new Error(`no draft row ${ref}`);
      await page
        .locator(".review-draft-row")
        .nth(i)
        .locator("textarea.review-reqs")
        .fill(lines.join("\n"));
    }
    const kept = await page.$$eval(".review-draft-row .review-ref", (ns) => ns.map((n) => n.value));
    await page.click("#review-save");
    await page.waitForFunction(
      (n) =>
        document.querySelectorAll(".review-comment").length === n &&
        document.querySelectorAll(".review-draft-row").length === 0,
      kept.length,
      { timeout: 60_000 },
    );
    return { removed, kept, ...(await view(page)) };
  },

  async "ledger-add"(page, a) {
    const items = JSON.parse(fs.readFileSync(a.ledger_file, "utf8")).items;
    await openFold(page, "review-ledger-box");
    await openFold(page, "review-ledger-advanced");
    for (const it of items) {
      const before = await page.$$eval("#review-ledger .review-ledger-item", (ns) => ns.length);
      await page.selectOption("#ledger-kind", it.kind);
      await page.selectOption("#ledger-scope", it.scope);
      await page.fill("#ledger-text", it.text);
      await page.selectOption("#ledger-check-type", it.check.type);
      await page.fill("#ledger-pattern", it.check.pattern ?? "");
      await page.fill("#ledger-flags", it.check.flags ?? "");
      await page.selectOption("#ledger-expect", it.check.expect);
      await clearMessage(page);
      await page.click("#review-ledger-add");
      await page.waitForFunction(
        (n) =>
          document.querySelectorAll("#review-ledger .review-ledger-item").length === n + 1 ||
          (document.getElementById("review-message")?.textContent.trim() ?? "") !== "",
        before,
        { timeout: 60_000 },
      );
    }
    return view(page);
  },

  async "check-all"(page) {
    await openFold(page, "review-ledger-box");
    await clearMessage(page);
    await page.click("#review-check-all");
    // Done when no ledger item waits for its first check, or an error shows.
    await page.waitForFunction(
      () =>
        document.getElementById("review-decisions")?.textContent.includes("not checked yet") ===
          false || (document.getElementById("review-message")?.textContent.trim() ?? "") !== "",
      null,
      { timeout: 10 * 60_000, polling: 500 },
    );
    await sleep(500);
    return view(page);
  },

  async preview(page, a) {
    await openCard(page, a.ref);
    // The preview button is under the Advanced of the card.
    await page.evaluate((wanted) => {
      const c = [...document.querySelectorAll(".review-comment")].find(
        (x) => x.querySelector(".review-comment-head strong")?.textContent.trim() === wanted,
      );
      const d = c?.querySelector("details.review-card-advanced");
      if (d && !d.open) d.querySelector("summary")?.click();
    }, a.ref);
    await card(page, a.ref).locator("button", { hasText: "Preview prompt" }).click();
    await page.waitForFunction(
      () => document.getElementById("review-preview")?.hidden === false,
      null,
      {
        timeout: 60_000,
      },
    );
    return { prompt: await text(page, "#review-preview") };
  },

  // Propose changes on one comment and wait until the turn is over.
  async propose(page, a) {
    if (await turnBusy(page)) throw new Error("a turn is already running");
    await openCard(page, a.ref);
    const btn = card(page, a.ref).locator("button", { hasText: "Propose changes" });
    await btn.click();
    await page.waitForFunction(() => document.getElementById("input")?.disabled, null, {
      timeout: 30_000,
    });
    const samples = [];
    const started = Date.now();
    while (await turnBusy(page)) {
      const p = await text(page, "#review-progress");
      if (p && samples[samples.length - 1] !== p && samples.length < 6) samples.push(p);
      if (Date.now() - started > TURN_TIMEOUT)
        throw new Error(`turn ${a.ref} still running after ${TURN_TIMEOUT} ms`);
      await sleep(1000);
    }
    // turn_done reloads and refreshes the pane; let that land.
    await sleep(4000);
    const assistant = await page.$$eval("#messages .msg.assistant", (ns) =>
      ns.length ? ns[ns.length - 1].textContent.trim() : "",
    );
    return {
      progress_samples: samples,
      seconds: Math.round((Date.now() - started) / 1000),
      assistant,
      ...(await view(page)),
    };
  },

  async refresh(page) {
    await clearMessage(page);
    await page.click("#review-refresh");
    await waitMessage(page, 120_000);
    return view(page);
  },

  // "Check revised paragraphs": a status refresh, then the check of the
  // paragraphs that hold two or more changes (one model request or more).
  async "check-paragraphs"(page) {
    await openFold(page, "review-more");
    await clearMessage(page);
    await page.click("#review-check-paragraphs");
    await waitMessage(
      page,
      20 * 60_000,
      "^(Checked \\d+ revised paragraph|No paragraph holds two or more changes)",
    );
    return view(page);
  },

  // "Rewrite this paragraph" under "Check revised paragraphs": the n-th such
  // button (a.nth, 0 by default) starts a review turn that writes the
  // paragraph anew; waits as `propose` does until the turn is over.
  async "rewrite-paragraph"(page, a = {}) {
    if (await turnBusy(page)) throw new Error("a turn is already running");
    await openFold(page, "review-more");
    const btn = page
      .locator("#review-consistency button", { hasText: "Rewrite this paragraph" })
      .nth(a.nth ?? 0);
    await btn.click();
    await page.waitForFunction(() => document.getElementById("input")?.disabled, null, {
      timeout: 30_000,
    });
    const started = Date.now();
    while (await turnBusy(page)) {
      if (Date.now() - started > TURN_TIMEOUT)
        throw new Error(`rewrite turn still running after ${TURN_TIMEOUT} ms`);
      await sleep(1000);
    }
    await sleep(4000);
    return { seconds: Math.round((Date.now() - started) / 1000), ...(await view(page)) };
  },

  // "Allow exception" of a blocked edit, pressed and confirmed: the row of
  // comment a.ref that reads "Blocked by <a.blocked_by>:". The daemon sends
  // the stored edit through the guarded write once more; an edit that is
  // refused again ends with its reason on the message line (an error).
  async "allow-exception"(page, a) {
    const row = () =>
      card(page, a.ref).locator(".review-blocked", { hasText: `Blocked by ${a.blocked_by}:` });
    await clearMessage(page);
    await openCard(page, a.ref);
    await row().locator("button", { hasText: "Allow exception" }).click();
    await row().locator("button", { hasText: "Apply despite" }).click();
    // An error ends the wait (the refusal is one); an edit that went in
    // leaves no message and the step runs out of time.
    await waitMessage(page, 180_000, "forbidden rule");
    return view(page);
  },

  async "clean-tool"(page) {
    await openFold(page, "review-submit-box");
    await clearMessage(page);
    await page.click("#review-clean-tool");
    await waitMessage(page, 120_000);
    return view(page);
  },

  // First press shows the counts on the button; second press deletes.
  async "delete-all"(page) {
    const btn = "#review-delete-all";
    await openFold(page, "review-submit-box");
    await clearMessage(page);
    await page.click(btn);
    await page.waitForFunction(
      () =>
        document
          .getElementById("review-delete-all")
          ?.textContent.startsWith("Back up and delete") ||
        (document.getElementById("review-message")?.textContent.trim() ?? "") !== "",
      null,
      { timeout: 60_000 },
    );
    const offer = await text(page, btn);
    await clearMessage(page);
    await page.click(btn);
    const message = await waitMessage(page, 120_000);
    return { offer, ...(await view(page)), message };
  },

  // ---- stage 2.5: the model menu of the chat box ----

  // The rows of #composer-model once the daemon's list (or its failure) is
  // in: taskpane.js loadModelMenu sets the tooltip either way.
  //
  // The list comes from the account and can be late once; the pane then
  // shows the aliases, which is right for the pane and wrong for this check.
  // An unavailable list is asked for once more: after the daemon's fallback
  // has run out the pane is loaded again (it asks for the list when it
  // connects) and brought back to the Review tab. `retried` and
  // `first_title` say that this happened.
  async models(page) {
    const read = async () => {
      await page.waitForFunction(
        () =>
          /^(Models your Claude account|Model list unavailable)/.test(
            document.getElementById("composer-model")?.title ?? "",
          ),
        null,
        { timeout: 90_000, polling: 200 },
      );
      return page.$eval("#composer-model", (m) => ({
        title: m.title,
        selected: m.value,
        options: [...m.options].map((o) => ({ value: o.value, label: o.textContent.trim() })),
      }));
    };
    const first = await read();
    if (!/^Model list unavailable/.test(first.title)) return first;
    await page.waitForTimeout(MODELS_RETRY_AFTER_MS);
    await page.reload();
    await steps["open-review"](page);
    return { ...(await read()), retried: true, first_title: first.title };
  },

  // Propose changes on one comment, then press Stop while the turn runs
  // (a.after_ms after it began, default 8 s) and wait until the pane is free.
  async "propose-stop"(page, a) {
    if (await turnBusy(page)) throw new Error("a turn is already running");
    await openCard(page, a.ref);
    await card(page, a.ref).locator("button", { hasText: "Propose changes" }).click();
    await page.waitForFunction(() => document.getElementById("input")?.disabled, null, {
      timeout: 30_000,
    });
    const started = Date.now();
    await sleep(a.after_ms ?? 8000);
    if (!(await turnBusy(page))) {
      return { stopped: false, reason: "the turn ended before Stop", ...(await view(page)) };
    }
    // The Stop button of the card of the comment (it presses the Stop of the chat).
    await card(page, a.ref).locator(".review-card-progress button").click();
    while (await turnBusy(page)) {
      if (Date.now() - started > 120_000) throw new Error("the turn did not end after Stop");
      await sleep(500);
    }
    await sleep(3000);
    return {
      stopped: true,
      seconds: Math.round((Date.now() - started) / 1000),
      ...(await view(page)),
    };
  },

  // ---- stage 2.5: Providers section (Setup tab) ----

  // The two menus, the row of every provider and the message line.
  async providers(page) {
    await openProviders(page);
    return providersView(page);
  },

  // Choose a provider (and a model) for a role: a.role "turn" | "oneshot",
  // a.id, a.model. A provider that is not in the menu or is disabled there
  // is not chosen; `refused` then holds the text of its menu row.
  async "select-provider"(page, a) {
    await openProviders(page);
    const menu = `#provider-${a.role}`;
    const option = await page
      .$eval(`${menu} option[value="${a.id}"]`, (o) => ({
        disabled: o.disabled,
        text: o.textContent,
        note: o.title,
      }))
      .catch(() => null);
    const selection = async () => (await providersView(page)).selection[a.role];
    if (!option) return { role: a.role, selected: await selection(), refused: `no ${a.id} row` };
    if (option.disabled) {
      const refused = option.note ? `${option.text}: ${option.note}` : option.text;
      return { role: a.role, selected: await selection(), refused };
    }
    const settled = (want) =>
      page.waitForFunction(
        ([role, w]) => {
          const m = document.getElementById("providers-message")?.textContent.trim() ?? "";
          const raw = document.getElementById("providers-root")?.getAttribute("data-selection");
          const s = raw ? JSON.parse(raw)[role] : null;
          return m !== "" || (s && s.id === w.id && (w.model == null || s.model === w.model));
        },
        [a.role, want],
        { timeout: 90_000, polling: 200 },
      );
    await page.evaluate(() => {
      const m = document.getElementById("providers-message");
      if (m) m.textContent = "";
    });
    await page.selectOption(menu, a.id);
    await settled({ id: a.id });
    if (a.model && !(await text(page, "#providers-message"))) {
      await page.selectOption(`${menu}-model`, a.model);
      await settled({ id: a.id, model: a.model });
    }
    const v = await providersView(page);
    return { role: a.role, selected: v.selection[a.role], message: v.message };
  },

  // ---- stage 2.5: Review tab ----

  // The review text from a file (a.path, a Windows path) instead of a paste.
  async "load-review-file"(page, a) {
    await clearMessage(page);
    await page.setInputFiles("#review-file-input", a.path);
    const message = await waitMessage(page, 60_000);
    const chars = await page.$eval("#review-input", (n) => n.value.length);
    return { message, chars };
  },

  async "export-ledger"(page) {
    await openFold(page, "review-ledger-box");
    await clearMessage(page);
    await page.click("#ledger-export");
    await waitMessage(page, 60_000);
    return view(page);
  },

  // a.path: a ledger JSON file (Windows path).
  async "import-ledger"(page, a) {
    await openFold(page, "review-ledger-box");
    await clearMessage(page);
    await page.setInputFiles("#ledger-file-input", a.path);
    await waitMessage(page, 60_000);
    return view(page);
  },

  // Candidate rules from the review text in the paste box (a.review_file
  // fills it first). a.pick: texts; a candidate is ticked when its line holds
  // one of them, and no other is. Without a.pick the pane's own ticks stay.
  // The ticked ones are added unless a.add is false.
  async "suggest-ledger"(page, a) {
    await showPasteBox(page);
    await openFold(page, "review-ledger-box");
    if (a.review_file) await page.fill("#review-input", fs.readFileSync(a.review_file, "utf8"));
    await clearMessage(page);
    await page.click("#ledger-suggest-review");
    const suggest_message = await waitMessage(page, 10 * 60_000, "candidate rule\\(s\\) by ");
    const rows = page.locator("#ledger-candidates label");
    const n = await rows.count();
    const candidates = [];
    for (let i = 0; i < n; i++) {
      const label = (await rows.nth(i).textContent()).replace(/\s+/g, " ").trim();
      const box = rows.nth(i).locator('input[type="checkbox"]');
      if (a.pick) await box.setChecked(a.pick.some((p) => label.includes(p)));
      candidates.push({ label, checked: await box.isChecked() });
    }
    if (a.add === false || !candidates.some((c) => c.checked)) {
      return { suggest_message, candidates, ...(await view(page)), message: "" };
    }
    await clearMessage(page);
    await page.click("#ledger-candidates-add");
    const message = await waitMessage(page, 60_000);
    return { suggest_message, candidates, ...(await view(page)), message };
  },

  // Propose all: every comment that is not done, one turn after another.
  // Waits until the Stop button of the run is gone and its summary shows.
  async "propose-all"(page) {
    if (await turnBusy(page)) throw new Error("a turn is already running");
    const cards = await page.$$eval(".review-comment", (ns) => ns.length);
    await clearMessage(page);
    await page.click("#review-propose-all");
    const order = [];
    const started = Date.now();
    const limit = TURN_TIMEOUT * Math.max(1, cards);
    const running = () =>
      page.evaluate(() => document.getElementById("review-propose-stop")?.hidden === false);
    // The run starts (Stop shows) or ends at once with a message (nothing open).
    await page.waitForFunction(
      () =>
        document.getElementById("review-propose-stop")?.hidden === false ||
        (document.getElementById("review-message")?.textContent.trim() ?? "") !== "",
      null,
      { timeout: 30_000, polling: 200 },
    );
    while (await running()) {
      const m = /Propose all: (\d+)\/(\d+) (\S+)\./.exec(await text(page, "#review-progress"));
      if (m && order[order.length - 1] !== m[3]) order.push(m[3]);
      if (Date.now() - started > limit)
        throw new Error(`Propose all still running after ${limit} ms (at ${order.join(",")})`);
      await sleep(1000);
    }
    // The summary line of the run, not the refresh message of its last turn.
    await page.waitForFunction(
      () =>
        /proposed\.$|^No open comment/.test(
          document.getElementById("review-message")?.textContent.trim() ?? "",
        ),
      null,
      { timeout: 120_000, polling: 200 },
    );
    await sleep(4000);
    return {
      order,
      seconds: Math.round((Date.now() - started) / 1000),
      progress_hidden: await page.$eval("#review-progress", (n) => n.hidden),
      ...(await view(page)),
    };
  },

  // The pane gets the focus back: the status refreshes with no Refresh click.
  // An auto-refresh does not start within 10 s of the one before; a.wait_ms
  // (default 11 s) passes first.
  async "focus-refresh"(page, a) {
    await sleep(a.wait_ms ?? 11_000);
    const before = await text(page, "#review-auto");
    await page.evaluate(() => window.dispatchEvent(new Event("focus")));
    const changed = await page
      .waitForFunction(
        (b) => {
          const t = document.getElementById("review-auto")?.textContent.trim() ?? "";
          return t !== "" && t !== b;
        },
        before,
        { timeout: 120_000, polling: 200 },
      )
      .then(
        () => true,
        () => false,
      );
    return { before, after: await text(page, "#review-auto"), changed, ...(await view(page)) };
  },

  // The snapshot the Refresh sends, taken straight from the pane's own
  // module: paragraph texts as Word reports them and the tracked runs read
  // from the paragraph OOXML.
  async snapshot(page) {
    await page.waitForFunction(
      () => document.getElementById("review-file")?.textContent.startsWith("Review file:"),
      null,
      { timeout: 60_000 },
    );
    const paragraphs = await page.evaluate(async (origin) => {
      const m = await import(`${origin}/shared/tools-word.js`);
      const s = await m.collectReviewSnapshot();
      return s.paragraphs.map((p) => ({
        index: p.index,
        text: p.text,
        w_ins: p.w_ins,
        w_del: p.w_del,
        tracked: p.tracked,
      }));
    }, ORIGIN);
    return { paragraphs };
  },

  // ---- stage 3: Response tab ----

  // Show the tab, then "Refresh from the manuscript": the tab was loaded when
  // the pane connected, before the review was split.
  async "open-response"(page) {
    await page.waitForFunction(
      () => {
        const shown = () =>
          document.querySelector('.tab-content[data-tab="response"]')?.hidden === false;
        if (!shown()) document.querySelector('button.tab[data-tab="response"]')?.click();
        return shown();
      },
      null,
      { timeout: 60_000, polling: 500 },
    );
    await clearResponse(page);
    await page.click("#response-refresh");
    await waitResponse(page, 120_000, "^Refreshed from the manuscript: ");
    return responseView(page);
  },

  "response-view": responseView,

  async "response-refresh"(page) {
    await clearResponse(page);
    await page.click("#response-refresh");
    await waitResponse(page, 120_000, "^Refreshed from the manuscript: ");
    return responseView(page);
  },

  // a.path: an existing response letter (.docx, Windows path). Nothing is
  // saved by this step; rows are the proposal lines.
  async "load-letter"(page, a) {
    await openProfileDetails(page);
    await clearResponse(page);
    await page.setInputFiles("#response-letter-file", a.path);
    await waitResponse(page, 5 * 60_000, "^Read the letter: ");
    const v = await responseView(page);
    return { rows: v.proposal_rows, ...v };
  },

  async "apply-proposal"(page) {
    await openProfileDetails(page);
    await clearResponse(page);
    await page.click("#response-proposal-apply");
    await waitResponse(page, 60_000, "^Applied to this paper: ");
    return responseView(page);
  },

  // a.lang, a.rationale_lang: "" puts the profile's own language back.
  async "response-set-lang"(page, a) {
    await openProfileDetails(page);
    await page.fill("#response-lang", a.lang ?? "");
    await page.fill("#response-rationale-lang", a.rationale_lang ?? "");
    await clearResponse(page);
    await page.click("#response-profile-save");
    await waitResponse(page, 60_000, "^Saved for this paper: profile ");
    return responseView(page);
  },

  // a.comments: [{ ref, reason }] declined as a whole. a.open_requests: every
  // request no accepted change answers is declined with a.request_reason.
  async "response-decline"(page, a) {
    for (const c of a.comments ?? []) {
      const card = responseCard(page, c.ref);
      await card.locator("input.response-decline-reason").fill(c.reason);
      await clearResponse(page);
      await card.locator("button", { hasText: "Decline this comment" }).click();
      await waitResponse(page, 60_000, `^Declined ${escapeRe(c.ref)}\\.`);
    }
    const declined_requests = [];
    const open = page.locator(".response-req:not(.response-req-declined)");
    while (a.open_requests && (await open.count()) > 0) {
      if (declined_requests.length >= 20) throw new Error("more than 20 open requests");
      const row = open.first();
      const label = (await row.locator("span").first().textContent()).trim();
      await row.locator("input.response-req-reason").fill(a.request_reason);
      await clearResponse(page);
      await row.locator("button", { hasText: "Decline" }).click();
      const said = await waitResponse(page, 60_000, "^Declined \\S+ of \\S+\\.$");
      if (!/^Declined /.test(said)) throw new Error(`declining "${label}": ${said}`);
      declined_requests.push(label);
    }
    return { declined_requests, ...(await responseView(page)) };
  },

  async "response-draft"(page, a) {
    await clearResponse(page);
    await responseCard(page, a.ref)
      .locator("button", { hasText: /^Draft$/ })
      .click();
    await waitResponse(page, TURN_TIMEOUT, `^Drafted ${escapeRe(a.ref)}( with .+?)?\\.( |$)`);
    return responseView(page);
  },

  async "response-draft-all"(page) {
    const cards = await page.$$eval(".response-card", (ns) => ns.length);
    await clearResponse(page);
    await page.click("#response-draft-all");
    const limit = TURN_TIMEOUT * Math.max(1, cards);
    await waitResponse(page, limit, "^(Drafted \\d+ responses?( with .+?)?\\.$|Nothing to draft)");
    return responseView(page);
  },

  // a.text: the whole answer; a.prepend: text put in front of the answer that
  // is there; neither: the answer is saved as it is (read and kept).
  async "response-edit"(page, a) {
    const card = responseCard(page, a.ref);
    const box = card.locator("textarea.response-text");
    const previous_text = await box.inputValue();
    if (a.text != null || a.prepend != null)
      await box.fill(a.text != null ? a.text : `${a.prepend}${previous_text}`);
    await clearResponse(page);
    await card.locator("button", { hasText: /^Save$/ }).click();
    await waitResponse(page, 120_000, `^Saved ${escapeRe(a.ref)}\\.$`);
    return { previous_text, ...(await responseView(page)) };
  },

  // Word renders the manuscript to PDF; the daemon reads its line numbers.
  async "response-locate"(page) {
    await clearResponse(page);
    await page.click("#response-locate");
    await waitResponse(page, 10 * 60_000, "^Line numbers updated: ");
    return responseView(page);
  },

  async "response-check"(page) {
    await clearResponse(page);
    await page.click("#response-check");
    await waitResponse(page, 10 * 60_000, "^Checked: ");
    return responseView(page);
  },

  // a.format "text": one plain-text file per reviewer; else the docx. Ends
  // with "Exported..." or, while a check fails, "Not exported: ...".
  async "response-export"(page, a) {
    await clearResponse(page);
    await page.click(a.format === "text" ? "#response-export-text" : "#response-export");
    await waitResponse(page, 10 * 60_000, "^(Exported|Not exported: )");
    return responseView(page);
  },

  // ---- stage 4: Tools tab ----

  async "open-tools"(page) {
    await page.waitForFunction(
      () => {
        const shown = () =>
          document.querySelector('.tab-content[data-tab="tools"]')?.hidden === false;
        if (!shown()) document.querySelector('button.tab[data-tab="tools"]')?.click();
        return shown();
      },
      null,
      { timeout: 60_000, polling: 500 },
    );
    await page.waitForFunction(
      () => {
        const line = document.getElementById("tools-doc")?.textContent ?? "";
        return line.startsWith("Saved file:") || line.startsWith("Save the document");
      },
      null,
      { timeout: 60_000 },
    );
    return toolsView(page);
  },

  "tools-view": toolsView,

  // a.paragraph: the paragraph to select in Word; a.lang: the target language.
  async "tools-translate"(page, a) {
    await selectParagraph(page, a.paragraph);
    await clearTools(page, "tools-selection");
    await page.fill("#tools-lang", a.lang);
    await page.click("#tools-translate");
    await waitTools(page, 5 * 60_000, { box: "tools-selection" });
    return toolsView(page);
  },

  // a.paragraph, a.question.
  async "tools-ask"(page, a) {
    await selectParagraph(page, a.paragraph);
    await clearTools(page, "tools-selection");
    await page.fill("#tools-question", a.question);
    await page.click("#tools-ask");
    await waitTools(page, 5 * 60_000, { box: "tools-selection" });
    return toolsView(page);
  },

  // a.guideline_file: the guideline text (a Windows path). Nothing is stored
  // by this step; rules are the candidates with the sentence each cites.
  async "tools-extract"(page, a) {
    await clearTools(page, "tools-rules");
    await page.fill("#tools-guideline", fs.readFileSync(a.guideline_file, "utf8"));
    await page.click("#tools-rules-extract");
    await waitTools(page, 5 * 60_000, { until: "rule candidates\\.|^No rule was found" });
    return toolsView(page);
  },

  // a.keep (a RegExp source): the candidates whose cited sentence matches
  // stay ticked, every other one is unticked; then "Add selected".
  async "tools-add"(page, a) {
    const kept = await page.evaluate((keep) => {
      const re = new RegExp(keep, "i");
      let n = 0;
      for (const l of document.querySelectorAll("#tools-rules label") ?? []) {
        const box = l.querySelector("input");
        const stays = re.test(l.querySelector(".review-preview")?.textContent ?? "");
        if (box) box.checked = stays;
        if (stays) n++;
      }
      return n;
    }, a.keep);
    await clearTools(page);
    await page.click("#tools-rules-add");
    await waitTools(page, 60_000, { until: "added to the ledger" });
    return { kept, ...(await toolsView(page)) };
  },

  async "tools-check"(page) {
    await clearTools(page, "tools-checks");
    await page.click("#tools-check");
    await waitTools(page, 5 * 60_000, {
      until: "findings\\. Nothing was changed\\.$|^The checks found nothing\\.$",
    });
    return toolsView(page);
  },

  // a.path: the earlier version (.docx, Windows path); a.mode: track_changes
  // or color.
  async "tools-marked"(page, a) {
    await clearTools(page, "tools-marked");
    await page.setInputFiles("#tools-baseline", a.path);
    await page.selectOption("#tools-marked-mode", a.mode);
    await page.click("#tools-marked-build");
    await waitTools(page, 10 * 60_000, { box: "tools-marked" });
    return toolsView(page);
  },

  async "tools-anonymize"(page) {
    await clearTools(page, "tools-anonymized");
    await page.click("#tools-anonymize");
    await waitTools(page, 10 * 60_000, { box: "tools-anonymized" });
    return toolsView(page);
  },

  view,
};

// Loaded by a test, the steps are only handed over; run as a program, one
// step is done.
module.exports = { steps, MODELS_RETRY_AFTER_MS };
if (require.main === module) {
  (async () => {
    const [step, argsFile] = process.argv.slice(2);
    const fn = steps[step];
    if (!fn) throw new Error(`unknown step ${step}`);
    const a = argsFile ? JSON.parse(fs.readFileSync(argsFile, "utf8")) : {};
    const { browser, page } = await connect();
    const shot = process.env.E2E_SHOT;
    const waiting = [];
    const timers = [];
    // While the step runs nothing is scrolled: the picture is what the user
    // was left with after the click. One picture at a time (the whole-tab
    // picture changes the layout for a moment and puts it back).
    let pictures = Promise.resolve();
    const during = (n, ms) =>
      timers.push(
        setTimeout(() => {
          const file = shot.replace(/\.png$/, `-wait${n}.png`);
          const row = { after_ms: ms };
          waiting.push(row);
          pictures = pictures
            .then(() => paneState(page))
            .then((state) => Object.assign(row, state))
            .then(() => progressInView(page))
            .then((seen) => (row.progress_in_view = seen))
            .then(() => takeShot(page, file))
            .then((files) => Object.assign(row, files))
            .catch((e) => (row.error = String(e?.message ?? e)));
        }, ms),
      );
    if (shot) {
      during(1, 4000);
      during(2, 20_000);
    }
    const after = async () => {
      for (const t of timers) clearTimeout(t);
      if (!shot) return {};
      await pictures;
      const taken = { shot: null, shot_full: null, shot_error: null, look_at: null };
      await (async () => {
        taken.viewport = await viewportOf(page);
        taken.look_at = await lookAt(page, a.ref);
        Object.assign(taken, await takeShot(page, shot));
      })().catch((e) => (taken.shot_error = String(e?.message ?? e)));
      return { ...taken, waiting };
    };
    try {
      let out;
      try {
        out = await fn(page, a);
      } catch (e) {
        const taken = await after();
        process.stdout.write(
          JSON.stringify({ ok: false, error: String(e?.message ?? e), ...taken }) + "\n",
        );
        process.exitCode = 1;
        return;
      }
      process.stdout.write(JSON.stringify({ ok: true, step, ...out, ...(await after()) }) + "\n");
    } finally {
      await browser.close().catch(() => {});
    }
  })().catch((e) => {
    process.stdout.write(JSON.stringify({ ok: false, error: String(e?.message ?? e) }) + "\n");
    process.exit(1);
  });
}
