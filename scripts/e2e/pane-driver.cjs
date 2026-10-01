// Drives the Draftspect Word task pane over the Chrome DevTools Protocol.
// Runs on Windows node (the pane is a WebView2 inside Word). One step per
// process: the runner calls it once per step and reads one JSON line.
//
//   node pane-driver.cjs <step> <args.json>
//
// Environment: PLAYWRIGHT_CORE (path of a playwright-core package),
// E2E_CDP_PORT (the --remote-debugging-port given to Word),
// E2E_PANE_ORIGIN (default http://127.0.0.1:47834).
// Element ids come from taskpane/word/index.html and the class names from
// taskpane/shared/review-pane.js.

const fs = require("node:fs");

const PW = process.env.PLAYWRIGHT_CORE;
const PORT = process.env.E2E_CDP_PORT || "9333";
const ORIGIN = process.env.E2E_PANE_ORIGIN || "http://127.0.0.1:47834";
const TURN_TIMEOUT = Number(process.env.E2E_TURN_TIMEOUT_MS || 20 * 60_000);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

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
    m.textContent = "";
    m.hidden = true;
  });
}
async function waitMessage(page, timeout = 120_000) {
  await page.waitForFunction(
    () => document.getElementById("review-message").textContent.trim() !== "",
    null,
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
      message_error: document.getElementById("review-message").classList.contains("form-error"),
      file: t(document.getElementById("review-file")),
      comments: q(".review-comment").map((c) => ({
        ref: t(c.querySelector(".review-comment-head strong")),
        reqs_line: t(c.querySelector(".review-reqs-line")),
        changes: q(".review-change-wrap", c).map((w) => ({
          badge: t(w.querySelector(".review-badge")),
          text: t(w.querySelector(".review-change")),
          flags: q(".review-flag", w).map(t),
        })),
        blocked: q(".review-blocked", c).map(t),
      })),
      decisions: q("#review-decisions > div").map(t),
      external: q("#review-external .review-external-item .review-violation").map(t),
      ledger: q("#review-ledger .review-ledger-item").map(t),
      violations: q("#review-violations .review-violation").map(t),
      agent_status: t(document.getElementById("agent-status")),
      delete_all_button: t(document.getElementById("review-delete-all")),
    };
  });
}

function card(page, ref) {
  return page.locator(".review-comment", {
    has: page.locator(".review-comment-head strong", {
      hasText: new RegExp(`^${ref.replace(".", "\\.")}$`),
    }),
  });
}

const turnBusy = (page) =>
  page.evaluate(() => {
    const p = document.getElementById("review-progress");
    return document.getElementById("input").disabled || !p.hidden;
  });

// Providers section of the Setup tab (ids from taskpane/word/index.html,
// rows from taskpane/shared/providers-pane.js). The list is asked for on
// connect; wait until it is drawn or its message says why it is not.
async function openProviders(page) {
  await page.click('button.tab[data-tab="setup"]');
  await page.waitForFunction(
    () =>
      document.querySelectorAll("#providers-list .provider-row").length > 0 ||
      document.getElementById("providers-message").textContent.trim() !== "",
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
    const raw = document.getElementById("providers-root").getAttribute("data-selection");
    return {
      menus: { turn: menu("provider-turn"), oneshot: menu("provider-oneshot") },
      models: { turn: menu("provider-turn-model"), oneshot: menu("provider-oneshot-model") },
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

const steps = {
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
      () => document.getElementById("review-file").textContent.startsWith("Review file:"),
      null,
      { timeout: 60_000 },
    );
    return view(page);
  },

  // a.review_file fills the paste box first; without it the text already
  // there is split (stage 2.5 loads it with load-review-file).
  async split(page, a) {
    if (a.review_file) await page.fill("#review-input", fs.readFileSync(a.review_file, "utf8"));
    await clearMessage(page);
    await page.click("#review-split");
    const message = await waitMessage(page, 10 * 60_000);
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
          document.getElementById("review-message").textContent.trim() !== "",
        before,
        { timeout: 60_000 },
      );
    }
    return view(page);
  },

  async "check-all"(page) {
    await clearMessage(page);
    await page.click("#review-check-all");
    // Done when no ledger item waits for its first check, or an error shows.
    await page.waitForFunction(
      () =>
        !document.getElementById("review-decisions").textContent.includes("not checked yet") ||
        document.getElementById("review-message").textContent.trim() !== "",
      null,
      { timeout: 10 * 60_000, polling: 500 },
    );
    await sleep(500);
    return view(page);
  },

  async preview(page, a) {
    await card(page, a.ref).locator("button", { hasText: "Preview prompt" }).click();
    await page.waitForFunction(() => !document.getElementById("review-preview").hidden, null, {
      timeout: 60_000,
    });
    return { prompt: await text(page, "#review-preview") };
  },

  // Propose changes on one comment and wait until the turn is over.
  async propose(page, a) {
    if (await turnBusy(page)) throw new Error("a turn is already running");
    const btn = card(page, a.ref).locator("button", { hasText: "Propose changes" });
    await btn.click();
    await page.waitForFunction(() => document.getElementById("input").disabled, null, {
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

  async "clean-tool"(page) {
    await clearMessage(page);
    await page.click("#review-clean-tool");
    await waitMessage(page, 120_000);
    return view(page);
  },

  // First press shows the counts on the button; second press deletes.
  async "delete-all"(page) {
    const btn = "#review-delete-all";
    await clearMessage(page);
    await page.click(btn);
    await page.waitForFunction(
      () =>
        document.getElementById("review-delete-all").textContent.startsWith("Back up and delete") ||
        document.getElementById("review-message").textContent.trim() !== "",
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
  async models(page) {
    await page.waitForFunction(
      () =>
        /^(Models your Claude account|Model list unavailable)/.test(
          document.getElementById("composer-model").title,
        ),
      null,
      { timeout: 90_000, polling: 200 },
    );
    return page.$eval("#composer-model", (m) => ({
      title: m.title,
      selected: m.value,
      options: [...m.options].map((o) => ({ value: o.value, label: o.textContent.trim() })),
    }));
  },

  // Propose changes on one comment, then press Stop while the turn runs
  // (a.after_ms after it began, default 8 s) and wait until the pane is free.
  async "propose-stop"(page, a) {
    if (await turnBusy(page)) throw new Error("a turn is already running");
    await card(page, a.ref).locator("button", { hasText: "Propose changes" }).click();
    await page.waitForFunction(() => document.getElementById("input").disabled, null, {
      timeout: 30_000,
    });
    const started = Date.now();
    await sleep(a.after_ms ?? 8000);
    if (!(await turnBusy(page))) {
      return { stopped: false, reason: "the turn ended before Stop", ...(await view(page)) };
    }
    await page.click("#stop-agent");
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
          const m = document.getElementById("providers-message").textContent.trim();
          const raw = document.getElementById("providers-root").getAttribute("data-selection");
          const s = raw ? JSON.parse(raw)[role] : null;
          return m !== "" || (s && s.id === w.id && (w.model == null || s.model === w.model));
        },
        [a.role, want],
        { timeout: 90_000, polling: 200 },
      );
    await page.evaluate(() => (document.getElementById("providers-message").textContent = ""));
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
    await clearMessage(page);
    await page.click("#ledger-export");
    await waitMessage(page, 60_000);
    return view(page);
  },

  // a.path: a ledger JSON file (Windows path).
  async "import-ledger"(page, a) {
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
    if (a.review_file) await page.fill("#review-input", fs.readFileSync(a.review_file, "utf8"));
    await clearMessage(page);
    await page.click("#ledger-suggest-review");
    const suggest_message = await waitMessage(page, 10 * 60_000);
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
      page.evaluate(() => !document.getElementById("review-propose-stop").hidden);
    // The run starts (Stop shows) or ends at once with a message (nothing open).
    await page.waitForFunction(
      () =>
        !document.getElementById("review-propose-stop").hidden ||
        document.getElementById("review-message").textContent.trim() !== "",
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
          document.getElementById("review-message").textContent.trim(),
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
          const t = document.getElementById("review-auto").textContent.trim();
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
      () => document.getElementById("review-file").textContent.startsWith("Review file:"),
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

  view,
};

(async () => {
  const [step, argsFile] = process.argv.slice(2);
  const fn = steps[step];
  if (!fn) throw new Error(`unknown step ${step}`);
  const a = argsFile ? JSON.parse(fs.readFileSync(argsFile, "utf8")) : {};
  const { browser, page } = await connect();
  try {
    const out = await fn(page, a);
    process.stdout.write(JSON.stringify({ ok: true, step, ...out }) + "\n");
  } finally {
    await browser.close().catch(() => {});
  }
})().catch((e) => {
  process.stdout.write(JSON.stringify({ ok: false, error: String(e?.message ?? e) }) + "\n");
  process.exit(1);
});
