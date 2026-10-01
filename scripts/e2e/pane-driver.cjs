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

const steps = {
  async "open-review"(page) {
    await page.click('button.tab[data-tab="review"]');
    await page.waitForFunction(
      () => document.getElementById("review-file").textContent.startsWith("Review file:"),
      null,
      { timeout: 60_000 },
    );
    return view(page);
  },

  async split(page, a) {
    await page.fill("#review-input", fs.readFileSync(a.review_file, "utf8"));
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
