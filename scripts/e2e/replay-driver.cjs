// The pane steps of replay.mjs: every step of pane-driver.cjs (which must be
// in the same folder) and three more. Runs on Windows node, one step per
// process, one JSON line out, as pane-driver.cjs does.
//
//   node replay-driver.cjs <step> <args.json>
//
// Environment: as pane-driver.cjs (PLAYWRIGHT_CORE, E2E_CDP_PORT,
// E2E_PANE_ORIGIN, E2E_TURN_TIMEOUT_MS).

const fs = require("node:fs");
const base = require("./pane-driver.cjs");

const PW = process.env.PLAYWRIGHT_CORE;
const PORT = process.env.E2E_CDP_PORT || "9333";
const ORIGIN = process.env.E2E_PANE_ORIGIN || "http://127.0.0.1:47834";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// As in pane-driver.cjs (it does not export its connect).
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

const turnBusy = (page) =>
  page.evaluate(() => {
    const p = document.getElementById("review-progress");
    return !!document.getElementById("input")?.disabled || p?.hidden === false;
  });

const steps = {
  ...base.steps,

  // a.model: an alias or model id of the model menu under the composer. A
  // model the menu does not list is not chosen: `refused` says so and lists
  // what the menu has.
  async "set-model"(page, a) {
    const menu = await base.steps.models(page);
    const values = menu.options.map((o) => o.value);
    if (!values.includes(a.model)) {
      return { selected: menu.selected, refused: `the model menu has no ${a.model}`, values };
    }
    if (menu.selected !== a.model) {
      await page.selectOption("#composer-model", a.model);
      await page.waitForFunction(
        (want) => document.getElementById("composer-model")?.value === want,
        a.model,
        { timeout: 30_000 },
      );
    }
    const selected = await page.evaluate(
      () => document.getElementById("composer-model")?.value ?? null,
    );
    return { selected, values };
  },

  // The headings of the document, from the pane's own module (what the
  // Response tab reads): [{ index, level, text }].
  async headings(page) {
    const headings = await page.evaluate(async (origin) => {
      const m = await import(`${origin}/shared/tools-word.js`);
      return m.collectHeadings();
    }, ORIGIN);
    return { headings };
  },

  // After a propose step that gave up: press Stop when a turn still runs and
  // wait until the pane is free, so the next comment can start.
  async "stop-turn"(page, a) {
    if (!(await turnBusy(page))) return { stopped: false, free: true };
    await page.click("#stop-agent");
    const started = Date.now();
    while (await turnBusy(page)) {
      if (Date.now() - started > (a.wait_ms ?? 120_000)) return { stopped: true, free: false };
      await sleep(500);
    }
    await sleep(3000);
    return { stopped: true, free: true };
  },
};

module.exports = { steps };
if (require.main === module) {
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
}
