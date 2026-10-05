// run-acceptance.mjs --screenshots: the names of the pictures, the rows of
// shots/timings.json, and that the option changes nothing unless it is given.
// The pictures themselves need Word and are not made here.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { MODEL_STEPS, shotName, stepTiming } from "../scripts/e2e/judge.mjs";

const read = (p) => readFileSync(new URL(p, import.meta.url), "utf8");

test("the picture of a step is named by its number, the step and the comment or tab it was for", () => {
  assert.equal(shotName(7, "propose", { ref: "R1.1" }), "07-propose-R1.1.png");
  assert.equal(shotName(3, "show-tab", { tab: "tools" }), "03-show-tab-tools.png");
  assert.equal(shotName(12, "mode14-refresh"), "12-mode14-refresh.png");
  assert.equal(shotName(104, "refresh", {}), "104-refresh.png");
  // Nothing else of the arguments gets into a file name: no path, no text.
  assert.equal(
    shotName(5, "split", { review_file: "D:\\x\\review.txt", ref: "../x" }),
    "05-split.png",
  );
});

test("a row of timings.json: how long, what the pane said meanwhile, and whether a model was waited for", () => {
  const row = stepTiming({
    n: 9,
    step: "propose",
    args: { ref: "R1.2" },
    shot: "09-propose-R1.2.png",
    seconds: 48.3,
    result: {
      ok: true,
      seconds: 44,
      message: "Proposed 1 change for R1.2.",
      progress_samples: ["Review R1.2: reading the paragraphs..."],
      shot: "D:\\run\\shots\\09-propose-R1.2.png",
      shot_full: "D:\\run\\shots\\09-propose-R1.2-full.png",
      shot_error: null,
      look_at: "card",
      viewport: { width: 648, height: 851, device_pixel_ratio: 1.5 },
      waiting: [
        {
          after_ms: 4000,
          progress_in_view: false,
          shot: "D:\\run\\shots\\09-propose-R1.2-wait1.png",
          shot_full: "D:\\run\\shots\\09-propose-R1.2-wait1-full.png",
          message: "",
          progress: "Review R1.2: 3 s",
          agent_status: "Working",
          input_disabled: true,
        },
        { after_ms: 20000, error: "Target closed" },
      ],
    },
  });
  assert.deepEqual(row, {
    n: 9,
    step: "propose",
    ref: "R1.2",
    model_call: true,
    ok: true,
    seconds: 48.3,
    turn_seconds: 44,
    shot: "09-propose-R1.2.png",
    shot_full: "09-propose-R1.2-full.png",
    shot_error: null,
    look_at: "card",
    viewport: { width: 648, height: 851, device_pixel_ratio: 1.5 },
    waiting: [
      {
        after_ms: 4000,
        progress_in_view: false,
        shot: "09-propose-R1.2-wait1.png",
        shot_full: "09-propose-R1.2-wait1-full.png",
        message: "",
        progress: "Review R1.2: 3 s",
        agent_status: "Working",
        input_disabled: true,
      },
      { after_ms: 20000, error: "Target closed" },
    ],
    progress: ["Review R1.2: reading the paragraphs..."],
    message: "Proposed 1 change for R1.2.",
  });
  // A step of the second Word run is the same step; one that failed says why;
  // a picture that could not be taken is not named.
  const failed = stepTiming({
    n: 40,
    step: "mode14-propose",
    args: { ref: "R1.1" },
    shot: "40-mode14-propose-R1.1.png",
    seconds: 600,
    result: { ok: false, error: "turn R1.1 still running", shot: null, shot_error: "timeout" },
  });
  assert.deepEqual(
    [failed.model_call, failed.ok, failed.message, failed.shot, failed.shot_error],
    [true, false, "turn R1.1 still running", null, "timeout"],
  );
  const tab = stepTiming({
    n: 1,
    step: "show-tab",
    args: { tab: "chat" },
    shot: "a.png",
    seconds: 4,
    result: {},
  });
  assert.deepEqual(
    [tab.model_call, tab.tab, tab.turn_seconds, tab.waiting],
    [false, "chat", null, []],
  );
  for (const s of ["refresh", "show-tab", "allow-exception", "edit-draft", "ledger-add"])
    assert.ok(!MODEL_STEPS.has(s), s);
});

test("the driver has the steps of a screenshot run and takes the pictures only when asked", () => {
  const { steps } = createRequire(import.meta.url)("../scripts/e2e/pane-driver.cjs");
  for (const s of ["show-tab", "refresh-offline", ...MODEL_STEPS]) assert.ok(steps[s], s);
  const driver = read("../scripts/e2e/pane-driver.cjs");
  assert.match(driver, /const shot = process\.env\.E2E_SHOT;/);
  assert.match(driver, /if \(!shot\) return \{\};/);
  assert.match(driver, /if \(shot\) \{\s*during\(1, 4000\);\s*during\(2, 20_000\);\s*\}/);
  // A step that failed is photographed too, and still ends as a failure.
  const tail = driver.slice(driver.indexOf("if (require.main === module)"));
  assert.ok(tail.indexOf("const taken = await after();") < tail.indexOf("process.exitCode = 1;"));
  // Two files a picture: the window as it is (never a full-page picture of
  // it: the pane scrolls inside an element, so that would be the window
  // again), and the whole tab with the layout grown and put back.
  assert.equal((driver.match(/page\s*\.screenshot\(/g) ?? []).length, 2);
  assert.match(driver, /await page\.screenshot\(\{ path: file \}\);/);
  const full = driver.slice(
    driver.indexOf("async function fullShot("),
    driver.indexOf("async function takeShot("),
  );
  assert.match(
    full,
    /try \{\s*await page\.screenshot\(\{ path: file, fullPage: true \}\);\s*\} finally \{/,
  );
  assert.match(full, /u\.e\.removeAttribute\("style"\)/);
  assert.match(full, /u\.e\.scrollTop = u\.top/);
  assert.match(full, /delete window\.__shotUndo/);
  // One picture at a time, and the pictures of the wait are over before the last one.
  assert.match(tail, /pictures = pictures\s*\.then\(/);
  assert.ok(tail.indexOf("await pictures;") < tail.indexOf("await takeShot(page, shot)"));
  // During the wait nothing is scrolled; after the step what it was for is.
  const wait = tail.slice(tail.indexOf("const during ="), tail.indexOf("const after ="));
  assert.doesNotMatch(wait, /lookAt\(/);
  assert.match(wait, /row\.progress_in_view = seen/);
  assert.ok(tail.indexOf("await lookAt(page, a.ref)") < tail.indexOf("await takeShot(page, shot)"));
});

test("the runner: nothing of it happens without --screenshots, and nothing of it is judged", () => {
  const runner = read("../scripts/e2e/run-acceptance.mjs");
  assert.match(runner, /screenshots: \{ type: "boolean", default: false \}/);
  assert.match(runner, /const shot = opt\.screenshots\s*\? path\.join\(SHOTS, shotName\(/);
  assert.match(runner, /\.\.\.\(shot \? \{ E2E_SHOT: winPath\(shot\) \} : \{\}\)/);
  assert.match(
    runner,
    /const showTab = \(tab\) => \(opt\.screenshots \? pane\("show-tab", \{ tab \}\) : null\);/,
  );
  assert.match(runner, /function wordWindowShot\(name\) \{\s*if \(!opt\.screenshots\) return;/);
  // The steps that exist only for the pictures are never noted as steps of the run.
  for (const step of ["show-tab", "refresh-offline"])
    assert.doesNotMatch(runner, new RegExp(`note\\([^)]*pane\\("${step}"`));
  assert.doesNotMatch(runner, /note\([^\n]*showTab\(/);
  // Stage 2: the tabs before anything is pasted; the daemon is stopped for the
  // picture only after the verify commands, and started again before the cleanup.
  const main = runner.slice(runner.indexOf("async function main() {"));
  assert.ok(
    main.indexOf('"library", "setup"]) showTab(tab);') < main.indexOf('pane("open-review")'),
  );
  const end = runner.slice(runner.lastIndexOf('if (STAGE === "s2") {'));
  const order = [
    "verify();",
    "await stopDaemon()",
    'pane("refresh-offline");',
    "await startDaemon(daemonLog",
    "await tidy();",
  ].map((x) => end.indexOf(x));
  assert.ok(order.every((x) => x >= 0));
  assert.deepEqual(
    order,
    [...order].sort((a, b) => a - b),
  );
  assert.match(runner, /"word-launch\.ps1",\s*"word-shot\.ps1",?\s*\]/);
});

test("word-shot.ps1 only reads the window of the process it is given", () => {
  const ps = read("../scripts/e2e/word-shot.ps1");
  // eslint-disable-next-line no-control-regex
  assert.doesNotMatch(ps, /[^\x00-\x7f]/);
  assert.match(ps, /\$proc\.ProcessName -ne "WINWORD"/);
  assert.match(ps, /PrintWindow\(\$handle, \$hdc, 2\)/);
  // The real size of the window is read (a scaled display gave the top left
  // corner only), and a picture with a large black area is not written.
  assert.ok(ps.indexOf("SetProcessDPIAware()") < ps.indexOf("GetWindowRect($handle"));
  assert.match(ps, /\$dropped = \(-not \$drawn\) -or \(\$share -gt \$MaxBlackShare\)/);
  assert.match(ps, /if \(-not \$dropped\) \{ \$bitmap\.Save\(/);
  assert.equal((ps.match(/\.Save\(/g) ?? []).length, 1);
  // Nothing that moves, shows, focuses, closes or types into a window, and
  // no picture of the screen (another window could be in it).
  for (const call of [
    "SetForegroundWindow",
    "ShowWindow",
    "SetWindowPos",
    "MoveWindow",
    "SendKeys",
    "CopyFromScreen",
    "CloseMainWindow",
    "Stop-Process",
    ".Kill(",
  ])
    assert.ok(!ps.includes(call), call);
  // PowerShell does not tell $Doc from $doc: one spelling per variable.
  const spellings = new Map();
  for (const [, name] of ps.matchAll(/\$([A-Za-z_][A-Za-z0-9_]*)/g)) {
    const key = name.toLowerCase();
    spellings.set(key, (spellings.get(key) ?? new Set()).add(name));
  }
  for (const [key, set] of spellings) assert.equal(set.size, 1, `$${key}: ${[...set].join(", ")}`);
});
