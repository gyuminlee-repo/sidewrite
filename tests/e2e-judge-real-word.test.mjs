// The steps of the acceptance runner that put in real Word what had only been
// checked without it: "Check now" of "Before submission", the "Figures and
// other files" button, and a run on the resident daemon (--resident). Their
// judges are pure and are fed here with inputs written by hand in the form
// the pane driver, word-com.ps1, browser-windows.ps1 and the daemon give
// them; none of these inputs was recorded from Word. The flows need Word and
// do not run here; what they ask of the driver and the scripts must exist.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import * as J from "../scripts/e2e/judge.mjs";
import { UNSAVED_TEXT, UNSAVED_UNKNOWN } from "../daemon/submission-check.mjs";

const read = (p) => readFileSync(new URL(p, import.meta.url), "utf8");
const byId = (rows) => Object.fromEntries(rows.map((r) => [r.id, r]));

// ---- "Check now" ---------------------------------------------------------

const LETTER = "/mnt/d/run/review_manuscript.response.docx";
const SHA = "a".repeat(64);
const revision = (extra = {}) => ({
  comments: [{ ref: "R1.1", response: { text: "Done." }, figure_changes: [] }],
  response_export: { path: LETTER, fail: 0, warn: 0, model_sha: "m", docx_sha: SHA },
  artifacts: {},
  ...extra,
});
const NOTES = [
  { id: "answers", state: "ready", text: "Every comment has an answer that passes the checks." },
  { id: "requests", state: "ready", text: "No request is open and no question to you is waiting." },
  {
    id: "format_counts",
    state: "not_applicable",
    text: "No format rule and no count rule is set for this paper.",
  },
];
// The rows as the pane draws them for a document with no revision and no
// comment, the letter exported and unchanged.
function paneRows({ saved, letter = "ready" }) {
  const copy = (what) =>
    saved
      ? `No ${what} is on record for this file. Finish the cleanup above and save the manuscript first (removing comments changes the saved file), then make the copy.`
      : UNSAVED_TEXT;
  return [
    ...NOTES.slice(0, 2),
    {
      id: "changes_decided",
      state: "ready",
      text: "No tracked change is waiting for a decision in this file.",
    },
    { id: "tool_comments", state: "ready", text: "The tool's comments are removed." },
    {
      id: "file_clean",
      state: "ready",
      text: "This file is clean: no tracked changes and no comments are left.",
    },
    NOTES[2],
    letter === "ready"
      ? {
          id: "letter",
          state: "ready",
          text: "The exported letter says what the answers and the accepted changes say now: review_manuscript.response.docx.",
        }
      : {
          id: "letter",
          state: "not_ready",
          text: "The answers, the accepted changes or the line numbers changed after review_manuscript.response.docx was exported. Export the letter again.",
        },
    { id: "marked_copy", state: "not_ready", text: copy("marked copy") },
    {
      id: "anonymized_copy",
      state: "not_applicable",
      text: "The journal profile does not say that an anonymized copy is needed.",
    },
    {
      id: "figures",
      state: "not_applicable",
      text: "No figure or file change is on record for this paper.",
    },
  ];
}
const press = ({ saved, letter } = { saved: false }) => ({
  ok: true,
  message: "",
  message_error: false,
  rows: paneRows({ saved, letter }),
  asof: saved
    ? "review_manuscript.docx, checked at 10:01. Press Check now again after a change."
    : "review_manuscript.docx, checked at 10:00. This document has changes that are not saved: save it, then press Check now again.",
});
const logOf = (saved, n = 11) =>
  `[tools] submission_check rows=${n} ready=5 not_ready=2 not_applicable=4 saved=${saved}\n`;
const word = (saved, extra = {}) => ({ ok: true, saved, revisions: [], comments: [], ...extra });
const good = (saved, letter = "ready") => ({
  step: press({ saved, letter }),
  word: word(saved),
  revision: revision(),
  letterSha: SHA,
  letter,
  log: logOf(saved),
});

test("check now: the texts the judge expects are the ones the daemon writes", () => {
  const src = read("../daemon/submission-check.mjs");
  assert.match(src, /`No \$\{what\} is on record for this file\. /);
  assert.match(src, /changed after \$\{name\} was exported/);
  assert.match(src, /text: `The exported letter says what the answers .*: \$\{name\}\.`/);
  assert.match(read("../taskpane/shared/review-pane.js"), /Press Check now again after a change\./);
  assert.equal(
    J.submissionLogLine(logOf(false)).saved,
    false,
    "the line aux-service.mjs writes is read",
  );
  assert.match(
    read("../daemon/aux-service.mjs"),
    /\[tools\] submission_check rows=\$\{rows\.length\}/,
  );
  assert.equal(J.submissionLogLine("nothing"), null);
});

test("check now: an unsaved document, then the saved one, then the letter out of date", () => {
  for (const [saved, letter] of [
    [false, "ready"],
    [true, "ready"],
    [true, "changed"],
  ]) {
    const ev = J.submissionEvidence(good(saved, letter));
    const fails = ev.rows.filter((r) => r.result === "FAIL");
    assert.deepEqual(fails, [], `${saved} ${letter}`);
    assert.equal(ev.pass, true);
    const rows = byId(ev.rows);
    for (const id of [
      "job",
      "saved",
      "changes_decided",
      "tool_comments",
      "file_clean",
      "marked_copy",
      "anonymized_copy",
      "letter",
      "figures",
      "as_of",
    ])
      assert.equal(rows[id].result, "PASS", id);
    // What the run cannot tell by itself is told, not judged.
    for (const id of ["answers", "requests", "format_counts"])
      assert.equal(rows[id].result, "NOTE");
    // The evidence of each row is the row as the pane shows it.
    assert.match(rows.letter.got, /^(ready|not_ready): /);
  }
});

// One input per judged row that must fail it, and only it where possible.
test("check now: every judged row fails on an input that is wrong for it", () => {
  const failOf = (input, id) => {
    const ev = J.submissionEvidence(input);
    assert.equal(ev.pass, false, id);
    assert.equal(byId(ev.rows)[id].result, "FAIL", `${id}: ${JSON.stringify(byId(ev.rows)[id])}`);
    return byId(ev.rows)[id];
  };
  const swap = (step, id, row) => ({
    ...step,
    rows: step.rows.map((r) => (r.id === id ? { ...r, ...row } : r)),
  });

  // job: the pane showed an error, or the daemon logged no job.
  failOf({ ...good(false), step: { ...press(), message: "x", message_error: true } }, "job");
  failOf({ ...good(false), log: "" }, "job");
  // saved: the pane sent another answer than Word gives over COM.
  failOf({ ...good(false), log: logOf(true) }, "saved");
  // changes_decided: ready while Word lists a revision.
  failOf({ ...good(false), word: word(false, { revisions: [{ type: 1 }] }) }, "changes_decided");
  // tool_comments: ready while a rationale comment of the tool is in Word.
  const toolComment = { text: "[R1.1] Reworded as asked." };
  const withTool = { ...good(false), word: word(false, { comments: [toolComment] }) };
  failOf(withTool, "tool_comments");
  // file_clean: ready while a comment of a person is in Word.
  failOf(
    { ...good(false), word: word(false, { comments: [{ text: "Check this." }] }) },
    "file_clean",
  );
  // marked_copy: the unsaved text after the save, the record's text before it.
  failOf(
    { ...good(true), step: swap(press({ saved: true }), "marked_copy", { text: UNSAVED_TEXT }) },
    "marked_copy",
  );
  failOf(
    {
      ...good(false),
      step: swap(press(), "marked_copy", { text: "No marked copy is on record for this file." }),
    },
    "marked_copy",
  );
  // A saved document whose pane could not say so is not the saved state.
  failOf(
    {
      ...good(true),
      step: swap(press({ saved: true }), "marked_copy", {
        text: `No marked copy is on record for this file. ${UNSAVED_UNKNOWN}`,
      }),
    },
    "marked_copy",
  );
  // Word did not say whether it is saved.
  failOf({ ...good(true), word: word(undefined) }, "marked_copy");
  // anonymized_copy: needed by the profile, yet "ready" while unsaved.
  failOf(
    {
      ...good(false),
      step: swap(press(), "anonymized_copy", {
        state: "ready",
        text: "The anonymized copy was made ...",
      }),
    },
    "anonymized_copy",
  );
  // letter: ready, but the file on the disk is not the one written.
  const notAsWritten = failOf(
    {
      ...good(true),
      letterSha: "b".repeat(64),
    },
    "letter",
  );
  assert.match(notAsWritten.expect, /not as written/);
  // letter: ready although the record says a check failed.
  failOf(
    {
      ...good(true),
      revision: revision({ response_export: { ...revision().response_export, fail: 2 } }),
    },
    "letter",
  );
  // letter: ready although the answers changed since the export.
  failOf({ ...good(true, "changed"), step: press({ saved: true, letter: "ready" }) }, "letter");
  // letter: out of date although nothing changed (a hash that is not stable).
  failOf({ ...good(true, "ready"), step: press({ saved: true, letter: "changed" }) }, "letter");
  // letter: ready, but naming another file.
  failOf(
    {
      ...good(true),
      step: swap(press({ saved: true }), "letter", {
        text: "The exported letter says ...: old.docx.",
      }),
    },
    "letter",
  );
  // letter: the row is missing.
  failOf(
    {
      ...good(true),
      step: {
        ...press({ saved: true }),
        rows: press({ saved: true }).rows.filter((r) => r.id !== "letter"),
      },
    },
    "letter",
  );
  // figures: none on record, yet the row asks for something.
  failOf(
    { ...good(true), step: swap(press({ saved: true }), "figures", { state: "not_ready" }) },
    "figures",
  );
  // as_of: the line under the rows says the other state.
  failOf({ ...good(false), step: { ...press(), asof: press({ saved: true }).asof } }, "as_of");
  failOf({ ...good(true), step: { ...press({ saved: true }), asof: press().asof } }, "as_of");
});

test("check now: a row of a supplementary file with the same id is told, not judged", () => {
  const step = press({ saved: true });
  step.rows.push({
    id: "changes_decided",
    state: "not_ready",
    text: "2 paragraphs of SI.docx still hold tracked changes. ...",
  });
  const ev = J.submissionEvidence({ ...good(true), step });
  assert.equal(ev.pass, true);
  assert.equal(ev.rows.filter((r) => r.id === "changes_decided" && r.result === "NOTE").length, 1);
});

test("check now: the letter inputs are compared, not the whole record", () => {
  const a = revision();
  assert.equal(J.letterInputsChanged(a, structuredClone(a)), false);
  assert.equal(
    J.letterInputsChanged(a, { ...structuredClone(a), doc_state: { saved: true } }),
    false,
  );
  const drafted = structuredClone(a);
  drafted.comments[0].response.text = "Erledigt.";
  assert.equal(J.letterInputsChanged(a, drafted), true);
  assert.equal(
    J.letterInputsChanged(a, { ...structuredClone(a), profile_overrides: { lang: "de" } }),
    true,
  );
  const declined = structuredClone(a);
  declined.comments[0].declined = { reason: "x" };
  assert.equal(J.letterInputsChanged(a, declined), true);
});

// ---- the figures button ----------------------------------------------------

const DOC = "D:\\run\\review_manuscript.docx";
const HREF = `http://127.0.0.1:47834/figures/?doc=${encodeURIComponent(DOC)}`;
const figuresStep = (extra = {}) => ({
  ok: true,
  href: HREF,
  spy: { supported: true, wrapped: true, calls: [HREF], errors: [] },
  pane_url_before: "http://127.0.0.1:47834/word/index.html",
  pane_url_after: "http://127.0.0.1:47834/word/index.html",
  figures_pages_before: [],
  figures_pages_after: [],
  ...extra,
});
const getLine = (doc, ua = "Mozilla/5.0 (Windows NT 10.0) Edg/140.0") =>
  `[figures] page GET doc=${JSON.stringify(doc)} ref=null ua=${JSON.stringify(ua)}`;

test("figures: the page loads are read from the line figure-http.mjs writes", () => {
  assert.match(read("../daemon/figure-http.mjs"), /`\[figures\] page GET doc=\$\{JSON\.stringify/);
  const gets = J.figuresPageGets(`x\n${getLine(DOC)}\n[figures] page GET doc=nope\n`);
  assert.deepEqual(gets, [{ doc: DOC, ref: null, ua: "Mozilla/5.0 (Windows NT 10.0) Edg/140.0" }]);
  assert.deepEqual(J.figuresPageGets(""), []);
});

test("figures: the API route and the link fallback both pass when the page was served outside the pane", () => {
  const gets = J.figuresPageGets(getLine(DOC));
  const api = J.figuresOpenEvidence({ step: figuresStep(), gets });
  assert.equal(api.pass, true, JSON.stringify(api.rows));
  assert.equal(api.route, "api");
  assert.equal(byId(api.rows)["user agent"].result, "NOTE");
  const fallback = J.figuresOpenEvidence({
    step: figuresStep({ spy: { supported: false, wrapped: false, calls: [], errors: [] } }),
    gets,
  });
  assert.equal(fallback.pass, true);
  assert.equal(fallback.route, "fallback");
});

test("figures: every judged row fails on an input that is wrong for it", () => {
  const gets = J.figuresPageGets(getLine(DOC));
  const failOf = (step, id, g = gets) => {
    const ev = J.figuresOpenEvidence({ step, gets: g });
    assert.equal(ev.pass, false, id);
    assert.equal(byId(ev.rows)[id].result, "FAIL", `${id}: ${JSON.stringify(ev.rows)}`);
  };
  failOf({ ok: false, error: "no link" }, "clicked");
  failOf(figuresStep({ href: "http://127.0.0.1:47834/figures/" }), "clicked");
  // The API is there and watched, but was not called, or with another address.
  failOf(figuresStep({ spy: { supported: true, wrapped: true, calls: [], errors: [] } }), "route");
  failOf(
    figuresStep({ spy: { supported: true, wrapped: true, calls: ["http://x/"], errors: [] } }),
    "route",
  );
  failOf(
    figuresStep({ spy: { supported: true, wrapped: true, calls: [HREF], errors: ["denied"] } }),
    "route",
  );
  // No page load, or one of another document.
  failOf(figuresStep(), "served", []);
  failOf(figuresStep(), "served", J.figuresPageGets(getLine("D:\\other.docx")));
  // The page opened inside Word's WebView2, or the pane itself navigated.
  failOf(figuresStep({ figures_pages_after: [HREF] }), "outside the pane");
  failOf(figuresStep({ pane_url_after: HREF }), "outside the pane");
});

test("figures: the window is closed only when it is surely the run's own", () => {
  const T = "Sidewrite figures and other files";
  const win = (pid, name, title = `${T} - Personal - Microsoft Edge`) => ({
    pid,
    name,
    title,
    has_window: true,
  });
  // No browser ran before; one new window shows the page.
  assert.deepEqual(J.browserCloseDecision({ before: [], after: [win(9, "msedge")], title: T }), {
    action: "close",
    pid: 9,
    name: "msedge",
  });
  const leave = (input, re) => {
    const d = J.browserCloseDecision({ title: T, ...input });
    assert.equal(d.action, "leave");
    assert.match(d.reason, re);
  };
  // The user's browser was running (also a background process with no window).
  leave(
    {
      before: [{ pid: 1, name: "msedge", title: "", has_window: false }],
      after: [win(9, "msedge")],
    },
    /running before the click/,
  );
  // The page opened as a tab of a window that was already there.
  leave(
    { before: [win(1, "msedge", `${T} - Edge`)], after: [win(1, "msedge", `${T} - Edge`)] },
    /no new browser window/,
  );
  // Another title (the user opened something else meanwhile).
  leave({ before: [], after: [win(9, "chrome", "Mail - Chrome")] }, /no new browser window/);
  leave({ before: [], after: [win(9, "msedge"), win(10, "chrome")] }, /2 new browser windows/);
  leave(
    { before: [], after: [win(9, "msedge"), win(10, "msedge", "Other - Edge")] },
    /more than one new msedge/,
  );
  leave({ before: null, after: [] }, /could not be listed/);
  // The title is the one of the page.
  assert.match(read("../taskpane/figures/index.html"), new RegExp(`<title>${T}</title>`));
});

test("figures: browser-windows.ps1 closes only by CloseMainWindow, after checking title and tabs", () => {
  const ps = read("../scripts/e2e/browser-windows.ps1");
  assert.match(ps, /"list" \{/);
  assert.match(ps, /"close" \{/);
  assert.match(ps, /CloseMainWindow\(\)/);
  assert.doesNotMatch(ps, /Stop-Process|\.Kill\(/);
  assert.match(ps, /\$tabs -ne 1/);
  assert.match(ps, /Contains\(\$Title\)/);
});

// ---- --resident --------------------------------------------------------------

test("resident: the log file of the unit and the owners of the ports", () => {
  const unit = [
    "[Service]",
    "ExecStart=/usr/local/bin/node /x/daemon/index.mjs",
    "StandardOutput=append:/home/u/.claude/office-addins/daemon.log",
    "StandardError=append:/home/u/.claude/office-addins/daemon.log",
  ].join("\n");
  assert.equal(J.unitLogFile(unit), "/home/u/.claude/office-addins/daemon.log");
  assert.equal(J.unitLogFile("[Service]\nExecStart=/x\n"), null);
  // The unit the install writes names such a file.
  assert.match(read("../scripts/install/unit.mjs"), /StandardOutput=append:\$\{logFile\}/);
  const ss = [
    'LISTEN 0 511 127.0.0.1:47833 0.0.0.0:* users:(("node",pid=4242,fd=21))',
    'LISTEN 0 511 127.0.0.1:47834 0.0.0.0:* users:(("node",pid=4242,fd=22))',
    'LISTEN 0 511 127.0.0.1:478340 0.0.0.0:* users:(("node",pid=7,fd=1))',
  ].join("\n");
  assert.deepEqual(J.listenerPids(ss, 47833), [4242]);
  assert.deepEqual(J.listenerPids(ss, 47834), [4242]);
  assert.deepEqual(J.listenerPids(ss, 9999), []);
});

test("resident: serving before and after, not restarted; each row fails on its own input", () => {
  const s = (extra = {}) => ({
    active: true,
    main_pid: 4242,
    ws: [4242],
    http: [4242],
    uptime_s: 1000,
    ...extra,
  });
  const ok = J.residentEvidence({ before: s(), after: s({ uptime_s: 1060 }), elapsed_s: 60 });
  assert.equal(ok.pass, true, JSON.stringify(ok.rows));
  const failOf = (input, id) => {
    const ev = J.residentEvidence(input);
    assert.equal(ev.pass, false);
    assert.equal(byId(ev.rows)[id].result, "FAIL", id);
  };
  // Another process holds a port (the run's own daemon, or none of the unit).
  failOf({ before: s({ http: [99] }), after: s({ uptime_s: 1060 }), elapsed_s: 60 }, "serving");
  failOf({ before: s({ active: false }), after: s({ uptime_s: 1060 }), elapsed_s: 60 }, "serving");
  // Restarted during the run: another MainPID, or an uptime that started again.
  failOf(
    {
      before: s(),
      after: s({ main_pid: 5000, ws: [5000], http: [5000], uptime_s: 20 }),
      elapsed_s: 60,
    },
    "not restarted",
  );
  failOf({ before: s(), after: s({ uptime_s: 30 }), elapsed_s: 60 }, "not restarted");
  failOf({ before: s(), after: s({ uptime_s: null }), elapsed_s: 60 }, "not restarted");
  // Stopped by the end of the run.
  failOf(
    { before: s(), after: s({ active: false, ws: [], http: [], uptime_s: null }), elapsed_s: 60 },
    "still serving",
  );
});

// ---- the runner, the driver and the scripts ---------------------------------

test("runner: the new steps are opt-in and ask only for what exists", () => {
  const runner = read("../scripts/e2e/run-acceptance.mjs");
  const driver = read("../scripts/e2e/pane-driver.cjs");
  const com = read("../scripts/e2e/word-com.ps1");
  for (const flag of ["check-now", "figures-button", "figures-timeout", "resident"])
    assert.match(runner, new RegExp(`"?${flag}"?: \\{ type: "(boolean|string)"`), flag);
  // Without the flags stage 3 does nothing new.
  const s3 = runner.slice(
    runner.indexOf("async function mainS3()"),
    runner.indexOf("// Stage 4 (S4-30"),
  );
  assert.match(s3, /opt\["check-now"\] \? await submissionCheck\(\) : null/);
  assert.match(s3, /if \(opt\["figures-button"\]\) await figuresButton\(\);/);
  assert.match(s3, /if \(checkedAt\) await submissionCheck\(checkedAt\);/);
  // The pane steps exist.
  for (const step of ["submission-check", "figures-open", "open-review", "view"])
    assert.match(driver, new RegExp(`\\n  (?:async )?"?${step}"?(?:\\(page|: |,)`), step);
  // The COM actions exist, state with Word's Saved.
  const block = runner.slice(
    runner.indexOf("// ---- real Word:"),
    runner.indexOf("// Stage 2 (procedure"),
  );
  assert.ok(block.length > 1000);
  for (const [, a] of block.matchAll(/\bcom\(\s*"([a-z-]+)"/g))
    assert.ok(com.includes(`  "${a}" {`), `COM action ${a}`);
  assert.match(com, /saved = \[bool\]\$d\.Saved/);
  // --resident neither stops the unit nor starts a daemon of the run.
  const setup = runner.slice(
    runner.indexOf("async function setup()"),
    runner.indexOf("// Stage 2 (procedure"),
  );
  assert.ok(
    setup.indexOf('if (STAGE === "resident") return setupResident();') <
      setup.indexOf("resident = pauseResident();"),
  );
  const resident = block.slice(block.indexOf("async function setupResident()"));
  assert.doesNotMatch(
    resident,
    /startDaemon|stopDaemon|pauseResident|systemctl", \["--user", "(stop|restart|start)"/,
  );
  assert.match(runner, /resident: mainResident/);
  // The isolation check of --resident reads the resident's log.
  assert.match(runner, /STAGE === "resident"\s*\? residentLogText\(\)/);
});

test("README: the new steps, the flag and the assumptions", () => {
  const readme = read("../scripts/e2e/README.md");
  for (const s of [
    "--check-now",
    "--figures-button",
    "--resident",
    "submission-check",
    "figures-open",
    "browser-windows.ps1",
  ])
    assert.ok(readme.includes(s), s);
});
