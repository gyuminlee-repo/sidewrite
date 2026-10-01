// scripts/e2e/pane-driver.cjs waits on the task pane's DOM from outside
// (page.waitForFunction) and reads it (page.evaluate). Right after a Word
// window opens, the pane's document can be there before its elements are:
// a condition that read `getElementById(...).textContent` then threw and the
// step failed ("Cannot read properties of null"), where it should have kept
// waiting. Every such function is taken from the source here and run on a
// document with no elements: a wait condition is false, a read does not
// throw. The driver itself needs Word and is not loaded.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const src = await readFile(new URL("../scripts/e2e/pane-driver.cjs", import.meta.url), "utf8");

// The first argument of each `<name>(` call in the source, as text.
function firstArgs(name) {
  const out = [];
  const open = `${name}(`;
  for (let at = src.indexOf(open); at !== -1; at = src.indexOf(open, at + 1)) {
    let i = at + open.length;
    let depth = 0;
    let prev = "(";
    for (; i < src.length; i++) {
      const c = src[i];
      if (c === '"' || c === "'" || c === "`") {
        for (i++; src[i] !== c; i++) if (src[i] === "\\") i++;
      } else if (c === "/" && src[i + 1] === "/") {
        while (src[i] !== "\n") i++;
        continue;
      } else if (c === "/" && "(,=:[!&|?{};".includes(prev)) {
        // A regex literal: to its closing slash, classes and escapes skipped.
        let cls = false;
        for (i++; cls || src[i] !== "/"; i++) {
          if (src[i] === "\\") i++;
          else if (src[i] === "[") cls = true;
          else if (src[i] === "]") cls = false;
        }
      } else if ("([{".includes(c)) depth++;
      else if (")]}".includes(c)) {
        if (depth === 0) break;
        depth--;
      } else if (c === "," && depth === 0) break;
      if (!/\s/.test(src[i])) prev = src[i];
    }
    out.push(src.slice(at + open.length, i).trim());
  }
  return out;
}

// A document of the task pane: elements by id, as plain objects.
function fakeDocument(byId = {}, all = {}) {
  return {
    getElementById: (id) => byId[id] ?? null,
    querySelector: (sel) => (all[sel] ?? [])[0] ?? null,
    querySelectorAll: (sel) => all[sel] ?? [],
  };
}
const el = (textContent = "", more = {}) => ({
  textContent,
  hidden: false,
  disabled: false,
  title: "",
  classList: { contains: () => false },
  getAttribute: () => null,
  click() {},
  ...more,
});
const windowStub = { dispatchEvent() {} };
// The function of a source text, bound to a document.
const bind = (text, document) =>
  new Function("document", "window", `return (${text});`)(document, windowStub);
// An argument of the shape the function takes.
const argFor = (text) => (/^\(\s*\[/.test(text) ? ["turn", { id: "claude" }] : 3);

const waits = firstArgs("waitForFunction");
const reads = firstArgs("page.evaluate").filter((t) => !t.includes("import("));

test("the source scan finds the wait conditions and the reads of the driver", () => {
  assert.ok(waits.length >= 15, `wait conditions: ${waits.length}`);
  assert.ok(reads.length >= 5, `reads: ${reads.length}`);
  for (const t of [...waits, ...reads]) assert.match(t, /=>/, t.slice(0, 60));
  // No read of an element that may not be there yet.
  assert.deepEqual(src.match(/getElementById\("[^"]+"\)\.(?!\s*$)/g) ?? [], []);
  console.log(`pane-driver: ${waits.length} wait conditions, ${reads.length} reads`);
});

test("no element yet: every wait condition is false and none throws", () => {
  for (const text of waits) {
    let got;
    assert.doesNotThrow(
      () => (got = bind(text, fakeDocument())(argFor(text))),
      `threw on an empty document: ${text.slice(0, 120)}`,
    );
    assert.ok(!got, `true on an empty document: ${text.slice(0, 120)}`);
  }
});

test("no element yet: every read gives an answer and none throws", async () => {
  for (const text of reads) {
    await assert.doesNotReject(
      async () => bind(text, fakeDocument())(argFor(text)),
      `threw on an empty document: ${text.slice(0, 120)}`,
    );
  }
});

// With the elements there the conditions decide as before.
const only = (part) => {
  const found = waits.filter((t) => t.includes(part));
  assert.ok(found.length >= 1, `no wait condition with ${part}`);
  return found;
};

test("the elements are there: the conditions decide on them", () => {
  // The review file line (open-review, and the snapshot step that failed).
  for (const text of only('startsWith("Review file:")')) {
    const on = (t) => !!bind(text, fakeDocument({ "review-file": el(t) }))();
    assert.deepEqual([on(""), on("Review file: /m/paper.revision.json")], [false, true]);
  }
  // The message line of a step: empty waits, a text ends the wait.
  for (const text of only("pattern")) {
    const on = (t, pattern, error = false) =>
      !!bind(
        text,
        fakeDocument({ "review-message": el(t, { classList: { contains: () => error } }) }),
      )(pattern);
    assert.deepEqual(
      [on("", null), on("Saved.", null), on("Waiting...", "^Saved"), on("Saved.", "^Saved")],
      [false, true, false, true],
    );
    assert.equal(on("Failed", "^Saved", true), true, "an error line ends the wait");
  }
  // The providers list is drawn, or its message says why not.
  for (const text of only("#providers-list .provider-row")) {
    const on = (rows, msg) =>
      !!bind(
        text,
        fakeDocument(
          { "providers-message": el(msg) },
          { "#providers-list .provider-row": Array(rows).fill(el()) },
        ),
      )();
    assert.deepEqual([on(0, ""), on(2, ""), on(0, "Daemon not ready")], [false, true, true]);
  }
  // The prompt preview shows.
  for (const text of only('"review-preview"')) {
    const on = (hidden) => !!bind(text, fakeDocument({ "review-preview": el("", { hidden }) }))();
    assert.deepEqual([on(true), on(false)], [false, true]);
  }
  // A turn started: the chat input is disabled.
  for (const text of waits.filter((t) => /getElementById\("input"\)\?\.disabled/.test(t))) {
    const on = (disabled) => !!bind(text, fakeDocument({ input: el("", { disabled }) }))();
    assert.deepEqual([on(false), on(true)], [false, true]);
  }
  // Propose all runs (Stop shows) or ended at once with a message.
  for (const text of only('"review-propose-stop"')) {
    const on = (hidden, msg) =>
      !!bind(
        text,
        fakeDocument({ "review-propose-stop": el("", { hidden }), "review-message": el(msg) }),
      )();
    assert.deepEqual(
      [on(true, ""), on(false, ""), on(true, "No open comment")],
      [false, true, true],
    );
  }
  // The ledger check ended: no item "not checked yet", or a message.
  for (const text of only("not checked yet")) {
    const on = (decisions, msg) =>
      !!bind(
        text,
        fakeDocument({ "review-decisions": el(decisions), "review-message": el(msg) }),
      )();
    assert.deepEqual(
      [on("L1 not checked yet", ""), on("L1 ok", ""), on("L1 not checked yet", "failed")],
      [false, true, true],
    );
  }
  // The model menu is filled, or says why not.
  for (const text of only('"composer-model"')) {
    const on = (title) => !!bind(text, fakeDocument({ "composer-model": el("", { title }) }))();
    assert.deepEqual(
      [on(""), on("Models your Claude account can use"), on("Model list unavailable (x)")],
      [false, true, true],
    );
  }
});
