// The figures routes on an HTTP server of their own, with the real bridge and
// a WebSocket client as the Word pane: who may ask, and for which document.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { fakeRunPy, figureWorld, sha256, listing } from "./helpers/figure-world.mjs";

test("the API answers 401 without the bridge token or with another one", async () => {
  const w = await figureWorld({ runPy: fakeRunPy() });
  try {
    for (const name of ["manuscripts", "state", "files", "render"]) {
      const r = await w.call(name, { token: null, query: { doc: w.doc } });
      assert.equal(r.status, 401, name);
      assert.equal(r.json.ok, false);
    }
    assert.equal((await w.call("manuscripts", { token: "figure-test-tokem" })).status, 401);
    const post = await w.propose({});
    assert.equal(post.status, 200);
    const noToken = await w.call("adopt", {
      token: null,
      body: { doc: w.doc, ref: "R1.1", id: post.json.id },
    });
    assert.equal(noToken.status, 401);
    assert.equal((await w.rev()).comments[0].figure_changes[0].state, "proposed");
    // With the token the same request is answered.
    assert.equal((await w.call("manuscripts")).status, 200);
  } finally {
    await w.cleanup();
  }
});

test("a request from another origin or to another host name is refused with 403, token or not", async () => {
  const w = await figureWorld({ runPy: fakeRunPy() });
  try {
    const foreign = await w.call("manuscripts", { headers: { Origin: "https://example.org" } });
    assert.equal(foreign.status, 403);
    const nullOrigin = await w.call("manuscripts", { headers: { Origin: "null" } });
    assert.equal(nullOrigin.status, 403);
    // The name of this server as another site would resolve it (DNS rebinding).
    const port = new URL(w.origin).port;
    assert.equal((await w.call("manuscripts", { host: `example.org:${port}` })).status, 403);
    assert.equal((await w.call("manuscripts", { host: `localhost:${port}` })).status, 403);
    assert.equal((await w.call("manuscripts", { host: "127.0.0.1:1" })).status, 403);
    // The page itself is refused the same way.
    assert.equal((await w.call("", { path: "/figures/", host: "example.org" })).status, 403);
    const post = await w.propose({});
    const adopt = await w.call("adopt", {
      headers: { Origin: "https://example.org" },
      body: { doc: w.doc, ref: "R1.1", id: post.json.id },
    });
    assert.equal(adopt.status, 403);
    assert.equal((await w.rev()).comments[0].figure_changes[0].state, "proposed");
    // Its own origin passes.
    assert.equal((await w.call("manuscripts", { headers: { Origin: w.origin } })).status, 200);
  } finally {
    await w.cleanup();
  }
});

test("only the document of a connected Word pane may be named", async () => {
  const w = await figureWorld({ runPy: fakeRunPy(), pane: false });
  try {
    assert.deepEqual((await w.call("manuscripts")).json.manuscripts, []);
    // The folder and its revision.json are there; no pane has the document.
    const closed = await w.call("state", { query: { doc: w.doc } });
    assert.equal(closed.status, 403);
    assert.match(closed.json.error, /not open in a Word pane/);
    assert.equal((await w.propose({})).status, 403);
    assert.equal((await w.rev()).comments[0].figure_changes.length, 0);

    // An Excel pane with the same path does not count.
    const excel = await w.openPane(w.doc, "excel");
    assert.equal((await w.call("state", { query: { doc: w.doc } })).status, 403);
    await excel.close();

    const pane = await w.openPane();
    const list = (await w.call("manuscripts")).json.manuscripts;
    assert.deepEqual(list, [{ doc: w.doc, path: w.doc }]);
    assert.equal((await w.call("state", { query: { doc: w.doc } })).status, 200);
    // Another file in the same folder is not that document.
    const other = await w.call("state", { query: { doc: join(w.dir, "other.docx") } });
    assert.equal(other.status, 403);
    assert.equal((await w.call("state", {})).status, 403);

    await pane.close();
    assert.equal((await w.call("state", { query: { doc: w.doc } })).status, 403);
  } finally {
    await w.cleanup();
  }
});

test("propose, state, update, reject and reopen over HTTP", async () => {
  const w = await figureWorld({ runPy: fakeRunPy() });
  try {
    const p = await w.propose({});
    assert.equal(p.status, 200);
    assert.deepEqual(p.json, { ok: true, id: "R1.1-f1", problems: [] });
    const body = { doc: w.doc, ref: "R1.1", id: "R1.1-f1" };
    const up = await w.call("update", {
      body: { ...body, patch: { summary: "A control curve was added." } },
    });
    assert.equal(up.status, 200);
    assert.equal((await w.call("reject", { body })).status, 200);
    const again = await w.call("reject", { body });
    assert.equal(again.status, 409);
    assert.match(again.json.error, /is rejected; it cannot become rejected/);
    assert.equal((await w.call("reopen", { body })).status, 200);
    const state = (await w.call("state", { query: { doc: w.doc } })).json;
    assert.equal(state.name, "paper.docx");
    const f = state.comments[0].figures[0];
    assert.equal(f.state, "proposed");
    assert.equal(f.summary, "A control curve was added.");
    const files = (await w.call("files", { query: { doc: w.doc } })).json;
    assert.deepEqual(files.files.map((x) => x.path).sort(), [
      "figures/fig2_after.svg",
      "figures/fig2_before.svg",
    ]);
    // What is no request of the API.
    assert.equal((await w.call("nothing")).status, 404);
    assert.equal((await w.call("state", { body: { doc: w.doc } })).status, 404);
    const bad = await w.call("propose", { body: null, headers: { "Content-Type": "text/plain" } });
    assert.equal(bad.status, 404, "a GET of a POST route");
  } finally {
    await w.cleanup();
  }
});

test("a render request cannot name a path: the picture is found by the recorded change", async () => {
  const w = await figureWorld({ runPy: fakeRunPy() });
  try {
    const secret = join(w.dir, "secret.png");
    await writeFile(secret, "not for the page");
    const { id } = (await w.propose({})).json;
    const q = { doc: w.doc, ref: "R1.1", id, side: "after" };
    const good = await w.call("render", { query: q });
    assert.equal(good.status, 200);
    assert.equal(good.headers["content-type"], "image/png");
    const f = (await w.rev()).comments[0].figure_changes[0];
    const cached = await readFile(join(w.dir, f.renders.after.path));
    assert.deepEqual(good.body, cached);

    // Extra names for a file are not read.
    for (const extra of [{ path: secret }, { file: secret }, { png: secret }, { out: secret }]) {
      const r = await w.call("render", { query: { ...q, ...extra } });
      assert.equal(r.status, 200);
      assert.deepEqual(r.body, cached);
    }
    // A path in place of an id, a side or a ref finds no record.
    for (const bad of [
      { id: secret },
      { id: "../../secret.png" },
      { side: "../../secret.png" },
      { side: secret },
      { ref: "../.." },
      { id: "" },
    ]) {
      const r = await w.call("render", { query: { ...q, ...bad } });
      assert.ok([400, 404].includes(r.status), `${JSON.stringify(bad)} gave ${r.status}`);
      assert.equal(r.body.includes("not for the page"), false);
    }
    // Nor does the page route read files by name.
    for (const path of [
      "/figures/../secret.png",
      "/figures/..%2f..%2fpackage.json",
      "/figures/x",
    ]) {
      const r = await w.call("", { path });
      assert.equal(r.status, 404, path);
    }
    assert.equal(await readFile(secret, "utf8"), "not for the page");
  } finally {
    await w.cleanup();
  }
});

test("the page and its two files are served; anything else on the server is not a figures route", async () => {
  const w = await figureWorld({ runPy: fakeRunPy() });
  try {
    const page = await w.call("", { path: "/figures/", token: null });
    assert.equal(page.status, 200);
    assert.match(page.headers["content-type"], /text\/html/);
    const html = page.body.toString("utf8");
    assert.match(html, /Open the manuscript in Word with the Sidewrite pane, then reload\./);
    assert.match(html, /<script src="\/figures\/figures\.js"><\/script>/);
    // No asset from another server, and no em dash in what the reader sees.
    const js = await w.call("", { path: "/figures/figures.js", token: null });
    const css = await w.call("", { path: "/figures/figures.css", token: null });
    assert.equal(js.status, 200);
    assert.equal(css.status, 200);
    for (const text of [html, js.body.toString("utf8"), css.body.toString("utf8")]) {
      assert.equal(/https?:\/\/(?!127\.0\.0\.1)/.test(text), false, "no outside address");
      assert.equal(text.includes("—"), false, "no em dash");
    }
    const redirect = await w.call("", { path: "/figures", token: null });
    assert.equal(redirect.status, 302);
    assert.equal(redirect.headers.location, "/figures/");
    // Not a figures route: the handler leaves it to the server.
    const other = await w.call("", { path: "/figuresque" });
    assert.equal(other.status, 404);
    assert.equal(other.body.toString(), "not a figures route");
    assert.equal(
      (await w.call("", { path: "/word/index.html" })).body.toString(),
      "not a figures route",
    );
  } finally {
    await w.cleanup();
  }
});

test("Adopt over HTTP writes only under the figures folder of that manuscript", async () => {
  const w = await figureWorld({ runPy: fakeRunPy() });
  try {
    const before = await listing(w.dir);
    const { id } = (await w.propose({})).json;
    const r = await w.call("adopt", { body: { doc: w.doc, ref: "R1.1", id } });
    assert.equal(r.status, 200);
    assert.equal(await sha256(join(w.dir, r.json.adopted.path)), r.json.adopted.sha256);
    const added = (await listing(w.dir)).filter((p) => !before.includes(p));
    assert.deepEqual(
      added.filter(
        (p) => !p.startsWith("figures/.renders/") && !p.startsWith("figures/submission/"),
      ),
      [],
    );
  } finally {
    await w.cleanup();
  }
});
