import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  findStyles,
  styleUrl,
  parseStyleInfo,
  fetchStyle,
  CSL_NOTICE,
  CSL_OFF,
} from "../daemon/csl-style.mjs";

const fixture = (name) => readFile(new URL(`./fixtures/${name}`, import.meta.url));
const RAW = "https://raw.githubusercontent.com/citation-style-language/styles/master";
const CHILD = "journal-of-invented-results";
const PARENT = "invented-parent-style";

// A stand-in for the styles repository: path -> bytes.
async function fakeRepo(extra = {}) {
  const files = {
    [`/dependent/${CHILD}.csl`]: await fixture("csl_dependent.csl"),
    [`/${PARENT}.csl`]: await fixture("csl_independent.csl"),
    ...extra,
  };
  const calls = [];
  const fetchFn = async (url) => {
    calls.push(url);
    assert.ok(url.startsWith(`${RAW}/`), url);
    const body = files[url.slice(RAW.length)];
    if (body instanceof Error) throw body;
    if (!body) return { ok: false, status: 404, arrayBuffer: async () => new ArrayBuffer(0) };
    return {
      ok: true,
      status: 200,
      arrayBuffer: async () =>
        body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength),
    };
  };
  return { fetchFn, calls, files };
}
const cache = () => mkdtemp(join(tmpdir(), "csl-"));

test("findStyles: an exact title first, then titles that contain the words, five at most", () => {
  const index = [
    {
      title: "Bioresource Technology Reports",
      name: "bioresource-technology-reports",
      dependent: 1,
    },
    { title: "Bioresource Technology", name: "bioresource-technology", dependent: 1 },
    { title: "Nature", name: "nature", dependent: 0 },
    ...Array.from({ length: 6 }, (_, i) => ({
      title: `Bioresource Technology Letters ${i}`,
      name: `btl-${i}`,
      dependent: 1,
    })),
    { titleShort: "BT", title: "Something Else", name: "something-else", dependent: 0 },
  ];
  const hits = findStyles("  bioresource   technology ", index);
  assert.equal(hits.length, 5);
  assert.equal(hits[0].name, "bioresource-technology");
  assert.equal(hits[1].name, "bioresource-technology-reports");
  assert.deepEqual(
    findStyles("bt", index).map((s) => s.name),
    ["something-else"],
  );
  assert.deepEqual(findStyles("no such journal", index), []);
  assert.deepEqual(findStyles("", index), []);
  assert.deepEqual(findStyles("nature", null), []);
});

test("styleUrl: dependent styles live in dependent/", () => {
  assert.equal(styleUrl({ name: "x", dependent: 1 }), `${RAW}/dependent/x.csl`);
  assert.equal(styleUrl({ name: "x", dependent: 0 }), `${RAW}/x.csl`);
  assert.equal(styleUrl({ name: "x" }), `${RAW}/x.csl`);
});

test("parseStyleInfo reads title, id, parent, format, rights, authors and contributors", async () => {
  const parent = parseStyleInfo((await fixture("csl_independent.csl")).toString("utf8"));
  assert.deepEqual(parent, {
    title: "Invented Parent Style",
    id: "http://www.zotero.org/styles/invented-parent-style",
    parent: null,
    format: "author-date",
    rights: "This work is licensed under a Creative Commons Attribution-ShareAlike 3.0 License",
    license: "http://creativecommons.org/licenses/by-sa/3.0/",
    authors: ["Test Author"],
    contributors: ["Second Tester", "Third & Fourth Tester"],
  });
  const child = parseStyleInfo((await fixture("csl_dependent.csl")).toString("utf8"));
  assert.equal(child.title, "Journal of Invented Results");
  assert.equal(child.parent, PARENT);
  assert.deepEqual([child.authors, child.contributors], [[], []]);
  assert.equal(parseStyleInfo("<html><body>Not found</body></html>"), null);
  assert.equal(parseStyleInfo(""), null);
});

test("a dependent style is fetched with its parent, once each, and both are kept as received", async () => {
  const { fetchFn, calls } = await fakeRepo();
  const cacheDir = await cache();
  const r = await fetchStyle({ slug: CHILD, enabled: true, fetchFn, cacheDir });
  assert.equal(r.ok, true);
  // Not known to be dependent: the top folder is asked first, then dependent/.
  assert.deepEqual(calls, [
    `${RAW}/${CHILD}.csl`,
    `${RAW}/dependent/${CHILD}.csl`,
    `${RAW}/${PARENT}.csl`,
  ]);
  assert.equal(r.info.title, "Journal of Invented Results");
  assert.equal(r.info.slug, CHILD);
  assert.equal(r.info.parent, PARENT);
  assert.equal(r.parent.title, "Invented Parent Style");
  assert.deepEqual(r.parent.contributors, ["Second Tester", "Third & Fourth Tester"]);
  assert.equal(r.notice, CSL_NOTICE);
  assert.deepEqual((await readdir(cacheDir)).sort(), [`${PARENT}.csl`, `${CHILD}.csl`].sort());
  assert.ok(
    (await readFile(join(cacheDir, `${CHILD}.csl`))).equals(await fixture("csl_dependent.csl")),
  );
  assert.ok(
    (await readFile(join(cacheDir, `${PARENT}.csl`))).equals(await fixture("csl_independent.csl")),
  );
  assert.deepEqual(r.files, [join(cacheDir, `${CHILD}.csl`), join(cacheDir, `${PARENT}.csl`)]);
});

test("a style known to be dependent is asked for in dependent/ directly; an independent one has no parent", async () => {
  const a = await fakeRepo();
  await fetchStyle({
    slug: CHILD,
    dependent: true,
    enabled: true,
    fetchFn: a.fetchFn,
    cacheDir: await cache(),
  });
  assert.deepEqual(a.calls, [`${RAW}/dependent/${CHILD}.csl`, `${RAW}/${PARENT}.csl`]);
  const b = await fakeRepo();
  const r = await fetchStyle({
    slug: PARENT,
    enabled: true,
    fetchFn: b.fetchFn,
    cacheDir: await cache(),
  });
  assert.deepEqual(b.calls, [`${RAW}/${PARENT}.csl`]);
  assert.deepEqual([r.ok, r.parent, r.info.format], [true, null, "author-date"]);
});

test("a cached style is read from the cache folder: no request", async () => {
  const cacheDir = await cache();
  const first = await fakeRepo();
  await fetchStyle({ slug: CHILD, enabled: true, fetchFn: first.fetchFn, cacheDir });
  const second = await fakeRepo();
  const r = await fetchStyle({ slug: CHILD, enabled: true, fetchFn: second.fetchFn, cacheDir });
  assert.equal(second.calls.length, 0);
  assert.deepEqual(
    [r.ok, r.info.title, r.parent.title],
    [true, "Journal of Invented Results", "Invented Parent Style"],
  );
});

test("off: no request and nothing written", async () => {
  const { fetchFn, calls } = await fakeRepo();
  const cacheDir = await cache();
  for (const enabled of [false, undefined, "true"]) {
    assert.deepEqual(await fetchStyle({ slug: CHILD, enabled, fetchFn, cacheDir }), {
      ok: false,
      error: CSL_OFF,
      notice: CSL_NOTICE,
    });
  }
  assert.equal(calls.length, 0);
  assert.deepEqual(await readdir(cacheDir), []);
});

test("a name that is not a style slug is refused before any request", async () => {
  const { fetchFn, calls } = await fakeRepo();
  const cacheDir = await cache();
  for (const slug of [
    "../secrets",
    "a/b",
    "Nature",
    "x.csl",
    "",
    null,
    "a b",
    "-x",
    "a".repeat(200),
  ]) {
    const r = await fetchStyle({ slug, enabled: true, fetchFn, cacheDir });
    assert.equal(r.ok, false, String(slug));
  }
  assert.equal(calls.length, 0);
  assert.deepEqual(await readdir(cacheDir), []);
});

test("an unknown style, a page that is not a style and a network failure are errors and cache nothing", async () => {
  const cacheDir = await cache();
  const none = await fakeRepo();
  const r = await fetchStyle({
    slug: "no-such-style",
    enabled: true,
    fetchFn: none.fetchFn,
    cacheDir,
  });
  assert.deepEqual([r.ok, /no style named no-such-style/.test(r.error)], [false, true]);
  assert.equal(none.calls.length, 2);
  const html = await fakeRepo({ "/odd.csl": Buffer.from("<html><body>hello</body></html>") });
  const odd = await fetchStyle({ slug: "odd", enabled: true, fetchFn: html.fetchFn, cacheDir });
  assert.deepEqual([odd.ok, /not a CSL style/.test(odd.error)], [false, true]);
  const down = await fakeRepo({ "/down.csl": new Error("ETIMEDOUT") });
  const d = await fetchStyle({ slug: "down", enabled: true, fetchFn: down.fetchFn, cacheDir });
  assert.deepEqual([d.ok, /ETIMEDOUT/.test(d.error)], [false, true]);
  assert.deepEqual(await readdir(cacheDir), []);
});

test("a missing parent is an error; the parent's own parent is never followed", async () => {
  const orphan = await fakeRepo({ [`/${PARENT}.csl`]: undefined });
  const r = await fetchStyle({
    slug: CHILD,
    dependent: true,
    enabled: true,
    fetchFn: orphan.fetchFn,
    cacheDir: await cache(),
  });
  assert.deepEqual([r.ok, /parent style invented-parent-style/.test(r.error)], [false, true]);
  // A parent that itself names a parent: one step only.
  const loop = await fakeRepo({ [`/${PARENT}.csl`]: await fixture("csl_dependent.csl") });
  const l = await fetchStyle({
    slug: CHILD,
    dependent: true,
    enabled: true,
    fetchFn: loop.fetchFn,
    cacheDir: await cache(),
  });
  assert.equal(loop.calls.length, 2);
  assert.equal(l.ok, true);
});

test("a cache file that is not a style is fetched again and replaced", async () => {
  const cacheDir = await cache();
  await writeFile(join(cacheDir, `${PARENT}.csl`), "garbage");
  const { fetchFn, calls } = await fakeRepo();
  const r = await fetchStyle({ slug: PARENT, enabled: true, fetchFn, cacheDir });
  assert.deepEqual([r.ok, calls.length], [true, 1]);
  assert.ok(
    (await readFile(join(cacheDir, `${PARENT}.csl`))).equals(await fixture("csl_independent.csl")),
  );
});

test("the notice names the project, the licence and says styles are not redistributed", () => {
  for (const words of [
    "CC BY-SA 3.0",
    "https://citationstyles.org/",
    "https://creativecommons.org/licenses/by-sa/3.0/",
    "not redistributed",
    "kept unchanged",
    "github",
  ]) {
    assert.ok(CSL_NOTICE.includes(words), words);
  }
});
