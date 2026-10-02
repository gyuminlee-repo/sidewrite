// The figures page (taskpane/figures/figures.js) on a fake DOM: the real
// script, run as the page at an address would run it, with the daemon's
// answers given by the test. No browser.
import { readFile } from "node:fs/promises";
import vm from "node:vm";

// What figures.js needs of a page. append and replaceChildren write the word
// "null" for a null, as a browser does.
export class PageEl {
  constructor(tag, id = null) {
    this.tagName = tag.toUpperCase();
    this.id = id;
    this.children = [];
    this.listeners = {};
    this.attributes = {};
    this.className = "";
    this.textContent = "";
    this.hidden = false;
    this.value = "";
    this.title = "";
    this.disabled = false;
    this.style = {};
    this.scrolled = 0;
    this.previousElementSibling = { hidden: false };
  }
  addEventListener(ev, fn) {
    (this.listeners[ev] ??= []).push(fn);
  }
  append(...kids) {
    for (const k of kids)
      this.children.push(
        k instanceof PageEl ? k : Object.assign(new PageEl("#text"), { textContent: String(k) }),
      );
  }
  replaceChildren(...kids) {
    this.children = [];
    this.append(...kids);
  }
  setAttribute(k, v) {
    this.attributes[k] = v;
  }
  scrollIntoView() {
    this.scrolled += 1;
  }
}

export const findAll = (node, pred, out = []) => {
  if (pred(node)) out.push(node);
  for (const k of node.children) findAll(k, pred, out);
  return out;
};
export const textOf = (n) => [n.textContent, ...n.children.map(textOf)].join("");

const tick = (ms) => new Promise((r) => setTimeout(r, ms));

// search: the query of the address ("?doc=...&ref=..."). manuscripts: what
// /figures/api/manuscripts answers. comments: the comments of the state.
// Returns the page after its first drawing: $ (by id), asked (every address
// it fetched), cards (the comment cards) and card(ref).
export async function figuresPage(search, { manuscripts, comments }) {
  const byId = new Map();
  const document = {
    getElementById(id) {
      if (!byId.has(id)) byId.set(id, new PageEl("div", id));
      return byId.get(id);
    },
    createElement: (tag) => new PageEl(tag),
  };
  const asked = [];
  const fetch = async (url) => {
    asked.push(url);
    const reply = (body) => ({ ok: true, json: async () => body, text: async () => body });
    if (url === "/bridge-token") return reply("token");
    const name = new URL(url, "http://127.0.0.1:1").pathname.split("/").pop();
    if (name === "manuscripts") return reply({ ok: true, manuscripts });
    if (name === "files") return reply({ ok: true, folder: "figures", files: [] });
    if (name === "state")
      return reply({
        ok: true,
        name: "paper.docx",
        folder: "figures",
        renderer: { available: true, reason: null },
        comments,
      });
    return { ok: false, status: 404, json: async () => ({ ok: false, error: "not in this test" }) };
  };
  const src = await readFile(new URL("../../taskpane/figures/figures.js", import.meta.url), "utf8");
  vm.runInNewContext(src, { document, fetch, location: { search }, URLSearchParams, URL });
  await tick(20);
  const $ = (id) => document.getElementById(id);
  const cards = findAll(
    $("comments"),
    (n) => n.tagName === "SECTION" && n.attributes["data-ref"] !== undefined,
  );
  return { $, asked, cards, card: (ref) => cards.find((c) => c.attributes["data-ref"] === ref) };
}
