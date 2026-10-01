// Pure parts of word-math-probe.mjs: a reading of word/document.xml that
// shows how equations and tracked changes stand in each other. No Word, no
// files.

// A tree of { tag, attrs, children, text, parent } from XML text. Enough for
// a document.xml as Word writes it: no DTD, no CDATA; comments and
// processing instructions are passed over; entities stay as written.
export function parseXml(xml) {
  const root = { tag: "#root", attrs: "", children: [], text: "", parent: null };
  const stack = [root];
  const token =
    /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<(\/?)([^\s<>/]+)((?:"[^"]*"|'[^']*'|[^<>"'])*?)(\/?)>|([^<]+)/g;
  for (const m of String(xml).matchAll(token)) {
    const top = stack[stack.length - 1];
    if (m[2] === undefined) {
      if (m[5] !== undefined) top.text += m[5];
    } else if (m[1]) {
      if (stack.length > 1) stack.pop();
    } else {
      const node = { tag: m[2], attrs: m[3].trim(), children: [], text: "", parent: top };
      top.children.push(node);
      if (!m[4]) stack.push(node);
    }
  }
  return root;
}

// The marks of a tracked change, as wrappers or as flags in properties.
const MARKS = ["w:ins", "w:del", "w:moveFrom", "w:moveTo"];
const TEXT = ["w:t", "w:delText", "m:t", "w:instrText", "w:delInstrText"];

const child = (node, ...path) =>
  path.reduce((n, tag) => n && n.children.find((c) => c.tag === tag), node);
const below = (node, out = []) => {
  for (const c of node.children) {
    out.push(c);
    below(c, out);
  }
  return out;
};
const textOf = (node) =>
  below(node)
    .filter((n) => TEXT.includes(n.tag))
    .map((n) => n.text)
    .join("");
// The change marks a properties element holds (also w:rPrChange and the like).
const flags = (pr) =>
  (pr?.children ?? [])
    .filter((c) => MARKS.includes(c.tag) || c.tag.endsWith("Change"))
    .map((c) => c.tag);
const properties = (node) =>
  `${node.tag}[${node.children.map((c) => (c.tag === "w:rPr" ? `w:rPr[${flags(c).join(", ")}]` : c.tag)).join(", ")}]`;

function outline(node, depth, out) {
  const pad = "  ".repeat(depth);
  if (node.tag === "w:pPr") return;
  if (node.tag === "w:r" || node.tag === "m:r") {
    const holds = node.children.map((c) => c.tag).filter((t) => t !== "w:rPr" && t !== "m:rPr");
    const rpr = child(node, "w:rPr");
    out.push(
      [
        `${pad}${node.tag} ${JSON.stringify(textOf(node))}`,
        `[${[...new Set(holds)].join(", ")}]`,
        ...(child(node, "m:rPr") ? ["m:rPr"] : []),
        ...(rpr ? [`w:rPr[${flags(rpr).join(", ")}]`] : []),
      ].join(" "),
    );
    return;
  }
  if (node.tag === "m:ctrlPr" || node.tag === "w:rPr") {
    out.push(pad + properties(node));
    return;
  }
  out.push(pad + node.tag);
  for (const c of node.children) outline(c, depth + 1, out);
}

// The paragraphs of a document.xml that hold an equation or a change mark,
// as lines: the paragraph with the state of its paragraph mark and its text,
// then what stands in what, one element a line. A run is one line with its
// text, what it holds, and the change marks in its w:rPr; m:ctrlPr is one
// line with what it holds.
export function mathSummary(xml) {
  const body = child(parseXml(xml), "w:document", "w:body");
  if (!body) return ["no w:body in this XML"];
  const out = [];
  below(body)
    .filter((n) => n.tag === "w:p")
    .forEach((p, i) => {
      if (!below(p).some((n) => n.tag.startsWith("m:") || MARKS.includes(n.tag))) return;
      const mark = flags(child(p, "w:pPr", "w:rPr"));
      out.push(
        `paragraph ${i + 1}  mark: ${mark.join(", ") || "-"}  text: ${JSON.stringify(textOf(p))}`,
      );
      for (const c of p.children) outline(c, 1, out);
    });
  return out;
}

// Where change marks stand, without the text: for every mark its two
// ancestors and what it holds, with how often that occurs. The answer to
// "does w:ins wrap m:oMath, or stand inside it, or sit in a w:rPr".
export function markRelations(xml) {
  const counts = new Map();
  for (const n of below(parseXml(xml))) {
    if (!MARKS.includes(n.tag)) continue;
    const up = [n.parent?.parent?.tag, n.parent?.tag].filter((t) => t && t !== "#root");
    const holds = [...new Set(n.children.map((c) => c.tag))];
    const key = `${[...up, n.tag].join(" > ")}${holds.length ? ` > {${holds.join(", ")}}` : " (a flag)"}`;
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  return [...counts]
    .sort(([a], [b]) => (a < b ? -1 : 1))
    .map(([key, n]) => `${String(n).padStart(3)} x ${key}`);
}
