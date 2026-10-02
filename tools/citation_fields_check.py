#!/usr/bin/env python3
"""Are the fields of a .docx still whole after it was edited?

  citation_fields_check.py --before A.docx --after B.docx [--expect-removed N] [--detail]

Reads both files (no Word) and compares the fields of every story (body,
footnotes, endnotes, comments, headers, footers): complex fields (w:fldChar
begin / separate / end with w:instrText, and w:delInstrText inside a tracked
deletion), w:fldSimple and w:hyperlink. Citations of Zotero
(ADDIN ZOTERO_ITEM, ADDIN ZOTERO_BIBL) and EndNote (ADDIN EN.CITE,
ADDIN EN.CITE.DATA, ADDIN EN.REFLIST) are such fields.

It holds when
  - every story of B is balanced: no end or separate without a begin, no
    second separate, no begin left open at the end of the story. A field may
    run over several paragraphs (a bibliography does); the depth at the start
    and the end of each paragraph is reported, not judged;
  - no field of B is split by a tracked change: its begin, separate and end
    markers and its instruction lie all outside a tracked change, or all in
    deletions, or all in insertions. A begin in a w:del with its end outside
    leaves half a field when the change is accepted or rejected;
  - the instructions of B are those of A as a multiset, compared as the exact
    text of the concatenated instruction; with --expect-removed N, B lacks
    exactly N of them and has no other;
  - the w:fldData of B are those of A (white space apart); with
    --expect-removed, B has none that A lacks;
  - the ZOTERO_* properties of docProps/custom.xml and the EN.* and ZOTERO*
    document variables of word/settings.xml are the same.
A itself must be balanced too: a broken A makes the comparison mean nothing.

Not judged, only listed (`text_changed`): a live Zotero or EndNote citation
whose visible text is no longer the text its instruction records
(plainCitation, DisplayText). The field is whole; the plugin will notice the
difference at its next refresh.

Prints one line per failure, then one JSON line. Exit code 0 when everything
holds, 1 when not, 2 when a file cannot be read. Standard library only.
"""
import argparse
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from collections import Counter
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
VT = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"
STORY = re.compile(r"^word/(document|footnotes|endnotes|comments|header\d*|footer\d*)\.xml$")
TRACKED = {W + "ins": "ins", W + "del": "del", W + "moveFrom": "del", W + "moveTo": "ins"}
KINDS = [
    ("ADDIN ZOTERO_ITEM", "ZOTERO_ITEM"),
    ("ADDIN ZOTERO_BIBL", "ZOTERO_BIBL"),
    ("ADDIN EN.CITE.DATA", "EN.CITE.DATA"),
    ("ADDIN EN.CITE", "EN.CITE"),
    ("ADDIN EN.REFLIST", "EN.REFLIST"),
]


def kind_of(code):
    c = code.strip()
    for prefix, name in KINDS:
        if c.startswith(prefix):
            return name
    return (c.split() or ["(empty)"])[0][:24]


def digest(code):
    return hashlib.sha1(code.encode("utf-8")).hexdigest()[:12]


def targets(z, part):
    """Relationship id -> target of one part."""
    rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if rels not in z.namelist():
        return {}
    return {r.get("Id"): r.get("Target", "") for r in ET.fromstring(z.read(rels))}


class Story:
    """The fields of one part, in document order."""

    def __init__(self, name, rels):
        self.name = name
        self.rels = rels
        self.fields = []
        self.errors = []
        self.paragraphs = []
        self.open = []
        self.para = []

    def error(self, what, field=None):
        p = self.para[-1] if self.para else None
        self.errors.append({"what": what, "story": self.name, "paragraph": p, "field": field})

    def add(self, field):
        p = self.para[-1] if self.para else None
        field.update(
            story=self.name,
            paragraph=p,
            index=len(self.fields),
            depth=len(self.open),
            layout=[],
            code="",
            data=None,
            result=[],
        )
        self.fields.append(field)
        return field

    def mark(self, field, token, ctx):
        item = f"{token}@{ctx or '-'}"
        if not field["layout"] or field["layout"][-1] != item:
            field["layout"].append(item)
        field.setdefault("ctx", {}).setdefault(token, set()).add(ctx or "-")

    def walk(self, el, ctx=None):
        tag = el.tag
        ctx = TRACKED.get(tag, ctx)
        if tag == W + "p":
            row = {"index": len(self.paragraphs), "text": "", "depth_start": len(self.open)}
            row.update(begin=0, separate=0, end=0)
            self.paragraphs.append(row)
            self.para.append(row)
        elif tag == W + "fldChar":
            self.marker(el, ctx)
        elif tag in (W + "instrText", W + "delInstrText"):
            if self.open and self.open[-1]["phase"] == "code":
                self.open[-1]["code"] += el.text or ""
                self.mark(self.open[-1], "instr", ctx)
            else:
                self.error("an instruction outside the code of a field")
        elif tag in (W + "t", W + "delText"):
            if self.para:
                self.para[-1]["text"] += el.text or ""
            if self.open and self.open[-1]["phase"] == "result":
                self.open[-1]["result"].append((el.text or "", ctx))
                self.mark(self.open[-1], "result", ctx)
        elif tag in (W + "fldSimple", W + "hyperlink"):
            self.simple(el, ctx)
        for child in el:
            self.walk(child, ctx)
        if tag == W + "p":
            self.para.pop()["depth_end"] = len(self.open)

    def marker(self, el, ctx):
        kind = el.get(W + "fldCharType")
        if self.para:
            self.para[-1][kind] = self.para[-1].get(kind, 0) + 1
        if kind == "begin":
            f = self.add({"form": "complex", "phase": "code", "closed": False})
            data = el.find(W + "fldData")
            if data is not None:
                f["data"] = "".join((data.text or "").split())
            self.mark(f, "begin", ctx)
            self.open.append(f)
        elif not self.open:
            self.error(f"{kind} without a begin")
        elif kind == "separate":
            f = self.open[-1]
            if f["phase"] == "result":
                self.error("a second separate in one field", f["index"])
            f["phase"] = "result"
            self.mark(f, "separate", ctx)
        elif kind == "end":
            f = self.open.pop()
            f["closed"] = True
            self.mark(f, "end", ctx)

    def simple(self, el, ctx):
        if el.tag == W + "hyperlink":
            where = self.rels.get(el.get(R_ID), "")
            anchor = el.get(W + "anchor")
            code = "HYPERLINK " + where + (f" #{anchor}" if anchor else "")
            f = self.add({"form": "hyperlink", "phase": "done", "closed": True})
        else:
            code = el.get(W + "instr", "")
            f = self.add({"form": "simple", "phase": "done", "closed": True})
            data = el.find(W + "fldData")
            if data is not None:
                f["data"] = "".join((data.text or "").split())
        f["code"] = code
        self.mark(f, "begin", ctx)
        self.mark(f, "end", ctx)

    def finish(self):
        for f in self.open:
            self.errors.append(
                {
                    "what": "a begin without an end",
                    "story": self.name,
                    "paragraph": f["paragraph"],
                    "field": f["index"],
                }
            )


def state_of(field):
    """live, deleted, inserted, or split (the markers and the instruction
    do not lie in one kind of place)."""
    ctx = field.get("ctx", {})
    places = set()
    for token in ("begin", "instr", "separate", "end"):
        places |= ctx.get(token, set())
    if len(places) != 1:
        return "split"
    return {"-": "live", "del": "deleted", "ins": "inserted"}[places.pop()]


def recorded_text(field):
    """The citation text the instruction records, or None."""
    code = field["code"]
    if field["kind"] == "ZOTERO_ITEM":
        try:
            data = json.loads(code[code.index("{") : code.rindex("}") + 1])
            return data.get("properties", {}).get("plainCitation")
        except ValueError:
            return None
    if field["kind"] == "EN.CITE":
        m = re.search(r"<DisplayText>(.*?)</DisplayText>", code)
        return m.group(1) if m else None
    return None


def read(path):
    out = {"path": str(path), "fields": [], "errors": [], "paragraphs": [], "props": {}, "vars": {}}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        for name in sorted(n for n in names if STORY.match(n)):
            story = Story(name, targets(z, name))
            story.walk(ET.fromstring(z.read(name)))
            story.finish()
            for f in story.fields:
                f["kind"] = "hyperlink" if f["form"] == "hyperlink" else kind_of(f["code"])
                f["id"] = digest(f["code"])
                f["state"] = state_of(f) if f["closed"] else "open"
                f["text"] = (f["paragraph"] or {}).get("text", "")
            out["fields"] += story.fields
            out["errors"] += story.errors
            out["paragraphs"] += [
                {"story": name, **{k: v for k, v in p.items() if k != "text"}, "text": p["text"][:40]}
                for p in story.paragraphs
                if p["begin"] or p["separate"] or p["end"] or p["depth_start"] or p["depth_end"]
            ]
        if "docProps/custom.xml" in names:
            for prop in ET.fromstring(z.read("docProps/custom.xml")):
                name = prop.get("name", "")
                if name.startswith("ZOTERO"):
                    out["props"][name] = "".join(prop.itertext())
        if "word/settings.xml" in names:
            for var in ET.fromstring(z.read("word/settings.xml")).iter(W + "docVar"):
                name = var.get(W + "name", "")
                if name.startswith(("EN.", "ZOTERO")):
                    out["vars"][name] = var.get(W + "val", "")
    return out


def brief(f):
    return {
        "id": f["id"],
        "kind": f["kind"],
        "state": f["state"],
        "story": f["story"],
        "paragraph": (f["paragraph"] or {}).get("index"),
        "text": f["text"][:40],
        "layout": " ".join(f["layout"]),
    }


def summary(doc):
    fields = doc["fields"]
    return {
        "fields": len(fields),
        "by_kind": dict(sorted(Counter(f["kind"] for f in fields).items())),
        "by_state": dict(sorted(Counter(f["state"] for f in fields).items())),
        "with_data": sum(1 for f in fields if f["data"]),
        "list": [{"id": f["id"], "kind": f["kind"], "state": f["state"]} for f in fields],
    }


def line(kind, f, before_layout, after_layout):
    return (
        f'FAIL {kind} | {f["story"]} p{(f["paragraph"] or {}).get("index", "?")} '
        f'"{f["text"][:40]}" | kind={f["kind"]} | before: {before_layout} | after: {after_layout}'
    )


def compare(before, after, expect_removed=None):
    failures = []
    lines = []
    by_id = {}
    for f in before["fields"]:
        by_id.setdefault(f["id"], []).append(f)

    def before_layout(f):
        twins = by_id.get(f["id"]) or []
        return " ".join(twins[0]["layout"]) if twins else "(no field with this instruction)"

    for name, doc in (("before", before), ("after", after)):
        for e in doc["errors"]:
            p = e["paragraph"] or {}
            failures.append(f"unbalanced-{name}")
            fields = [f for f in doc["fields"] if f["index"] == e["field"] and f["story"] == e["story"]]
            kind = fields[0]["kind"] if fields else "?"
            layout = " ".join(fields[0]["layout"]) if fields else "-"
            row = f"begin {p.get('begin', 0)} separate {p.get('separate', 0)} end {p.get('end', 0)}"
            lines.append(
                f'FAIL unbalanced ({name}): {e["what"]} | {e["story"]} p{p.get("index", "?")} '
                f'"{p.get("text", "")[:40]}" | kind={kind} | markers of the paragraph: {row} | '
                f"layout: {layout}"
            )
    split = [f for f in after["fields"] if f["state"] == "split"]
    for f in split:
        failures.append("split-field")
        lines.append(line("split-field", f, before_layout(f), " ".join(f["layout"])))

    a = Counter(f["id"] for f in before["fields"])
    b = Counter(f["id"] for f in after["fields"])
    missing = a - b
    added = b - a
    n_missing = sum(missing.values())
    if added:
        failures.append("instruction-added-or-changed")
    if expect_removed is None and missing:
        failures.append("instruction-missing-or-changed")
    if expect_removed is not None and n_missing != expect_removed:
        failures.append("removed-count")
        lines.append(f"FAIL removed-count | {n_missing} instruction(s) gone, expected {expect_removed}")
    for fid in added:
        f = next(x for x in after["fields"] if x["id"] == fid)
        lines.append(line("instruction-added-or-changed", f, "(absent)", " ".join(f["layout"])))
    if expect_removed is None or n_missing != expect_removed:
        for fid in missing:
            f = by_id[fid][0]
            lines.append(line("instruction-missing-or-changed", f, " ".join(f["layout"]), "(absent)"))

    da = Counter(f["data"] for f in before["fields"] if f["data"])
    db = Counter(f["data"] for f in after["fields"] if f["data"])
    data = {"before": sum(da.values()), "after": sum(db.values()), "added": sum((db - da).values())}
    data["missing"] = sum((da - db).values())
    if data["added"] or (expect_removed is None and data["missing"]):
        failures.append("fldData")
        lines.append(f'FAIL fldData | missing {data["missing"]}, not in the original {data["added"]}')

    props = {}
    for key in ("props", "vars"):
        names = sorted(set(before[key]) | set(after[key]))
        diff = [n for n in names if before[key].get(n) != after[key].get(n)]
        props[key] = {"before": len(before[key]), "after": len(after[key]), "differ": diff}
        if diff:
            failures.append("property" if key == "props" else "document-variable")
            for n in diff:
                state = "lost" if n not in after[key] else "new" if n not in before[key] else "changed"
                lines.append(f"FAIL {key} | {n} {state}")

    changed = []
    for f in after["fields"]:
        want = recorded_text(f) if f["state"] == "live" else None
        if want is None:
            continue
        accepted = "".join(t for t, ctx in f["result"] if ctx != "del")
        rejected = "".join(t for t, ctx in f["result"] if ctx != "ins")
        if accepted != want or rejected != want:
            changed.append({**brief(f), "recorded": want, "accepted": accepted, "rejected": rejected})

    return lines, {
        "ok": not failures,
        "failures": sorted(set(failures)),
        "expect_removed": expect_removed,
        "before": summary(before),
        "after": summary(after),
        "missing": [brief(by_id[i][0]) for i in missing for _ in range(missing[i])],
        "added": [brief(next(x for x in after["fields"] if x["id"] == i)) for i in added],
        "split": [brief(f) for f in split],
        "unbalanced": [
            {"what": e["what"], "story": e["story"], "paragraph": (e["paragraph"] or {}).get("index")}
            for e in after["errors"]
        ],
        "unbalanced_before": len(before["errors"]),
        "fld_data": data,
        "properties": props["props"],
        "doc_vars": props["vars"],
        "text_changed": changed,
        "paragraphs": after["paragraphs"],
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--before", required=True)
    ap.add_argument("--after", required=True)
    ap.add_argument("--expect-removed", type=int, default=None)
    ap.add_argument(
        "--detail",
        action="store_true",
        help="the JSON line also lists every field (before.list, after.list) and the "
        "markers of every paragraph that holds one (paragraphs)",
    )
    args = ap.parse_args(argv)
    try:
        before, after = read(args.before), read(args.after)
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError) as e:
        print(json.dumps({"ok": False, "failures": ["unreadable"], "error": str(e)}))
        return 2
    lines, result = compare(before, after, args.expect_removed)
    if not args.detail:
        del result["paragraphs"], result["before"]["list"], result["after"]["list"]
    for text in lines:
        print(text)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
