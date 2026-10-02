#!/usr/bin/env python3
"""Are the fields of a .docx still whole after it was edited?

  citation_fields_check.py --before A.docx --after B.docx [--expect-removed N]
                           [--fld-data judge|note] [--detail]

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
  - no field of B is hollowed: a complex field that is still there (its
    markers are not all inside tracked deletions) while its result, the runs
    between separate and end, has no text left once the tracked changes are
    accepted, although the result had text (in B itself, inside a deletion,
    or in the same field of A). Deleting the visible text of a citation and
    leaving its markers does this: the citation looks deleted, the field
    stays, and the reference manager writes the text again at its next
    refresh and keeps the reference in the bibliography. A field whose
    markers AND result all lie inside tracked deletions is deleted as a
    whole (state "deleted", shown as "in deletions") and is not hollowed;
  - no result of a field that is still there differs from the result the
    same field has in A (compared with the tracked changes accepted on both
    sides): a result with a tracked deletion or insertion inside it, or
    other words. This fails as result-changed, because the text of a field
    is written by the reference manager: a tool that edits it leaves a
    citation that no longer says what its instruction records, and the next
    refresh overwrites the edit or asks the author about it;
  - the w:fldData of B are those of A (white space apart); with
    --expect-removed, B has none that A lacks. With --fld-data note a
    difference is printed as a NOTE line and does not fail: for a file whose
    w:fldData Word itself rewrites on every save (the invented EndNote data
    of the test fixture) the difference says nothing about a tool;
  - the ZOTERO_* properties of docProps/custom.xml and the EN.* and ZOTERO*
    document variables of word/settings.xml are the same.
A itself must be balanced too: a broken A makes the comparison mean nothing.

The results of w:fldSimple and w:hyperlink are not read: hollowed and
result-changed are told for complex fields only.

Not judged, only listed:
  text_changed  a live Zotero or EndNote citation whose visible text is no
                longer the text its instruction records (plainCitation,
                DisplayText). A is not looked at: a citation the author
                changed by hand before the edits is listed here and does not
                fail; one an edit changed fails as result-changed above.
  flattened     a field that lies whole in a tracked deletion while its
                visible text stands again, as plain inserted text, in the
                same paragraph: an edit replaced a range that held the field.
                Nothing is half, but once the change is accepted the citation
                is plain text and the plugin no longer knows it.

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
            row.update(begin=0, separate=0, end=0, inserted="")
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
                if ctx == "ins" and not self.open:
                    self.para[-1]["inserted"] += el.text or ""
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
                {
                    "story": name,
                    **{k: v for k, v in p.items() if k not in ("text", "inserted")},
                    "text": p["text"][:40],
                }
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


def accepted_text(field):
    """The result as it reads once every tracked change is accepted."""
    return "".join(t for t, ctx in field["result"] if ctx != "del")


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


def compare(before, after, expect_removed=None, fld_data="judge"):
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
    data["differs"] = bool(data["added"] or (expect_removed is None and data["missing"]))
    data["judged"] = fld_data != "note"
    if data["differs"]:
        counts = f'missing {data["missing"]}, not in the original {data["added"]}'
        if data["judged"]:
            failures.append("fldData")
            lines.append(f"FAIL fldData | {counts}")
        else:
            lines.append(f"NOTE fldData | {counts} (not judged: --fld-data note)")

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

    # A field that is still there with nothing left of its result (hollowed),
    # or with another result than the same field has in the before file.
    was = {}
    for f in before["fields"]:
        if f["state"] != "deleted":
            was.setdefault(f["id"], []).append(accepted_text(f))
    hollowed = []
    result_changed = []
    for f in after["fields"]:
        if f["form"] != "complex" or not f["closed"] or f["state"] == "deleted":
            continue
        now = accepted_text(f)
        twins = was.get(f["id"], [])
        # Its own text, deleted or not: a result that lies whole in a pending
        # deletion is hollowed in every copy that still holds the deletion.
        own = "".join(t for t, _ in f["result"])
        if now in twins and (now.strip() or not own.strip()):
            twins.remove(now)
            continue
        if not now.strip() and (own.strip() or "".join(twins).strip()):
            old = next((t for t in twins if t.strip()), own)
            hollowed.append({**brief(f), "was": old})
            failures.append("hollowed")
            lines.append(
                line("hollowed", f, before_layout(f), " ".join(f["layout"]))
                + f" | result: {json.dumps(old, ensure_ascii=False)} -> {json.dumps(now)}"
            )
        elif twins:
            old = twins.pop(0)
            result_changed.append({**brief(f), "was": old, "accepted": now})
            failures.append("result-changed")
            lines.append(
                line("result-changed", f, before_layout(f), " ".join(f["layout"]))
                + f" | result: {json.dumps(old, ensure_ascii=False)} -> "
                + json.dumps(now, ensure_ascii=False)
            )

    changed = []
    for f in after["fields"]:
        want = recorded_text(f) if f["state"] == "live" else None
        if want is None:
            continue
        accepted = "".join(t for t, ctx in f["result"] if ctx != "del")
        rejected = "".join(t for t, ctx in f["result"] if ctx != "ins")
        if accepted != want or rejected != want:
            changed.append({**brief(f), "recorded": want, "accepted": accepted, "rejected": rejected})

    # A field that lies whole in a deletion while its visible text was typed
    # again next to it as plain inserted text: the edit took the field for
    # words. Whole by every rule above; after accept-all the citation is text.
    flattened = []
    for f in after["fields"]:
        shown = "".join(t for t, _ in f["result"]).strip()
        if f["state"] == "deleted" and shown and shown in (f["paragraph"] or {}).get("inserted", ""):
            flattened.append({**brief(f), "shown": shown})

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
        "hollowed": hollowed,
        "result_changed": result_changed,
        "text_changed": changed,
        "flattened": flattened,
        "paragraphs": after["paragraphs"],
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--before", required=True)
    ap.add_argument("--after", required=True)
    ap.add_argument("--expect-removed", type=int, default=None)
    ap.add_argument(
        "--fld-data",
        choices=("judge", "note"),
        default="judge",
        help="note: a difference of w:fldData is printed and does not fail",
    )
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
    lines, result = compare(before, after, args.expect_removed, args.fld_data)
    if not args.detail:
        del result["paragraphs"], result["before"]["list"], result["after"]["list"]
    for text in lines:
        print(text)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
