#!/usr/bin/env python3
"""Check an anonymized docx: what identifying metadata is left, and that the text is unchanged.

Every category of tools/anonymize_docx.py is counted in --after and gets one of three states:

  cleared          nothing of it is left
  kept by option   something is left and it was not asked to be cleared (kept by default, or --keep)
  not cleared      something is left that should be gone, or that no tool here can clear

Categories (default): names, org, custom_props, comment_authors, people, change_authors, dates,
core_dates, core_text, rsids, doc_vars, attached_template, doc_ids, custom_xml, thumbnail, image_meta
(cleared by default); title, para_ids (kept by default); hidden_text, protection, image_uninspected,
embedded (reported only: the anonymizer never changes them, so they are "not cleared" whenever present).
Pass the same --keep / --also that the anonymizer was given.

--after is also compared with --before: the text of every w:t (body, notes, headers, footers,
comments) and the number of revision marks must be the same, and no relationship or content-type
entry may point at a part that is gone.

The XML is read with a parser, so the check does not depend on prefixes or quoting. An author name
written in the text itself stays there; `name_mentions` says how often, for information only.

Exit code: 0 when every category that should be cleared is at 0, text and marks are equal and the
package is consistent; 1 otherwise; 2 when a file cannot be read. Reported-only categories do not
change the exit code: read `not_cleared` in the output.

Usage: verify-anonymized.py --before SOURCE.docx --after RESULT.docx
       [--author Anonymous] [--initials A] [--keep CATEGORY,...] [--also CATEGORY,...] [--json]
"""

import argparse
import collections
import json
import posixpath
import re
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
BIBLIOGRAPHY = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"

FIXED_DATE = "1970-01-01T00:00:00Z"

# id -> (default, identity-bearing, what it is). tools/anonymize_docx.py carries the same ids and defaults.
CATEGORIES = collections.OrderedDict([
    ("names", ("clear", True, "creator and last-modified-by in core.xml")),
    ("org", ("clear", True, "Company, Manager, HyperlinkBase in app.xml")),
    ("custom_props", ("clear", True, "custom document properties")),
    ("comment_authors", ("clear", True, "comment authors and initials")),
    ("people", ("clear", True, "people.xml persons and presence accounts")),
    ("change_authors", ("clear", True, "tracked-change authors")),
    ("dates", ("clear", True, "dates of tracked changes and comments")),
    ("core_dates", ("clear", True, "created, modified, last printed, revision number")),
    ("core_text", ("clear", True, "keywords, description, subject, category, identifier, status")),
    ("rsids", ("clear", True, "revision save ids (editing-session fingerprints)")),
    ("doc_vars", ("clear", True, "document variables in settings.xml")),
    ("attached_template", ("clear", True, "attached template and its path")),
    ("doc_ids", ("clear", True, "document ids in settings.xml")),
    ("custom_xml", ("clear", True, "custom XML items other than bibliography sources")),
    ("thumbnail", ("clear", True, "thumbnail picture of page one")),
    ("image_meta", ("clear", True, "EXIF, XMP, IPTC, comments and text chunks in JPEG and PNG pictures")),
    ("title", ("keep", False, "document title in core.xml")),
    ("para_ids", ("keep", False, "paragraph ids outside comments.xml")),
    ("hidden_text", ("report", True, "hidden text runs (never removed: the text must stay)")),
    ("protection", ("report", False, "document or write protection settings")),
    ("image_uninspected", ("report", True, "TIFF pictures (their metadata is not inspected)")),
    ("embedded", ("report", True, "embedded objects (they keep their own metadata)")),
])  # fmt: skip

BODY_PART = re.compile(r"^word/(?:document|footnotes|endnotes|header\d*|footer\d*)\.xml$")
REVISION_TAGS = (
    "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange",
    "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange", "numberingChange",
    "cellIns", "cellDel", "cellMerge",
)  # fmt: skip


def q(ns, tag):
    return "{%s}%s" % (ns, tag)


def local(name):
    return name.rsplit("}", 1)[-1]


class Package:
    """The members of a docx: names, raw bytes on demand, parsed XML parts (None when not well-formed)."""

    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.names = z.namelist()
            self.raw = {}
            self.xml = {}
            for name in self.names:
                is_xml = name.endswith((".xml", ".rels")) and (
                    name.startswith(("docProps/", "word/", "customXml/", "_rels/")) or name == "[Content_Types].xml"
                )
                if is_xml:
                    try:
                        self.xml[name] = ET.fromstring(z.read(name))
                    except ET.ParseError:
                        self.xml[name] = None
                elif name.startswith(("word/media/", "word/embeddings/", "docProps/")):
                    self.raw[name] = z.read(name)

    def part(self, name):
        return self.xml.get(name)

    def word_parts(self):
        return [(n, r) for n, r in sorted(self.xml.items()) if n.startswith("word/") and n.endswith(".xml") and r is not None]


def texts(root, ns, tag):
    return [(e.text or "").strip() for e in root.iter(q(ns, tag))] if root is not None else []


# ---------------------------------------------------------------- pictures (an own reading, not the anonymizer's)
def jpeg_has_meta(data):
    """True when a JPEG carries EXIF/XMP (APP1), IPTC (APP13) or a comment; None when it cannot be read."""
    if data[:2] != b"\xff\xd8":
        return None
    pos = 2
    while pos + 4 <= len(data):
        if data[pos] != 0xFF:
            return None
        marker = data[pos + 1]
        if marker == 0xFF:
            pos += 1
        elif marker == 0xDA:
            return False
        elif marker == 0x01 or 0xD0 <= marker <= 0xD8:
            pos += 2
        elif marker in (0xE1, 0xED, 0xFE):
            return True
        else:
            pos += 2 + struct.unpack(">H", data[pos + 2 : pos + 4])[0]
    return None


def png_has_meta(data):
    """True when a PNG carries a text, EXIF or time chunk; None when it cannot be read."""
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    pos = 8
    while pos + 12 <= len(data):
        length, kind = struct.unpack(">I", data[pos : pos + 4])[0], data[pos + 4 : pos + 8]
        if kind in (b"tEXt", b"zTXt", b"iTXt", b"eXIf", b"tIME"):
            return True
        if kind == b"IEND":
            return False
        pos += 12 + length
    return None


# ---------------------------------------------------------------- counting
def count_marks(pkg, author, initials):
    """[(category, key, count)] of everything left in the package. Keys are "<part>:<target>"."""
    out = []

    def add(category, key, n, always=False):
        if n or always:
            out.append((category, key, n))

    core, app = pkg.part("docProps/core.xml"), pkg.part("docProps/app.xml")
    add("names", "docProps/core.xml:creator", sum(1 for t in texts(core, DC, "creator") if t and t != author), True)
    add("names", "docProps/core.xml:lastModifiedBy", sum(1 for t in texts(core, CP, "lastModifiedBy") if t and t != author), True)
    for tag in ("Company", "Manager", "HyperlinkBase"):
        add("org", "docProps/app.xml:" + tag, sum(1 for t in texts(app, APP, tag) if t), tag != "HyperlinkBase")
    custom = pkg.part("docProps/custom.xml")
    add("custom_props", "docProps/custom.xml:property", len(list(custom.iter(q(CUSTOM, "property")))) if custom is not None else 0, True)
    for ns, tag in ((DCTERMS, "created"), (DCTERMS, "modified"), (CP, "lastPrinted")):
        add("core_dates", "docProps/core.xml:" + tag, sum(1 for t in texts(core, ns, tag) if t not in ("", FIXED_DATE)))
    add("core_dates", "docProps/core.xml:revision", sum(1 for t in texts(core, CP, "revision") if t not in ("", "1")))
    for ns, tag in ((CP, "keywords"), (DC, "description"), (DC, "subject"), (CP, "category"), (DC, "identifier"), (CP, "contentStatus")):
        add("core_text", "docProps/core.xml:" + tag, sum(1 for t in texts(core, ns, tag) if t))
    add("title", "docProps/core.xml:title", sum(1 for t in texts(core, DC, "title") if t))

    comments = pkg.part("word/comments.xml")
    found = list(comments.iter()) if comments is not None else []
    add("comment_authors", "word/comments.xml:author", sum(1 for e in found if e.get(q(W, "author"), author) not in ("", author)), True)
    add("comment_authors", "word/comments.xml:initials", sum(1 for e in found if e.get(q(W, "initials"), initials) not in ("", initials)), True)
    people = pkg.part("word/people.xml")
    persons = list(people.iter(q(W15, "person"))) if people is not None else []
    add("people", "word/people.xml:person", sum(1 for p in persons if p.get(q(W15, "author"), "") != author), True)
    add("people", "word/people.xml:presenceInfo", len(list(people.iter(q(W15, "presenceInfo")))) if people is not None else 0, True)

    for name, root in pkg.word_parts():
        attrs = [(k, v) for e in root.iter() for k, v in e.attrib.items()]
        if name not in ("word/comments.xml", "word/people.xml"):
            n = sum(1 for k, v in attrs if k == q(W, "author") and v != author)
            add("change_authors", name + ":change_author", n, bool(BODY_PART.match(name)))
        add("dates", name + ":date", sum(1 for k, v in attrs if (k == q(W, "date") or local(k) == "dateUtc") and v != FIXED_DATE))
        add("rsids", name + ":rsid", sum(1 for k, _ in attrs if k.startswith(q(W, "rsid"))))
        if name != "word/comments.xml":
            add("para_ids", name + ":paraId", sum(1 for k, _ in attrs if k in (q(W14, "paraId"), q(W14, "textId"))))
        if BODY_PART.match(name):
            hidden = 0
            for r in root.iter(q(W, "r")):
                vanish = r.find("%s/%s" % (q(W, "rPr"), q(W, "vanish")))
                on = vanish is not None and vanish.get(q(W, "val"), "true") not in ("0", "false", "off")
                if on and any((t.text or "") for t in r.iter(q(W, "t"))):
                    hidden += 1
            add("hidden_text", name + ":hidden_text", hidden)

    settings = pkg.part("word/settings.xml")
    if settings is not None:
        tags = collections.Counter(e.tag for e in settings.iter())
        add("rsids", "word/settings.xml:rsids", tags[q(W, "rsid")] + tags[q(W, "rsidRoot")])
        add("doc_vars", "word/settings.xml:docVar", tags[q(W, "docVar")])
        add("attached_template", "word/settings.xml:attachedTemplate", tags[q(W, "attachedTemplate")])
        add("doc_ids", "word/settings.xml:docId", tags[q(W14, "docId")] + tags[q(W15, "docId")])
        add("protection", "word/settings.xml:protection", tags[q(W, "documentProtection")] + tags[q(W, "writeProtection")])
    rels = pkg.part("word/_rels/settings.xml.rels")
    if rels is not None:
        n = sum(1 for e in rels.iter(q(REL, "Relationship")) if e.get("Type", "").endswith("/attachedTemplate"))
        add("attached_template", "word/_rels/settings.xml.rels:attachedTemplate", n)

    items = [n for n in pkg.names if re.fullmatch(r"customXml/item\d+\.xml", n)]
    other = [n for n in items if pkg.part(n) is None or pkg.part(n).tag != q(BIBLIOGRAPHY, "Sources")]
    add("custom_xml", "customXml:item", len(other))
    add("thumbnail", "docProps:thumbnail", sum(1 for n in pkg.names if re.fullmatch(r"docProps/thumbnail\.[^/]+", n)))
    meta = tiff = 0
    for name in pkg.names:
        low = name.lower()
        if not name.startswith("word/media/"):
            continue
        if low.endswith((".jpg", ".jpeg")):
            meta += jpeg_has_meta(pkg.raw[name]) is not False
        elif low.endswith(".png"):
            meta += png_has_meta(pkg.raw[name]) is not False
        elif low.endswith((".tif", ".tiff")):
            tiff += 1
    add("image_meta", "word/media:with_metadata", meta)
    add("image_uninspected", "word/media:tiff", tiff)
    add("embedded", "word/embeddings:object", sum(1 for n in pkg.names if n.startswith("word/embeddings/") and not n.endswith("/")))
    return out


def dangling(pkg):
    """Internal relationship targets and content-type overrides that name a member the package does not have."""
    names, out = set(pkg.names), set()
    for name, root in pkg.xml.items():
        if root is None:
            continue
        if name.endswith(".rels"):
            base = posixpath.dirname(posixpath.dirname(name))
            for e in root.iter(q(REL, "Relationship")):
                target = e.get("Target", "")
                if e.get("TargetMode") == "External" or not target:
                    continue
                member = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
                if member.split("#")[0] not in names:
                    out.add("%s -> %s" % (name, member))
        elif name == "[Content_Types].xml":
            for e in root.iter(q(TYPES, "Override")):
                if e.get("PartName", "").lstrip("/") not in names:
                    out.add("%s -> %s" % (name, e.get("PartName")))
    return out


def body_text(pkg):
    """{part: text of every w:t} for the body parts and the comments."""
    return {
        name: "".join(e.text or "" for e in root.iter(q(W, "t")))
        for name, root in pkg.word_parts()
        if BODY_PART.match(name) or name == "word/comments.xml"
    }


def markup(pkg):
    """Counter of revision marks per part."""
    tags = {q(W, t) for t in REVISION_TAGS}
    return collections.Counter(
        (name, e.tag) for name, root in pkg.word_parts() if BODY_PART.match(name)
        for e in root.iter() if e.tag in tags
    )  # fmt: skip


def names_in(pkg):
    """Author names the source carries in its marks (three letters or more)."""
    names = set(texts(pkg.part("docProps/core.xml"), DC, "creator") + texts(pkg.part("docProps/core.xml"), CP, "lastModifiedBy"))
    for _, root in pkg.word_parts():
        for e in root.iter():
            names.update(v for k, v in e.attrib.items() if k in (q(W, "author"), q(W15, "author")))
    return {n for n in names if len(n) >= 3}


def verify(before_path, after_path, author="Anonymous", initials="A", keep=(), also=()):
    before, after = Package(before_path), Package(after_path)
    expected = ({c for c, v in CATEGORIES.items() if v[0] == "clear"} | set(also)) - set(keep)
    marks = count_marks(after, author, initials)
    counts = collections.Counter()
    for category, _, n in marks:
        counts[category] += n
    categories = collections.OrderedDict()
    for cid, (default, identity, what) in CATEGORIES.items():
        n = counts[cid]
        if n == 0:
            status = "cleared"
        elif cid in expected or default == "report":
            status = "not cleared"
        else:
            status = "kept by option"
        categories[cid] = {"status": status, "count": n, "default": default, "identity": identity, "what": what}
    mentions = {}
    after_text = body_text(after)
    for name in names_in(before) - {author}:
        n = sum(t.count(name) for t in after_text.values())
        if n:
            mentions[name] = n
    new_dangling = sorted(dangling(after) - dangling(before))
    return {
        "categories": categories,
        "remaining": collections.OrderedDict((key, n) for _, key, n in marks),
        "total": sum(counts[c] for c in expected),
        "not_cleared": [c for c, v in categories.items() if v["status"] == "not cleared"],
        "body_equal": body_text(before) == after_text,
        "markup_equal": markup(before) == markup(after),
        "structure_ok": not new_dangling,
        "dangling": new_dangling,
        "unreadable_parts": sorted(name for name, root in after.xml.items() if root is None),
        "name_mentions": mentions,
    }


def passed(rep):
    return (
        rep["total"] == 0 and rep["body_equal"] and rep["markup_equal"] and rep["structure_ok"] and not rep["unreadable_parts"]
    )


def categories_arg(text):
    """argparse type: a comma-separated list of switchable category names."""
    names = [t.strip() for t in text.split(",") if t.strip()]
    unknown = [n for n in names if n not in CATEGORIES or CATEGORIES[n][0] == "report"]
    if unknown:
        raise argparse.ArgumentTypeError("not a switchable category: %s" % ", ".join(unknown))
    return names


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--before", required=True, help="the source docx")
    ap.add_argument("--after", required=True, help="the anonymized docx")
    ap.add_argument("--author", default="Anonymous", help="the anonymous name that is allowed")
    ap.add_argument("--initials", default="A", help="the anonymous initials that are allowed")
    ap.add_argument("--keep", type=categories_arg, default=[], help="categories the anonymizer was told to keep")
    ap.add_argument("--also", type=categories_arg, default=[], help="categories the anonymizer was told to clear as well")
    ap.add_argument("--json", action="store_true", help="print the report as JSON")
    a = ap.parse_args(argv)
    try:
        rep = verify(a.before, a.after, a.author, a.initials, a.keep, a.also)
    except (OSError, zipfile.BadZipFile) as e:
        print("cannot read %s or %s: %s" % (a.before, a.after, e), file=sys.stderr)
        return 2
    if a.json:
        print(json.dumps(rep, ensure_ascii=False))
    else:
        for cid, c in rep["categories"].items():
            print("%-18s %-15s %3d  %s" % (cid, c["status"], c["count"], c["what"]))
        print("total %d, body_equal %s, markup_equal %s, structure_ok %s" % (rep["total"], rep["body_equal"], rep["markup_equal"], rep["structure_ok"]))
        if rep["not_cleared"]:
            print("NOT CLEARED: %s" % ", ".join(rep["not_cleared"]))
        for name in rep["unreadable_parts"]:
            print("not well-formed XML: %s" % name)
        for line in rep["dangling"]:
            print("points at a missing part: %s" % line)
        for n in rep["name_mentions"].values():
            print("note: a source author name still stands %d time(s) in the text (the text is not anonymized)" % n)
    return 0 if passed(rep) else 1


if __name__ == "__main__":
    sys.exit(main())
