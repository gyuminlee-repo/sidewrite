#!/usr/bin/env python3
"""Check that an anonymized docx keeps no author mark and that its text is unchanged.

Counts, part by part, the author marks that are still in --after:

  docProps/core.xml    creator, lastModifiedBy   (anything but the anonymous name or empty)
  docProps/app.xml     Company, Manager          (anything but empty)
  docProps/custom.xml  property                  (every custom property)
  word/comments.xml    author, initials          (anything but the anonymous name / initials)
  word/people.xml      person, presenceInfo      (a person under another name; every presenceInfo)
  word/document.xml, footnotes, endnotes, headers, footers
                       change_author             (a tracked-change mark under another name)

and compares --after with --before: the text of every w:t (body, notes, headers,
footers, comments) and the number of revision marks must be the same.

Reads the XML with a parser, so it does not depend on prefixes or quoting. It
does not look at dates, rsids, settings.xml or customXml/ (not covered yet), and
it does not judge the body text: an author name written in the text stays there
(`name_mentions` reports how often, for information only).

Exit code: 0 when nothing is left and text and marks are equal, 1 otherwise,
2 when a file cannot be read.

Usage: verify-anonymized.py --before SOURCE.docx --after RESULT.docx
       [--author Anonymous] [--initials A] [--json]
"""

import argparse
import collections
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"

BODY_PART = re.compile(r"^word/(?:document|footnotes|endnotes|header\d*|footer\d*)\.xml$")
REVISION_TAGS = (
    "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange",
    "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange", "numberingChange",
    "cellIns", "cellDel", "cellMerge",
)  # fmt: skip


def q(ns, tag):
    return "{%s}%s" % (ns, tag)


def read_parts(path):
    """{part name: parsed root} for every XML part that matters here."""
    out = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith(("docProps/", "word/")) and name.endswith(".xml"):
                try:
                    out[name] = ET.fromstring(z.read(name))
                except ET.ParseError:
                    out[name] = None
    return out


def text_of(root, ns, tag):
    return [(e.text or "").strip() for e in root.iter(q(ns, tag))]


def remaining(parts, author, initials):
    """Counts of what is left, keyed "<part>:<target>"."""
    rem = collections.OrderedDict()
    core, app = parts.get("docProps/core.xml"), parts.get("docProps/app.xml")
    for tag, ns in (("creator", DC), ("lastModifiedBy", CP)):
        rem["docProps/core.xml:" + tag] = (
            sum(1 for t in text_of(core, ns, tag) if t and t != author) if core is not None else 0
        )
    for tag in ("Company", "Manager"):
        rem["docProps/app.xml:" + tag] = sum(1 for t in text_of(app, APP, tag) if t) if app is not None else 0
    custom = parts.get("docProps/custom.xml")
    rem["docProps/custom.xml:property"] = len(list(custom.iter(q(CUSTOM, "property")))) if custom is not None else 0
    comments = parts.get("word/comments.xml")
    found = list(comments.iter(q(W, "comment"))) if comments is not None else []
    rem["word/comments.xml:author"] = sum(1 for c in found if c.get(q(W, "author"), "") not in ("", author))
    rem["word/comments.xml:initials"] = sum(1 for c in found if c.get(q(W, "initials"), "") not in ("", initials))
    people = parts.get("word/people.xml")
    persons = list(people.iter(q(W15, "person"))) if people is not None else []
    rem["word/people.xml:person"] = sum(1 for p in persons if p.get(q(W15, "author"), "") != author)
    rem["word/people.xml:presenceInfo"] = len(list(people.iter(q(W15, "presenceInfo")))) if people is not None else 0
    for name in sorted(parts):
        if BODY_PART.match(name) and parts[name] is not None:
            rem[name + ":change_author"] = sum(
                1 for e in parts[name].iter() if e.get(q(W, "author"), author) != author
            )
    return rem


def body_text(parts):
    """{part: text of every w:t} for the body parts and the comments."""
    return {
        name: "".join(e.text or "" for e in root.iter(q(W, "t")))
        for name, root in parts.items()
        if root is not None and (BODY_PART.match(name) or name == "word/comments.xml")
    }


def markup(parts):
    """Counter of revision marks per part."""
    tags = {q(W, t) for t in REVISION_TAGS}
    return collections.Counter(
        (name, e.tag) for name, root in parts.items() if root is not None and BODY_PART.match(name)
        for e in root.iter() if e.tag in tags
    )  # fmt: skip


def names_in(parts):
    """Author names the source carries in its marks (three letters or more)."""
    names = set()
    for part, ns, tag in (("docProps/core.xml", DC, "creator"), ("docProps/core.xml", CP, "lastModifiedBy")):
        if parts.get(part) is not None:
            names.update(text_of(parts[part], ns, tag))
    for name, root in parts.items():
        if root is not None and name.startswith("word/"):
            for e in root.iter():
                names.update(v for k, v in e.attrib.items() if k in (q(W, "author"), q(W15, "author")))
    return {n for n in names if len(n) >= 3}


def verify(before_path, after_path, author="Anonymous", initials="A"):
    before, after = read_parts(before_path), read_parts(after_path)
    rem = remaining(after, author, initials)
    broken = sorted(name for name, root in after.items() if root is None)
    mentions = {}
    for name in names_in(before) - {author}:
        n = sum(t.count(name) for t in body_text(after).values())
        if n:
            mentions[name] = n
    return {
        "remaining": rem,
        "total": sum(rem.values()),
        "body_equal": body_text(before) == body_text(after),
        "markup_equal": markup(before) == markup(after),
        "unreadable_parts": broken,
        "name_mentions": mentions,
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--before", required=True, help="the source docx")
    ap.add_argument("--after", required=True, help="the anonymized docx")
    ap.add_argument("--author", default="Anonymous", help="the anonymous name that is allowed")
    ap.add_argument("--initials", default="A", help="the anonymous initials that are allowed")
    ap.add_argument("--json", action="store_true", help="print the report as JSON")
    a = ap.parse_args(argv)
    try:
        rep = verify(a.before, a.after, a.author, a.initials)
    except (OSError, zipfile.BadZipFile) as e:
        print("cannot read %s or %s: %s" % (a.before, a.after, e), file=sys.stderr)
        return 2
    if a.json:
        print(json.dumps(rep, ensure_ascii=False, indent=1))
    else:
        for key, n in rep["remaining"].items():
            print("%-45s %d" % (key, n))
        print("total %d, body_equal %s, markup_equal %s" % (rep["total"], rep["body_equal"], rep["markup_equal"]))
        for name in rep["unreadable_parts"]:
            print("not well-formed XML: %s" % name)
        for name, n in rep["name_mentions"].items():
            print("note: a source author name still stands %d time(s) in the text (the text is not anonymized)" % n)
    ok = rep["total"] == 0 and rep["body_equal"] and rep["markup_equal"] and not rep["unreadable_parts"]
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
