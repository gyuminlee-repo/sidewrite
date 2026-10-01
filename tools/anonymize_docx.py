#!/usr/bin/env python3
"""Write a copy of a docx with its author marks and other identifying metadata cleared.

The text of the document is never changed. What is cleared is chosen by category; every
category can be switched (--keep a,b / --also c,d). Defaults:

  cleared by default
    names              docProps/core.xml creator, lastModifiedBy       -> the anonymous name
    org                docProps/app.xml Company, Manager, HyperlinkBase -> empty
    custom_props       docProps/custom.xml custom properties            -> removed
    comment_authors    word/comments.xml w:author, w:initials           -> the anonymous name / initials
    people             word/people.xml persons and presenceInfo         -> one person, the anonymous name
    change_authors     w:author of tracked changes (every word/*.xml)   -> the anonymous name
    dates              w:date of tracked changes and comments, dateUtc  -> a fixed date (--dates remove: removed)
    core_dates         core.xml created, modified, lastPrinted, revision -> the fixed date, revision 1
    core_text          core.xml keywords, description, subject, category, identifier, contentStatus -> empty
    rsids              w:rsid* attributes and the rsids table in settings.xml -> removed (they tell editing
                       sessions apart; they are optional in the file format)
    doc_vars           settings.xml document variables                   -> removed
    attached_template  settings.xml attachedTemplate and the template path in its relationship -> removed
    doc_ids            settings.xml w14:docId, w15:docId                 -> removed
    custom_xml         customXml/ items other than the bibliography sources, with their references -> removed
    thumbnail          docProps/thumbnail.* (a picture of page one) and its relationship -> removed
    image_meta         EXIF, XMP, IPTC and comments in word/media JPEG files; text, EXIF and time chunks in PNG
                       files -> removed. A picture whose EXIF says it is rotated is left as it is (removing
                       the tag could turn the picture) and the verifier reports it.

  kept by default (clear with --also)
    title              core.xml title -> empty
    para_ids           w14:paraId, w14:textId outside comments.xml -> removed. Inside comments.xml they stay
                       in any case: commentsExtended.xml and commentsIds.xml link comments through them.

  never changed, only reported by the verifier
    hidden text (w:vanish), document protection, TIFF pictures, embedded objects, app.xml TotalTime,
    Template and application version, and every name written in the text itself.

Both values of --dates are valid in the file format (the date attribute is optional). Not checked: whether
Word opens every result without a repair prompt; check the result with scripts/verify-anonymized.py.

The XML parts are edited as text, attribute by attribute inside tags only, not re-serialised, so namespace
declarations and everything else stay byte for byte; each edited part is then parsed to make sure it is
still well-formed. Pictures are edited by dropping whole metadata segments; the picture data is not touched.

The output is never the source: an output path that is the source file under any name (the same path, a
link, another spelling) is refused. The copy is written to a temporary file next to the output and moved
into place when it is complete.

Standard library only.

Usage: anonymize_docx.py --in SOURCE.docx --out COPY.docx [--author Anonymous] [--initials A]
       [--keep CATEGORY,...] [--also CATEGORY,...] [--dates fixed|remove]
Exit code: 0 done (a JSON report on stdout), 2 refused or failed (the reason on stderr).
"""

import argparse
import copy
import json
import os
import posixpath
import re
import struct
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
DATE_UTC_NS = (
    "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
)
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
BIBLIOGRAPHY = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"

FIXED_DATE = "1970-01-01T00:00:00Z"
NAME = r"[A-Za-z_][\w.\-]*"

# id -> cleared by default. scripts/verify-anonymized.py carries the same table.
CATEGORIES = {
    "names": True, "org": True, "custom_props": True, "comment_authors": True, "people": True,
    "change_authors": True, "dates": True, "core_dates": True, "core_text": True, "rsids": True,
    "doc_vars": True, "attached_template": True, "doc_ids": True, "custom_xml": True, "thumbnail": True,
    "image_meta": True, "title": False, "para_ids": False,
}  # fmt: skip


class Refused(Exception):
    """The request cannot be carried out; the message says why."""


# ---------------------------------------------------------------- XML as text
def prefixes(xml, *uris):
    """Every non-empty prefix bound to one of `uris` in this part."""
    out = set()
    for uri in uris:
        out.update(p for p, _ in re.findall(r"\bxmlns:(%s)\s*=\s*([\"'])%s\2" % (NAME, re.escape(uri)), xml))
    return out


def attr_value(text):
    """`text` as a double-quoted XML attribute value."""
    return '"%s"' % escape(text, {'"': "&quot;"})


TAG = re.compile(r"<[^<>!?/][^<>]*>")
ATTR = re.compile(r"(\s+)(%s:%s)(\s*=\s*)(\"[^\"]*\"|'[^']*')" % (NAME, NAME))
DROP = object()


def edit_attrs(xml, decide):
    """Rewrite attributes inside start tags only: decide(qname, quoted value) -> None (keep), DROP, or a new quoted value."""

    def attr(a):
        new = decide(a.group(2), a.group(4))
        if new is None:
            return a.group(0)
        return "" if new is DROP else a.group(1) + a.group(2) + a.group(3) + new

    return TAG.sub(lambda m: ATTR.sub(attr, m.group(0)), xml)


def set_text(xml, prefix, name, value, counts, key, keep=()):
    """Set the text of every `<prefix:name>` element to `value`; count the ones that change."""
    tag = "%s:%s" % (prefix, name) if prefix else name
    new = escape(value)
    pat = re.compile(r"(<%s(?:\s[^>]*)?>)(.*?)(</%s\s*>)" % (re.escape(tag), re.escape(tag)), re.S)

    def sub(m):
        old = m.group(2)
        if old == new or old.strip() in keep or (not old.strip() and not new):
            return m.group(0)
        counts[key] = counts.get(key, 0) + 1
        return m.group(1) + new + m.group(3)

    return pat.sub(sub, xml)


def drop_elements(xml, qname, counts, key, child=None):
    """Remove every `<qname>` element (with its content); count it, or its `child` elements when given."""
    q = re.escape(qname)
    pat = re.compile(r"<%s(?:\s[^>]*)?(?:/>|>.*?</%s\s*>)" % (q, q), re.S)

    def sub(m):
        inner = re.sub(r"^<[^>]*>", "", m.group(0))
        n = len(re.findall(r"<%s[\s/>]" % child, inner)) if child else 1
        counts[key] = counts.get(key, 0) + n
        return ""

    return pat.sub(sub, xml)


def root_span(xml):
    """(end of the root start tag, start of the root end tag); None when the root has no content."""
    blank = re.sub(r"<\?.*?\?>|<!--.*?-->", lambda x: " " * len(x.group()), xml, flags=re.S)
    m = re.search(r"<((?:%s:)?%s)(?:\s[^>]*?)?(/?)>" % (NAME, NAME), blank)
    if not m or m.group(2):
        return None
    end = xml.rfind("</" + m.group(1))
    return (m.end(), end) if end >= m.end() else None


def scrub_core(xml, author, on, counts):
    for p in prefixes(xml, DC):
        if "names" in on:
            xml = set_text(xml, p, "creator", author, counts, "creator")
        if "core_text" in on:
            for tag in ("description", "subject", "identifier"):
                xml = set_text(xml, p, tag, "", counts, tag)
        if "title" in on:
            xml = set_text(xml, p, "title", "", counts, "title")
    for p in prefixes(xml, CP):
        if "names" in on:
            xml = set_text(xml, p, "lastModifiedBy", author, counts, "lastModifiedBy")
        if "core_text" in on:
            for tag in ("keywords", "category", "contentStatus"):
                xml = set_text(xml, p, tag, "", counts, tag)
        if "core_dates" in on:
            xml = set_text(xml, p, "lastPrinted", FIXED_DATE, counts, "lastPrinted")
            xml = set_text(xml, p, "revision", "1", counts, "revision")
    if "core_dates" in on:
        for p in prefixes(xml, DCTERMS):
            for tag in ("created", "modified"):
                xml = set_text(xml, p, tag, FIXED_DATE, counts, tag)
    return xml


def scrub_people(xml, author, counts):
    span = root_span(xml)
    for p in prefixes(xml, W15):
        persons = re.findall(r"<%s:person\b[^>]*>" % re.escape(p), xml)
        presence = len(re.findall(r"<%s:presenceInfo\b" % re.escape(p), xml))
        one = "<%s:person %s:author=%s/>" % (p, p, attr_value(author))
        if span and persons and (presence or persons != [one]):
            counts["person"] = len(persons)
            if presence:
                counts["presenceInfo"] = presence
            return xml[: span[0]] + one + xml[span[1] :]
    return xml


def scrub_word(name, xml, author, initials, on, dates, counts):
    """Attributes of one word/*.xml part, and the settings elements."""
    w, w14 = prefixes(xml, W), prefixes(xml, W14)
    utc = prefixes(xml, *DATE_UTC_NS)
    in_comments = name == "word/comments.xml"
    author_on = ("comment_authors" if in_comments else "change_authors") in on
    new_author, new_initials, new_date = attr_value(author), attr_value(initials), attr_value(FIXED_DATE)

    def hit(key):
        counts[key] = counts.get(key, 0) + 1

    def decide(qname, value):
        prefix, local = qname.split(":", 1)
        if prefix in w:
            if local == "author" and author_on and value != new_author:
                hit("author")
                return new_author
            if local == "initials" and in_comments and "comment_authors" in on and value != new_initials:
                hit("initials")
                return new_initials
            if local == "date" and "dates" in on and (dates == "remove" or value != new_date):
                hit("date")
                return DROP if dates == "remove" else new_date
            if local.startswith("rsid") and "rsids" in on:
                hit("rsid")
                return DROP
        if prefix in utc and local == "dateUtc" and "dates" in on and (dates == "remove" or value != new_date):
            hit("date")
            return DROP if dates == "remove" else new_date
        if prefix in w14 and local in ("paraId", "textId") and "para_ids" in on and not in_comments:
            hit("paraId")
            return DROP
        return None

    xml = edit_attrs(xml, decide)
    if name == "word/settings.xml":
        for p in w:
            if "rsids" in on:
                xml = drop_elements(xml, p + ":rsids", counts, "rsid", child=p + ":rsid\\w*")
            if "doc_vars" in on:
                xml = drop_elements(xml, p + ":docVars", counts, "docVar", child=p + ":docVar")
            if "attached_template" in on:
                xml = drop_elements(xml, p + ":attachedTemplate", counts, "attachedTemplate")
        if "doc_ids" in on:
            for p in w14 | prefixes(xml, W15):
                xml = drop_elements(xml, p + ":docId", counts, "docId")
    return xml


def scrub_part(name, xml, author, initials, on, dates="fixed"):
    """Return (new text, {target: count}) for one XML part."""
    counts = {}
    if name == "docProps/core.xml":
        xml = scrub_core(xml, author, on, counts)
    elif name == "docProps/app.xml":
        if "org" in on:
            for tag in ("Company", "Manager", "HyperlinkBase"):
                for p in set(re.findall(r"<(?:(%s):)?%s[\s>]" % (NAME, tag), xml)):
                    xml = set_text(xml, p, tag, "", counts, tag)
    elif name == "docProps/custom.xml":
        n = len(re.findall(r"<(?:%s:)?property[\s>]" % NAME, xml))
        span = root_span(xml)
        if "custom_props" in on and n and span:
            xml = xml[: span[0]] + xml[span[1] :]
            counts["property"] = n
    elif name == "word/people.xml":
        if "people" in on:
            xml = scrub_people(xml, author, counts)
    elif name.startswith("word/") and name.endswith(".xml"):
        xml = scrub_word(name, xml, author, initials, on, dates, counts)
    return xml, counts


# ---------------------------------------------------------------- relationships and content types
def rels_base(rels_name):
    """The folder a relationship part's targets are relative to: 'a/_rels/b.xml.rels' -> 'a'."""
    return posixpath.dirname(posixpath.dirname(rels_name))


def resolve(rels_name, target):
    """The zip member a relationship target names."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(rels_base(rels_name), target))


RELATIONSHIP = re.compile(r"<(?:%s:)?Relationship\b[^>]*?/>|<(?:%s:)?Relationship\b[^>]*?>\s*</(?:%s:)?Relationship\s*>" % ((NAME,) * 3))


def attr_of(tag, name):
    m = re.search(r"\s%s\s*=\s*(\"[^\"]*\"|'[^']*')" % name, tag)
    return m.group(1)[1:-1] if m else None


def scrub_rels(name, xml, dropped, on, counts):
    """Remove relationships to removed members, and the attached-template relationship."""

    def sub(m):
        tag = m.group(0)
        kind, target = attr_of(tag, "Type") or "", attr_of(tag, "Target") or ""
        external = attr_of(tag, "TargetMode") == "External"
        if not external and resolve(name, target) in dropped:
            counts["relationship"] = counts.get("relationship", 0) + 1
            return ""
        if "attached_template" in on and kind.endswith("/attachedTemplate"):
            counts["attachedTemplate"] = counts.get("attachedTemplate", 0) + 1
            return ""
        return tag

    return RELATIONSHIP.sub(sub, xml)


def scrub_content_types(xml, dropped, counts):
    def sub(m):
        part = (attr_of(m.group(0), "PartName") or "").lstrip("/")
        if part in dropped:
            counts["override"] = counts.get("override", 0) + 1
            return ""
        return m.group(0)

    return re.sub(r"<(?:%s:)?Override\b[^>]*?/>" % NAME, sub, xml)


def members_to_drop(zin, on):
    """Zip members that are removed whole: the thumbnail, and custom XML items that are not bibliography sources."""
    names = zin.namelist()
    drop = set()
    if "thumbnail" in on:
        drop.update(n for n in names if re.fullmatch(r"docProps/thumbnail\.[^/]+", n))
    if "custom_xml" in on:
        for n in names:
            if not re.fullmatch(r"customXml/item\d+\.xml", n):
                continue
            try:
                root = ET.fromstring(zin.read(n)).tag
            except ET.ParseError:
                root = ""
            if root == "{%s}Sources" % BIBLIOGRAPHY:
                continue
            drop.add(n)
            rels = "customXml/_rels/%s.rels" % posixpath.basename(n)
            if rels in names:
                drop.add(rels)
                for tag in RELATIONSHIP.findall(zin.read(rels).decode("utf-8", "replace")):
                    target = resolve(rels, attr_of(tag, "Target") or "")
                    if attr_of(tag, "TargetMode") != "External" and target in names:
                        drop.add(target)
    return drop


# ---------------------------------------------------------------- pictures
def exif_orientation(tiff):
    """The orientation an EXIF block names: None (no tag), 1..8, or -1 when the block cannot be read."""
    order = {b"II": "<", b"MM": ">"}.get(bytes(tiff[:2]))
    if order is None or len(tiff) < 8 or struct.unpack(order + "H", tiff[2:4])[0] != 42:
        return -1
    at = struct.unpack(order + "I", tiff[4:8])[0]
    if at + 2 > len(tiff):
        return -1
    n = struct.unpack(order + "H", tiff[at : at + 2])[0]
    for i in range(n):
        entry = tiff[at + 2 + 12 * i : at + 14 + 12 * i]
        if len(entry) < 12:
            return -1
        if struct.unpack(order + "H", entry[:2])[0] == 0x0112:
            return struct.unpack(order + "H", entry[8:10])[0]
    return None


JPEG_META = (0xE1, 0xED, 0xFE)  # APP1 (EXIF, XMP), APP13 (IPTC), comment


def strip_jpeg(data):
    """(new bytes, segments removed, state): 'none' no metadata, 'stripped', 'rotated' (left as it is), 'unreadable'."""
    if data[:2] != b"\xff\xd8":
        return data, 0, "unreadable"
    pos, spans, orientation = 2, [], None
    while pos + 4 <= len(data):
        if data[pos] != 0xFF:
            return data, 0, "unreadable"
        marker = data[pos + 1]
        if marker == 0xFF:
            pos += 1
            continue
        if marker == 0xDA:  # start of scan: the picture data follows and is never touched
            if not spans:
                return data, 0, "none"
            if orientation not in (None, 1):
                return data, 0, "rotated"
            out, last = [], 0
            for a, b in spans:
                out.append(data[last:a])
                last = b
            return b"".join(out) + data[last:], len(spans), "stripped"
        if marker == 0x01 or 0xD0 <= marker <= 0xD8:
            pos += 2
            continue
        length = struct.unpack(">H", data[pos + 2 : pos + 4])[0]
        end = pos + 2 + length
        if marker == 0xD9 or length < 2 or end > len(data):
            return data, 0, "unreadable"
        if marker in JPEG_META:
            spans.append((pos, end))
            if marker == 0xE1 and data[pos + 4 : pos + 10] == b"Exif\x00\x00" and orientation is None:
                orientation = exif_orientation(data[pos + 10 : end])
        pos = end
    return data, 0, "unreadable"


PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
PNG_META = (b"tEXt", b"zTXt", b"iTXt", b"eXIf", b"tIME")


def strip_png(data):
    """Like strip_jpeg, for the text, EXIF and time chunks of a PNG."""
    if data[:8] != PNG_SIGNATURE:
        return data, 0, "unreadable"
    pos, keep, removed, orientation, ended = 8, [data[:8]], 0, None, False
    while pos + 12 <= len(data) and not ended:
        length = struct.unpack(">I", data[pos : pos + 4])[0]
        kind, end = data[pos + 4 : pos + 8], pos + 12 + length
        if end > len(data):
            return data, 0, "unreadable"
        if kind in PNG_META:
            removed += 1
            if kind == b"eXIf":
                orientation = exif_orientation(data[pos + 8 : pos + 8 + length])
        else:
            keep.append(data[pos:end])
        ended = kind == b"IEND"
        pos = end
    if not ended:
        return data, 0, "unreadable"
    if not removed:
        return data, 0, "none"
    if orientation not in (None, 1):
        return data, 0, "rotated"
    return b"".join(keep) + data[pos:], removed, "stripped"


def strip_picture(name, data):
    lower = name.lower()
    if lower.endswith((".jpg", ".jpeg")):
        return strip_jpeg(data)
    if lower.endswith(".png"):
        return strip_png(data)
    return data, 0, "none"


# ---------------------------------------------------------------- the copy
def anonymize_member(name, data, author, initials, on, dates, dropped):
    """Return (new bytes, counts) for one zip member; the same bytes for anything that is not a target."""
    counts = {}
    if name.startswith("word/media/"):
        if "image_meta" in on:
            new, n, state = strip_picture(name, data)
            if state == "stripped":
                return new, {"segments": n}
            if state in ("rotated", "unreadable"):
                return data, {"left_" + state: 1}
        return data, {}
    is_rels = name.endswith(".rels")
    if not (is_rels or name == "[Content_Types].xml" or (name.startswith(("docProps/", "word/")) and name.endswith(".xml"))):
        return data, {}
    try:
        xml = data.decode("utf-8")
    except UnicodeDecodeError:
        raise Refused("%s is not UTF-8; this tool does not rewrite it" % name)
    if is_rels:
        new = scrub_rels(name, xml, dropped, on, counts)
    elif name == "[Content_Types].xml":
        new = scrub_content_types(xml, dropped, counts)
    else:
        new, counts = scrub_part(name, xml, author, initials, on, dates)
    if not counts:
        return data, {}
    out = new.encode("utf-8")
    try:
        ET.fromstring(out)
    except ET.ParseError as e:
        raise Refused("%s would not be well-formed after the edit (%s); nothing was written" % (name, e))
    return out, counts


def same_file(a, b):
    """True when two paths name one file: equal real paths, or one existing file reached two ways."""
    ra, rb = os.path.realpath(a), os.path.realpath(b)
    if ra == rb or os.path.normcase(ra) == os.path.normcase(rb):
        return True
    return os.path.exists(a) and os.path.exists(b) and os.path.samefile(a, b)


def anonymize(src, out, author="Anonymous", initials="A", on=None, dates="fixed"):
    """Write the anonymized copy of `src` to `out`; return the report."""
    on = {c for c, default in CATEGORIES.items() if default} if on is None else set(on)
    if not os.path.isfile(src):
        raise Refused("no such file: %s" % src)
    if same_file(src, out):
        raise Refused("the output is the same file as the source: %s" % out)
    out_dir = os.path.dirname(os.path.abspath(out))
    if not os.path.isdir(out_dir):
        raise Refused("the output folder does not exist: %s" % out_dir)
    if not zipfile.is_zipfile(src):
        raise Refused("not a docx (not a zip file): %s" % src)
    changed = {}
    fd, tmp = tempfile.mkstemp(prefix=".anonymize-", suffix=".tmp", dir=out_dir)
    try:
        with os.fdopen(fd, "wb") as fh, zipfile.ZipFile(src) as zin, zipfile.ZipFile(fh, "w") as zout:
            if "word/document.xml" not in zin.namelist():
                raise Refused("not a docx (no word/document.xml): %s" % src)
            dropped = members_to_drop(zin, on)
            zout.comment = zin.comment
            for info in zin.infolist():
                if info.filename in dropped:
                    continue
                data, counts = anonymize_member(info.filename, zin.read(info.filename), author, initials, on, dates, dropped)
                if counts:
                    changed[info.filename] = counts
                zout.writestr(copy.copy(info), data, compress_type=info.compress_type)
        if same_file(src, out):  # the output became a link to the source while the copy was written
            raise Refused("the output is the same file as the source: %s" % out)
        os.replace(tmp, out)
    finally:
        if os.path.exists(tmp):  # anything but a completed move: leave no half-written file
            os.unlink(tmp)
    return {
        "output": str(out),
        "author": author,
        "initials": initials,
        "cleared": sorted(on),
        "kept": sorted(set(CATEGORIES) - on),
        "dates": dates,
        "changed": changed,
        "removed": sorted(dropped),
    }


def categories(text):
    """argparse type: a comma-separated list of category names."""
    names = [t.strip() for t in text.split(",") if t.strip()]
    unknown = [n for n in names if n not in CATEGORIES]
    if unknown:
        raise argparse.ArgumentTypeError("unknown category %s; known: %s" % (", ".join(unknown), ", ".join(CATEGORIES)))
    return names


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--in", dest="src", required=True, help="the source docx (never modified)")
    ap.add_argument("--out", required=True, help="the anonymized copy to write")
    ap.add_argument("--author", default="Anonymous", help="the name written in place of every author")
    ap.add_argument("--initials", default="A", help="the initials written on comments")
    ap.add_argument("--keep", type=categories, default=[], help="categories to leave as they are")
    ap.add_argument("--also", type=categories, default=[], help="categories to clear that are kept by default")
    ap.add_argument("--dates", choices=("fixed", "remove"), default="fixed", help="what becomes of change and comment dates")
    a = ap.parse_args(argv)
    on = ({c for c, default in CATEGORIES.items() if default} | set(a.also)) - set(a.keep)
    try:
        report = anonymize(a.src, a.out, a.author, a.initials, on, a.dates)
    except (Refused, OSError, zipfile.BadZipFile) as e:
        print(str(e), file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
