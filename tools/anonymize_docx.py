#!/usr/bin/env python3
"""Write a copy of a docx with its author marks replaced by an anonymous name.

What is changed (nothing else is touched; the text of the document stays as it is):

  docProps/core.xml    dc:creator, cp:lastModifiedBy        -> the anonymous name
  docProps/app.xml     Company, Manager                     -> empty
  docProps/custom.xml  every custom property                -> removed
  word/comments.xml    w:author, w:initials                 -> the anonymous name / initials
  word/people.xml      every person and presenceInfo        -> one person, the anonymous name
  every word/*.xml     w:author of tracked changes          -> the anonymous name

Not changed: dates of comments and changes, rsid attributes, word/settings.xml,
customXml/, and any name written in the text itself. Check the result with
scripts/verify-anonymized.py.

The parts are edited as text, not re-serialised, so namespace declarations and
everything else in a part stay byte for byte; each edited part is then parsed
to make sure it is still well-formed. The zip keeps its part order.

The output is never the source: an output path that is the source file under
any name (the same path, a link, another spelling) is refused. The copy is
written to a temporary file next to the output and moved into place when it is
complete.

Standard library only.

Usage: anonymize_docx.py --in SOURCE.docx --out COPY.docx [--author Anonymous] [--initials A]
Exit code: 0 done (a JSON report on stdout), 2 refused or failed (the reason on stderr).
"""

import argparse
import copy
import json
import os
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"

NAME = r"[A-Za-z_][\w.\-]*"


class Refused(Exception):
    """The request cannot be carried out; the message says why."""


def prefixes(xml, uri):
    """Every prefix bound to `uri` in this part ('' is the default namespace)."""
    found = re.findall(r"\bxmlns(?::(%s))?\s*=\s*([\"'])%s\2" % (NAME, re.escape(uri)), xml)
    return {p for p, _ in found}


def attr_value(text):
    """`text` as a double-quoted XML attribute value."""
    return '"%s"' % escape(text, {'"': "&quot;"})


def set_attr(xml, prefix, name, value, counts, key):
    """Set every `prefix:name="..."` attribute to `value`; count the ones that change."""
    new = attr_value(value)
    pat = re.compile(r"(\s%s:%s\s*=\s*)(\"[^\"]*\"|'[^']*')" % (re.escape(prefix), name))

    def sub(m):
        if m.group(2) == new:
            return m.group(0)
        counts[key] = counts.get(key, 0) + 1
        return m.group(1) + new

    return pat.sub(sub, xml)


def set_text(xml, prefix, name, value, counts, key):
    """Set the text of every `<prefix:name>` element to `value`; count the ones that change."""
    tag = "%s:%s" % (prefix, name) if prefix else name
    new = escape(value)
    pat = re.compile(r"(<%s(?:\s[^>]*)?>)(.*?)(</%s\s*>)" % (re.escape(tag), re.escape(tag)), re.S)

    def sub(m):
        if m.group(2) == new or (not m.group(2).strip() and not new):
            return m.group(0)
        counts[key] = counts.get(key, 0) + 1
        return m.group(1) + new + m.group(3)

    return pat.sub(sub, xml)


def root_span(xml):
    """(end of the root start tag, start of the root end tag); None when the root has no content."""
    blank = re.sub(r"<\?.*?\?>|<!--.*?-->", lambda x: " " * len(x.group()), xml, flags=re.S)
    m = re.search(r"<((?:%s:)?%s)(?:\s[^>]*?)?(/?)>" % (NAME, NAME), blank)
    if not m or m.group(2):
        return None
    end = xml.rfind("</" + m.group(1))
    return (m.end(), end) if end >= m.end() else None


def scrub_part(name, xml, author, initials):
    """Return (new text, {target: count}) for one part."""
    counts = {}
    if name == "docProps/core.xml":
        for p in prefixes(xml, DC):
            xml = set_text(xml, p, "creator", author, counts, "creator")
        for p in prefixes(xml, CP):
            xml = set_text(xml, p, "lastModifiedBy", author, counts, "lastModifiedBy")
    elif name == "docProps/app.xml":
        for tag in ("Company", "Manager"):
            for p in set(re.findall(r"<(?:(%s):)?%s[\s>]" % (NAME, tag), xml)):
                xml = set_text(xml, p, tag, "", counts, tag)
    elif name == "docProps/custom.xml":
        n = len(re.findall(r"<(?:%s:)?property[\s>]" % NAME, xml))
        span = root_span(xml)
        if n and span:
            xml = xml[: span[0]] + xml[span[1] :]
            counts["property"] = n
    elif name == "word/people.xml":
        span = root_span(xml)
        for p in prefixes(xml, W15):
            persons = re.findall(r"<%s:person\b[^>]*>" % re.escape(p), xml)
            presence = len(re.findall(r"<%s:presenceInfo\b" % re.escape(p), xml))
            one = "<%s:person %s:author=%s/>" % (p, p, attr_value(author))
            if p and span and persons and (presence or persons != [one]):
                xml = xml[: span[0]] + one + xml[span[1] :]
                counts["person"] = len(persons)
                if presence:
                    counts["presenceInfo"] = presence
                break
    elif name.startswith("word/") and name.endswith(".xml"):
        for p in prefixes(xml, W):
            if p:
                xml = set_attr(xml, p, "author", author, counts, "author")
                if name == "word/comments.xml":
                    xml = set_attr(xml, p, "initials", initials, counts, "initials")
    return xml, counts


def anonymize_bytes(name, data, author, initials):
    """Return (new bytes, counts) for one zip member; the same bytes for anything that is not a target."""
    if not (name.startswith("docProps/") or name.startswith("word/")) or not name.endswith(".xml"):
        return data, {}
    try:
        xml = data.decode("utf-8")
    except UnicodeDecodeError:
        raise Refused("%s is not UTF-8; this tool does not rewrite it" % name)
    new, counts = scrub_part(name, xml, author, initials)
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


def anonymize(src, out, author="Anonymous", initials="A"):
    """Write the anonymized copy of `src` to `out`; return the report."""
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
            zout.comment = zin.comment
            for info in zin.infolist():
                data, counts = anonymize_bytes(info.filename, zin.read(info.filename), author, initials)
                if counts:
                    changed[info.filename] = counts
                zout.writestr(copy.copy(info), data, compress_type=info.compress_type)
        if same_file(src, out):  # the output became a link to the source while the copy was written
            raise Refused("the output is the same file as the source: %s" % out)
        os.replace(tmp, out)
    finally:
        if os.path.exists(tmp):  # anything but a completed move: leave no half-written file
            os.unlink(tmp)
    return {"output": str(out), "author": author, "initials": initials, "changed": changed}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--in", dest="src", required=True, help="the source docx (never modified)")
    ap.add_argument("--out", required=True, help="the anonymized copy to write")
    ap.add_argument("--author", default="Anonymous", help="the name written in place of every author")
    ap.add_argument("--initials", default="A", help="the initials written on comments")
    a = ap.parse_args(argv)
    try:
        report = anonymize(a.src, a.out, a.author, a.initials)
    except (Refused, OSError, zipfile.BadZipFile) as e:
        print(str(e), file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
