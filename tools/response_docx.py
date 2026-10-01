#!/usr/bin/env python3
"""Write a response letter .docx from a paragraph model. Standard library only.

  python3 tools/response_docx.py --model model.json --out letter.docx

Model: {"font": str|null, "size_pt": number|null,
        "paragraphs": [{"kind": str, "text": str, "color": "RRGGBB"|null}]}

One paragraph per model row, one run per paragraph. The run carries the
colour; headings (kinds in BOLD) are bold. The same model always gives the
same bytes (fixed zip timestamps), so a re-export without changes is no diff.

Prints one JSON line. Exit 0 written, 3 the target is locked (open in Word),
1 any other error.
"""

import argparse
import json
import os
import re
import sys
import uuid
import zipfile
from xml.sax.saxutils import escape, quoteattr

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    "</Types>"
)
RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
    'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
)
BOLD = {"reviewer_head", "section_head", "comment_head"}
HEX = re.compile(r"^[0-9A-Fa-f]{6}$")
# Characters XML 1.0 does not allow (Word's U+0005 comment anchor among them).
CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
STAMP = (2000, 1, 1, 0, 0, 0)


class Locked(Exception):
    """The target file cannot be replaced: another program holds it open."""


def run_xml(text, color=None, bold=False, font=None, size_pt=None):
    props = []
    if font:
        f = quoteattr(str(font))
        props.append("<w:rFonts w:ascii=%s w:hAnsi=%s w:cs=%s/>" % (f, f, f))
    if bold:
        props.append("<w:b/>")
    if color:
        if not HEX.match(str(color)):
            raise ValueError("colour is not RRGGBB: %r" % (color,))
        props.append('<w:color w:val="%s"/>' % color.upper())
    if size_pt:
        props.append('<w:sz w:val="%d"/>' % round(float(size_pt) * 2))
    rpr = "<w:rPr>%s</w:rPr>" % "".join(props) if props else ""
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, escape(CONTROL.sub("", str(text))))


def document_xml(model):
    font, size = model.get("font"), model.get("size_pt")
    paras = [
        "<w:p>%s</w:p>" % run_xml(p.get("text", ""), p.get("color"), p.get("kind") in BOLD, font, size)
        for p in model["paragraphs"]
    ]
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="%s"><w:body>%s%s</w:body></w:document>' % (W, "".join(paras), SECT)
    )


def write_zip(path, parts):
    """path: a file name or a binary file open for writing. parts: [(name, text)].
    Fixed timestamps and attributes: same input, same bytes."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:
            info = zipfile.ZipInfo(name, STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o644 << 16
            z.writestr(info, data.encode("utf-8") if isinstance(data, str) else data)


def write_docx(model, out):
    """Write the letter to `out` and return the paragraph count.

    The zip is written to a temp file of this call's own in the folder of
    `out` (same file system, so the move over `out` is atomic) and created
    exclusively: two writers of one path never share a temp file, and a file
    that was already there is never touched. Only that own file is cleaned up.
    """
    xml = document_xml(model)  # raises before anything is written
    out = str(out)
    folder, base = os.path.split(os.path.abspath(out))
    tmp = os.path.join(folder, ".%s.%d.%s.tmp" % (base, os.getpid(), uuid.uuid4().hex))
    with open(tmp, "xb") as fh:  # from here on the file is ours
        try:
            write_zip(fh, [("[Content_Types].xml", CONTENT_TYPES), ("_rels/.rels", RELS), ("word/document.xml", xml)])
            fh.close()
            try:
                os.replace(tmp, out)
            except PermissionError as e:
                raise Locked(str(e)) from e
        finally:
            fh.close()
            if os.path.exists(tmp):
                os.remove(tmp)
    return len(model["paragraphs"])


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--model", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    try:
        with open(a.model, encoding="utf-8") as fh:
            model = json.load(fh)
        n = write_docx(model, a.out)
    except Locked as e:
        print(json.dumps({"ok": False, "code": "locked", "error": str(e)}))
        return 3
    except (OSError, ValueError, KeyError, TypeError, AttributeError) as e:
        print(json.dumps({"ok": False, "code": "error", "error": str(e)}))
        return 1
    print(json.dumps({"ok": True, "path": a.out, "paragraphs": n}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
