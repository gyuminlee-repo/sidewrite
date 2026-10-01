#!/usr/bin/env python3
"""Build a small .docx whose paragraphs hold tracked insertions and deletions
with the run elements that are text without being w:t: non-breaking hyphen,
soft hyphen, tab, line break, carriage return, symbol.

The acceptance run opens it in Word and compares, for every paragraph, the
text Word reports (Paragraph.text) with the tracked runs the task pane reads
from the paragraph OOXML (taskpane/shared/review-view.js trackedRuns): each
tracked run must be found in the paragraph text.

Usage: make_tracked_special.py out.docx [--count]
(--count: also print the number of paragraphs with a tracked run, last line)
"""

import sys
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

AUTHOR = 'w:author="Fixture" w:date="2026-01-01T00:00:00Z"'

# (name, kind, run content XML inside the tracked run). T is w:t for an
# insertion and w:delText for a deletion.
CASES = [
    ("plain", "ins", "<T>plain words</T>"),
    ("no-break hyphen", "ins", "<T>non</T><w:noBreakHyphen/><T>linear</T>"),
    ("soft hyphen", "ins", "<T>micro</T><w:softHyphen/><T>scope</T>"),
    ("tab", "ins", "<T>left</T><w:tab/><T>right</T>"),
    ("line break", "ins", "<T>first</T><w:br/><T>second</T>"),
    ("carriage return", "ins", "<T>one</T><w:cr/><T>two</T>"),
    ("symbol", "ins", '<T>alpha </T><w:sym w:font="Symbol" w:char="F061"/><T> end</T>'),
    ("symbol wingdings", "ins", '<T>mark </T><w:sym w:font="Wingdings" w:char="F0FC"/><T> end</T>'),
    ("entities and spaces", "ins", "<T>a &amp; b  &lt;c&gt; </T>"),
    ("deleted no-break hyphen", "del", "<T>non</T><w:noBreakHyphen/><T>linear</T>"),
    ("deleted soft hyphen", "del", "<T>micro</T><w:softHyphen/><T>scope</T>"),
    ("deleted tab", "del", "<T>left</T><w:tab/><T>right</T>"),
    ("deleted symbol", "del", '<T>alpha </T><w:sym w:font="Symbol" w:char="F061"/><T> end</T>'),
    ("no-break hyphen only", "ins", "<w:noBreakHyphen/>"),
]


def paragraph_xml(index, kind, content):
    tag = "w:t" if kind == "ins" else "w:delText"
    content = content.replace("<T>", f'<{tag} xml:space="preserve">').replace("</T>", f"</{tag}>")
    return (
        f"<w:p {nsdecls('w')}>"
        f'<w:r><w:t xml:space="preserve">Case {index} before </w:t></w:r>'
        f'<w:{kind} w:id="{100 + index}" {AUTHOR}><w:r>{content}</w:r></w:{kind}>'
        '<w:r><w:t xml:space="preserve"> after.</w:t></w:r>'
        "</w:p>"
    )


def build(out_path):
    doc = Document()
    # Compatibility mode 15, as make-test-manuscript.py writes it.
    for cs in doc.settings.element.find(qn("w:compat")).findall(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            cs.set(qn("w:val"), "15")
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))
    for i, (_name, kind, content) in enumerate(CASES):
        sect.addprevious(parse_xml(paragraph_xml(i, kind, content)))
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--count"]
    if len(args) != 1:
        sys.exit(__doc__)
    print(build(args[0]))
    if "--count" in sys.argv:
        print(len(CASES))
