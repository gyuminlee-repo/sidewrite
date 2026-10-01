#!/usr/bin/env python3
"""Build a fake response letter of an earlier round, for tests. Standard library only.

  python3 scripts/make-test-response.py --variant plain tests/fixtures/prior_response.docx

The letter answers tests/fixtures/fake_review.txt (one reviewer, five comments)
about an invented manuscript. Everything in it is made up. The answers are
C00000, location lines and quotes 0070C0. Each variant writes the same letter
in another way Word files do, so the reader (tools/response_read.py) is tested
on each:

  plain      colours as direct w:color on the runs
  table      one two-cell table row per comment (comment | answer)
  tracked    one answer holds w:ins, w:del and w:moveFrom runs
  styled     colours from a character style (one basedOn step), a paragraph
             style and a theme colour; one run with w:color "auto"
  eastasian  translation paragraphs in Hangul with East Asian font settings
  highlight  a highlighted run and a shaded run inside coloured text

The same variant always gives the same bytes.
"""

import argparse
import sys
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import response_docx as rd  # noqa: E402  (deterministic zip writer and package parts)

VARIANTS = ("plain", "table", "tracked", "styled", "eastasian", "highlight")
BODY, REVISED, GLOSS = "C00000", "0070C0", "7F7F7F"
W_NS = rd.W
STYLES_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
STYLES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
TRACK = 'w:author="Author" w:date="2000-01-01T00:00:00Z"'

INTRO = ("Fake response letter for software testing. It answers tests/fixtures/fake_review.txt about an "
         "invented manuscript. No real study, journal, or person is involved.")
TITLE = "Response to Reviewers"
REVIEWER = "Reviewer #1"
# (comment, answer, location line or None, quote or None)
COMMENTS = [
    ("Comment 1: Please state the unit of the incubation temperature as degrees Celsius.",
     "We now give the unit as degrees Celsius.",
     "[Revised in 2. Materials and Methods, line 9]",
     "The incubation temperature was held at 30 °C."),
    ("Comment 2: Please call the extract a clarified lysate.",
     "We replaced the term in that paragraph.",
     "[Revised in 2. Materials and Methods, lines 6-7]",
     "A clarified lysate was prepared by sonication."),
    ("Comment 3: Please state how growth was followed.",
     "We state how growth was followed.",
     "[Revised in 3. Results, lines 12-13]",
     "Growth was followed by optical density at 600 nm (Fig. 2)."),
    ("Comment 4: Please delete the sentence about where the flasks came from.",
     "We deleted the sentence.",
     "[Removed from 1. Introduction]",
     None),
    ("Comment 5: Please name the antifoam agent that was used.",
     "No antifoam agent was used in this study.",
     "[No change made]",
     None),
]
GLOSSES = {0: "[T] 단위를 추가했다.", 1: "[T] 용어를 바꾸었다."}


def run(text, color=None, rstyle=None, theme=None, highlight=None, shade=None, east_asian=False, tag="w:t"):
    props = []
    if rstyle:
        props.append('<w:rStyle w:val="%s"/>' % rstyle)
    if east_asian:
        props.append('<w:rFonts w:eastAsia="Malgun Gothic" w:hint="eastAsia"/>')
    if color:
        props.append('<w:color w:val="%s"%s/>' % (color, ' w:themeColor="%s"' % theme if theme else ""))
    if highlight:
        props.append('<w:highlight w:val="%s"/>' % highlight)
    if shade:
        props.append('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade)
    if east_asian:
        props.append('<w:lang w:eastAsia="ko-KR"/>')
    rpr = "<w:rPr>%s</w:rPr>" % "".join(props) if props else ""
    return '<w:r>%s<%s xml:space="preserve">%s</%s></w:r>' % (rpr, tag, escape(text), tag)


def para(runs, pstyle=None):
    ppr = '<w:pPr><w:pStyle w:val="%s"/></w:pPr>' % pstyle if pstyle else ""
    return "<w:p>%s%s</w:p>" % (ppr, runs if isinstance(runs, str) else "".join(runs))


def answer(variant, i, text):
    """Paragraphs of the answer to comment i (0-based)."""
    if variant == "styled":
        return [para(run(text, rstyle="ResponseChar"))]
    if variant == "tracked" and i == 0:
        head, tail = "We ", "the unit as degrees Celsius."
        return [para([
            run(head, BODY),
            '<w:ins w:id="1" %s>%s</w:ins>' % (TRACK, run("added ", BODY)),
            '<w:del w:id="2" %s>%s</w:del>' % (TRACK, run("removed ", BODY, tag="w:delText")),
            '<w:moveFrom w:id="3" %s>%s</w:moveFrom>' % (TRACK, run("moved here ", BODY)),
            run(tail, BODY),
        ])]
    if variant == "eastasian" and i == 2:
        # Latin and Hangul runs in one paragraph.
        return [para([run("We state how growth was followed ", BODY), run("(광학 밀도).", BODY, east_asian=True)])]
    if variant == "highlight" and i == 1:
        return [para([run("We replaced ", BODY), run("the term", BODY, shade="FFFF00"), run(" in that paragraph.", BODY)])]
    out = [para(run(text, BODY))]
    if variant == "eastasian" and i in GLOSSES:
        out.append(para(run(GLOSSES[i], GLOSS, east_asian=True)))
    return out


def location(variant, text):
    if text == "[No change made]":
        # styled: a direct "auto" on a run whose character style has a colour.
        return para(run(text, "auto", rstyle="ResponseChar") if variant == "styled" else run(text, BODY))
    if variant == "styled":
        return para(run(text), pstyle="LocationLine")
    return para(run(text, REVISED))


def quote(variant, i, text):
    if variant == "styled" and i == 0:
        return para(run(text, REVISED, theme="accent1"))
    if variant == "highlight" and i == 2:
        before, ref, after = text.partition("Fig. 2")
        return para([run(before, REVISED), run(ref, REVISED, highlight="yellow"), run(after, REVISED)])
    return para(run(text, REVISED))


def block(variant, i):
    """(comment paragraphs, answer paragraphs) of comment i."""
    comment, ans, loc, q = COMMENTS[i]
    right = answer(variant, i, ans)
    if loc:
        right.append(location(variant, loc))
    if q:
        right.append(quote(variant, i, q))
    return [para(run(comment))], right


def cell(paragraphs):
    return '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>%s</w:tc>' % "".join(paragraphs)


def body_xml(variant):
    title = run(TITLE, "auto") if variant == "styled" else run(TITLE)
    out = [para(run(INTRO)), para(title), para(run(REVIEWER))]
    if variant == "table":
        rows = []
        for i in range(len(COMMENTS)):
            left, right = block(variant, i)
            if i == len(COMMENTS) - 1:
                right.append("<w:p/>")  # an empty paragraph inside a cell
            rows.append("<w:tr>%s%s</w:tr>" % (cell(left), cell(right)))
        out.append('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/>'
                   '<w:gridCol w:w="4500"/></w:tblGrid>%s</w:tbl>' % "".join(rows))
        out.append("<w:p/>")  # Word wants a paragraph after a table
    else:
        for i in range(len(COMMENTS)):
            left, right = block(variant, i)
            out.extend(left + right)
    return "".join(out)


def document_xml(variant):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="%s"><w:body>%s%s</w:body></w:document>' % (W_NS, body_xml(variant), rd.SECT))


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="%s">'
          '<w:style w:type="character" w:styleId="ResponseBase"><w:name w:val="Response Base"/>'
          '<w:rPr><w:color w:val="%s"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="ResponseChar"><w:name w:val="Response Char"/>'
          '<w:basedOn w:val="ResponseBase"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="LocationLine"><w:name w:val="Location Line"/>'
          '<w:rPr><w:color w:val="%s"/></w:rPr></w:style>'
          '</w:styles>' % (W_NS, BODY, REVISED))


def parts(variant):
    if variant not in VARIANTS:
        raise ValueError("unknown variant %r (one of %s)" % (variant, ", ".join(VARIANTS)))
    out = [("[Content_Types].xml", rd.CONTENT_TYPES), ("_rels/.rels", rd.RELS), ("word/document.xml", document_xml(variant))]
    if variant == "styled":
        override = '<Override PartName="/word/styles.xml" ContentType="%s"/>' % STYLES_TYPE
        out[0] = ("[Content_Types].xml", rd.CONTENT_TYPES.replace("</Types>", override + "</Types>"))
        out.append(("word/_rels/document.xml.rels",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="%s" Target="styles.xml"/></Relationships>' % STYLES_REL))
        out.append(("word/styles.xml", STYLES))
    return out


def build(path, variant="plain"):
    rd.write_zip(str(path), parts(variant))
    return path


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("out", nargs="?", default="prior_response.docx")
    ap.add_argument("--variant", choices=VARIANTS, default="plain")
    a = ap.parse_args(argv)
    build(a.out, a.variant)
    print("wrote %s (%s)" % (a.out, a.variant))
    return 0


if __name__ == "__main__":
    sys.exit(main())
