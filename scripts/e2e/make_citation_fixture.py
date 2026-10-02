#!/usr/bin/env python3
"""Fake manuscript with reference-manager fields, for the `cite` acceptance stage.

Every sentence, author, title and identifier is invented. The document holds
the fields the Zotero and EndNote Word plugins write, in the OOXML form they
have in a saved .docx, so that tracked edits next to them can be judged by
tools/citation_fields_check.py.

Where each shape comes from:

  Zotero citation      instruction " ADDIN ZOTERO_ITEM CSL_CITATION {json} "
                       zotero-word-for-windows-integration field.cpp
                       (FIELD_PREFIXES " ADDIN ZOTERO_", setCode "%s%s "),
                       zotero integration.js (`ITEM CSL_CITATION ${code}`,
                       the keys of the csl-citation.json schema)
  Zotero bibliography  " ADDIN ZOTERO_BIBL {json} CSL_BIBLIOGRAPHY "
                       integration.js (`BIBL ${code} CSL_BIBLIOGRAPHY`); the
                       keys uncited/omitted/custom of {json}: own knowledge
  ZOTERO_PREF_n        custom document properties, 255 characters each, 1-based
                       zoteroWinWordIntegration.h (PREFS_PROPERTY,
                       MAX_PROPERTY_LENGTH), document.cpp setProperty;
                       the <data ...> text: integration.js DocumentData.serialize
  w:fldData            a child of w:fldChar whose content the application
                       defines: ISO/IEC 29500-1 17.16 (fldData, parent fldChar)
  EndNote citation     " ADDIN EN.CITE <EndNote><Cite>...</Cite></EndNote>" and
                       the nested " ADDIN EN.CITE.DATA " field: secondary
                       sources only (scripts that read such documents). The
                       tags inside <record>, the content of w:fldData (here the
                       base64 of the same XML) and that both begin markers
                       carry it: own knowledge, NOT confirmed from a source
  EndNote list         " ADDIN EN.REFLIST ": the same secondary sources
  EN.* docVars         EN.InstantFormat, EN.Layout, EN.Libraries in
                       settings.xml: the names from a secondary source, the
                       values own knowledge, NOT confirmed
  how Word splits an instruction over runs, complex REF field, w:hyperlink:
                       ISO/IEC 29500-1 17.16; where the splits fall is chosen here

Usage: python3 scripts/e2e/make_citation_fixture.py [output.docx]
       python3 scripts/e2e/make_citation_fixture.py --cases   (the texts the
           stage edits, as JSON; standard library only)
       python3 scripts/e2e/make_citation_fixture.py --dump x.docx   (paragraph
           texts as JSON, for tests; standard library only)
       python3 scripts/e2e/make_citation_fixture.py --dump-reviewed x.docx   (the
           same with the fields as Word's getReviewedText shows them)

Writing the fixture needs python-docx (a development tool; the fixture is
committed as tests/fixtures/citation_manuscript.docx).
"""
import base64
import json
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

FIXTURE = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "citation_manuscript.docx"
FIXED_TIME = (2026, 1, 1, 0, 0, 0)

SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
LIBRARY = "http://zotero.org/users/local/Fx7Invnt/items/"

# Invented references: key -> (item id, family, given, year, title).
REFS = {
    "arden": (101, "Arden", "Q.", "2019", "Shaken flasks as a first vessel"),
    "belloc": (102, "Belloc", "M.", "2020", "Slow growth in narrow flasks"),
    "corvane": (103, "Corvane", "I.", "2021", "The medium, not the flask"),
    "dellmar": (104, "Dellmar", "S.", "2018", "Fill volume and oxygen transfer"),
    "eskridge": (105, "Eskridge", "P.", "2022", "Baffles and sampling"),
    "fenwyck": (106, "Fenwyck", "A.", "2023", "Foam in invented cultures"),
}
EN_REFS = {
    "garrowby": (7, "Garrowby", "T.", "2017", "Foam plugs in invented flasks"),
    "hallivane": (8, "Hallivane", "R.", "2016", "A reader calibrated against dry weight"),
}
JOURNAL = "Journal of Invented Cultures"

ZOTERO_DATA = (
    '<data data-version="3" zotero-version="7.0.11"><session id="Fx7Sess1"/>'
    '<style id="http://www.zotero.org/styles/apa" locale="en-US" hasBibliography="1" '
    'bibliographyStyleHasBeenSet="1"/><prefs><pref name="fieldType" value="Field"/>'
    '<pref name="automaticJournalAbbreviations" value="true"/></prefs></data>'
)
MAX_PROPERTY_LENGTH = 255

EN_DOCVARS = {
    "EN.InstantFormat": "<ENInstantFormat><Enabled>0</Enabled><ScanUnformatted>1</ScanUnformatted>"
    "<ScanChanges>1</ScanChanges><Suspended>0</Suspended></ENInstantFormat>",
    "EN.Layout": "<ENLayout><Style>APA 7th</Style><LeftDelim>{</LeftDelim><RightDelim>}</RightDelim>"
    "<FontName>Calibri</FontName><FontSize>11</FontSize><ReflistTitle></ReflistTitle>"
    "<StartingRefnum>1</StartingRefnum><FirstLineIndent>0</FirstLineIndent>"
    "<HangingIndent>720</HangingIndent><LineSpacing>0</LineSpacing><SpaceAfter>0</SpaceAfter>"
    "</ENLayout>",
    "EN.Libraries": "<Libraries><item db-id=\"x0invented0db0id0fixture00\">invented-library"
    "<record-ids><item>7</item><item>8</item></record-ids></item></Libraries>",
}

BOOKMARK = "_Ref100000001"
LINK = "https://example.invalid/invented-protocol"

# What the stage does without a model, one tracked edit each (run-acceptance
# mainCite). Every find is in the document exactly once.
# expect: what the tool has to answer (judge.mjs citeCaseRow).
#   "applied"  the edit goes in and no field is touched
#   "refused"  refused as FieldNotEditable, Word unchanged: the range lies in
#              the text of a field, cuts through one, or would replace one
#   "removed"  the edit goes in and exactly `fields` field(s) lie whole in
#              the tracked deletion (begin through end)
# The cases that must be refused share a paragraph with a later edit of the
# run (the two in the baffles paragraph with the turn of R1.3, the EndNote
# one with the turn of R1.2); when one of them is applied against the
# expectation, its own row fails first and names the field.
CASES = [
    {
        "name": "next-to",
        "what": "a phrase next to a Zotero citation, in its sentence",
        "find": "are cheap to run",
        "replace": "are inexpensive to run",
        "expect": "applied",
    },
    {
        "name": "spans",
        "what": "a replacement whose find holds the whole visible text of a Zotero citation",
        "find": "slow growth (Belloc, 2020) in narrow flasks",
        "replace": "reduced growth (Belloc, 2020) in slim flasks",
        "expect": "refused",
    },
    {
        "name": "straddles",
        "what": "a phrase that starts before a Zotero citation and ends inside its text",
        "find": "that reading (Corvane,",
        "replace": "this view (Corvane;",
        "expect": "refused",
    },
    {
        "name": "cite-text-zotero",
        "what": "the visible text of a Zotero citation alone, deleted",
        "find": "(Arden, 2019; Eskridge, 2022)",
        "replace": "",
        "expect": "refused",
    },
    {
        "name": "cite-text-endnote",
        "what": "the visible text of an EndNote citation alone, deleted",
        "find": "(Hallivane, 2016)",
        "replace": "",
        "expect": "refused",
    },
    {
        "name": "cite-text-replace",
        "what": "the visible text of a Zotero citation replaced by other text",
        "find": "(Dellmar, 2018)",
        "replace": "(Dellmar and others, 2018)",
        "expect": "refused",
    },
    {
        "name": "delete-sentence",
        "what": "a whole sentence with an EndNote citation, deleted as one edit",
        "find": "The plugs were autoclaved twice before use (Garrowby, 2017).",
        "replace": "",
        "expect": "removed",
        "fields": 1,
    },
    {
        "name": "delete-sentence-zotero",
        "what": "a whole sentence with a Zotero citation, deleted as one edit",
        "find": "The invented strain grows fastest when a tenth of the flask is filled (Arden, 2019).",
        "replace": "",
        "expect": "removed",
        "fields": 1,
    },
    {
        "name": "formatted",
        "what": "a phrase next to a field in the paragraph with italic and superscript runs",
        "find": "as counted under the microscope",
        "replace": "as counted by microscopy",
        "expect": "applied",
    },
]
# The paragraph the whole-paragraph rewrite is tried on, and the words changed.
REWRITE = {
    "name": "rewrite",
    "what": "office_tracked_rewrite on a paragraph with two citation fields",
    "paragraph_has": "Foam formed in every invented run",
    "changes": [["held down", "suppressed"], ["did not change", "did not alter"]],
}

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def run(text, fmt=""):
    rpr = {"i": "<w:rPr><w:i/></w:rPr>", "sup": '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'}
    return f'<w:r>{rpr.get(fmt, "")}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def fld(kind, data=None):
    inner = f'<w:fldData xml:space="preserve">{data}</w:fldData>' if data else ""
    return f'<w:r><w:fldChar w:fldCharType="{kind}">{inner}</w:fldChar></w:r>'


def instr(code, pieces=1):
    """The instruction, over `pieces` runs (Word splits a long one)."""
    size = -(-len(code) // pieces)
    parts = [code[i : i + size] for i in range(0, len(code), size)]
    return "".join(
        f'<w:r><w:instrText xml:space="preserve">{escape(p)}</w:instrText></w:r>' for p in parts
    )


def field(code, result, pieces=1):
    return fld("begin") + instr(code, pieces) + fld("separate") + result + fld("end")


def zotero_code(keys, citation_id):
    text = "(" + "; ".join(f"{REFS[k][1]}, {REFS[k][3]}" for k in keys) + ")"
    items = []
    for k in keys:
        item_id, family, given, year, title = REFS[k]
        items.append(
            {
                "id": item_id,
                "uris": [f"{LIBRARY}INV{item_id:05d}"],
                "itemData": {
                    "id": item_id,
                    "type": "article-journal",
                    "container-title": JOURNAL,
                    "title": title,
                    "author": [{"family": family, "given": given}],
                    "issued": {"date-parts": [[year]]},
                },
            }
        )
    data = {
        "citationID": citation_id,
        "properties": {"formattedCitation": text, "plainCitation": text, "noteIndex": 0},
        "citationItems": items,
        "schema": SCHEMA,
    }
    return " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(data, separators=(",", ":")) + " ", text


def zotero(keys, citation_id, pieces=1):
    code, text = zotero_code(keys, citation_id)
    return field(code, run(text), pieces)


def endnote_xml(key):
    num, family, given, year, title = EN_REFS[key]
    return (
        f"<EndNote><Cite><Author>{family}</Author><Year>{year}</Year><RecNum>{num}</RecNum>"
        f"<DisplayText>({family}, {year})</DisplayText><record><rec-number>{num}</rec-number>"
        f'<foreign-keys><key app="EN" db-id="x0invented0db0id0fixture00" timestamp="1700000000">'
        f'{num}</key></foreign-keys><ref-type name="Journal Article">17</ref-type>'
        f"<contributors><authors><author>{family}, {given}</author></authors></contributors>"
        f"<titles><title>{title}</title><secondary-title>{JOURNAL}</secondary-title></titles>"
        f"<periodical><full-title>{JOURNAL}</full-title></periodical><volume>{num}</volume>"
        f"<dates><year>{year}</year></dates><urls></urls></record></Cite></EndNote>"
    )


def endnote_result(key):
    _num, family, _given, year, _title = EN_REFS[key]
    return (
        '<w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve">'
        f"({family}, {year})</w:t></w:r>"
    )


def endnote_short(key):
    """The record inside the instruction."""
    return field(" ADDIN EN.CITE " + endnote_xml(key), endnote_result(key), pieces=3)


def endnote_long(key):
    """The record in w:fldData, with the nested EN.CITE.DATA field."""
    data = base64.b64encode(endnote_xml(key).encode("utf-8")).decode("ascii")
    return (
        fld("begin", data)
        + instr(" ADDIN EN.CITE ")
        + fld("begin", data)
        + instr(" ADDIN EN.CITE.DATA ")
        + fld("end")
        + fld("separate")
        + endnote_result(key)
        + fld("end")
    )


def para(content, style=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p {W} {R}>{ppr}{content}</w:p>"


def reference(key, table):
    _num, family, given, year, title = table[key]
    return f"{family}, {given} ({year}). {title}. {JOURNAL}."


def body(link_rid):
    z = zotero
    out = [
        para(run("Citation fields in an invented flask study"), "Title"),
        para(
            run(
                "Fake manuscript for software testing. Every sentence, name and reference "
                "is invented."
            )
        ),
        para(run("Introduction"), "Heading1"),
        para(
            run("Shaken flasks remain the usual first vessel for an invented culture ")
            + z(["arden"], "cFx00001")
            + run(" and are cheap to run. Their oxygen supply is limited by the liquid surface.")
        ),
        para(
            run("Earlier invented work reported slow growth ")
            + z(["belloc"], "cFx00002")
            + run(" in narrow flasks at low shaking speed.")
        ),
        para(
            run("A later invented survey disagreed with that reading ")
            + z(["corvane"], "cFx00003", pieces=2)
            + run(" and blamed the medium.")
        ),
        para(
            run("Oxygen transfer falls as the fill volume rises ")
            + z(["dellmar", "eskridge"], "cFx00004", pieces=4)
            + run(". The invented strain grows fastest when a tenth of the flask is filled ")
            + z(["arden"], "cFx00005")
            + run(".")
        ),
        para(
            run("Cells of ")
            + run("Examplea fictiva", "i")
            + run(" reached 10")
            + run("9", "sup")
            + run(" cells per millilitre within a day ")
            + z(["belloc"], "cFx00006")
            + run(", as counted under the microscope.")
        ),
        para(
            run("Foam formed in every invented run ")
            + z(["fenwyck"], "cFx00007")
            + run(". It was held down with a silicone agent, which did not change the growth rate ")
            + z(["corvane"], "cFx00008", pieces=3)
            + run(".")
        ),
        para(
            run(
                "It has been said by several invented authors that it is the case that baffles "
                "raise oxygen transfer "
            )
            + z(["dellmar"], "cFx00009")
            + run(", and it has also been said that they raise foaming ")
            + z(["fenwyck"], "cFx00010")
            + run(", and in addition it has been said that they make sampling harder ")
            + z(["arden", "eskridge"], "cFx00011", pieces=5)
            + run(".")
        ),
        para(run("Methods"), "Heading1"),
        para(
            run("Flasks were closed with foam plugs. The plugs were autoclaved twice before use ")
            + endnote_short("garrowby")
            + run(". Cultures were shaken at a fixed speed.")
        ),
        para(
            run(
                "Optical density was read every hour. The reader had been calibrated against "
                "dry weight in an earlier invented study "
            )
            + endnote_long("hallivane")
            + run(". Samples were kept on ice.")
        ),
        para(
            run("The plateau is described in the section ")
            + field(f" REF {BOOKMARK} \\h ", run("Growth plateau"))
            + run(", and the invented protocol is kept at ")
            + f'<w:hyperlink r:id="{link_rid}" w:history="1">'
            + '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>'
            + "<w:t>the protocol page</w:t></w:r></w:hyperlink>"
            + run(".")
        ),
        para(
            f'<w:bookmarkStart w:id="1" w:name="{BOOKMARK}"/>'
            + run("Growth plateau")
            + '<w:bookmarkEnd w:id="1"/>',
            "Heading1",
        ),
        para(run("Every invented culture stopped growing before the second day.")),
        para(run("References"), "Heading1"),
    ]
    # The Zotero bibliography: one field over all its paragraphs.
    keys = sorted(REFS)
    bibl = ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '
    for i, k in enumerate(keys):
        head = fld("begin") + instr(bibl) + fld("separate") if i == 0 else ""
        tail = fld("end") if i == len(keys) - 1 else ""
        out.append(para(head + run(reference(k, REFS)) + tail, "Bibliography"))
    # The EndNote list, the same way.
    keys = sorted(EN_REFS)
    for i, k in enumerate(keys):
        head = fld("begin") + instr(" ADDIN EN.REFLIST ") + fld("separate") if i == 0 else ""
        tail = fld("end") if i == len(keys) - 1 else ""
        out.append(para(head + run(reference(k, EN_REFS)) + tail, "EndNoteBibliography"))
    return out


def custom_xml():
    chunks = [
        ZOTERO_DATA[i : i + MAX_PROPERTY_LENGTH]
        for i in range(0, len(ZOTERO_DATA), MAX_PROPERTY_LENGTH)
    ]
    props = "".join(
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" '
        f'pid="{n + 2}" name="ZOTERO_PREF_{n + 1}"><vt:lpwstr>{escape(c)}</vt:lpwstr></property>'
        for n, c in enumerate(chunks)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f"{props}</Properties>"
    )


def add_custom_part(path):
    """docProps/custom.xml with its content type and relationship; fixed times."""
    ct = (
        '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.custom-properties+xml"/>'
    )
    rel = (
        '<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/'
        '2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
    )
    with zipfile.ZipFile(path) as z:
        parts = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts + [("docProps/custom.xml", custom_xml().encode("utf-8"))]:
            if name == "[Content_Types].xml":
                data = data.replace(b"</Types>", ct.encode() + b"</Types>")
            elif name == "_rels/.rels":
                data = data.replace(b"</Relationships>", rel.encode() + b"</Relationships>")
            z.writestr(zipfile.ZipInfo(name, FIXED_TIME), data, zipfile.ZIP_DEFLATED)


def build(out_path=FIXTURE):
    import datetime

    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn

    doc = Document()
    compat = doc.settings.element.find(qn("w:compat"))
    for cs in compat.findall(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            cs.set(qn("w:val"), "15")
    doc_vars = "".join(
        f"<w:docVar w:name={quoteattr(k)} w:val={quoteattr(v)}/>" for k, v in EN_DOCVARS.items()
    )
    compat.addnext(parse_xml(f"<w:docVars {W}>{doc_vars}</w:docVars>"))
    for name in ("Bibliography", "EndNote Bibliography"):
        if name not in [s.name for s in doc.styles]:
            doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    if "Hyperlink" not in [s.name for s in doc.styles]:
        doc.styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER)
    rid = doc.part.relate_to(LINK, RT.HYPERLINK, is_external=True)
    sect = doc.element.body.find(qn("w:sectPr"))
    for xml in body(rid):
        sect.addprevious(parse_xml(xml))
    fixed = datetime.datetime(*FIXED_TIME)
    props = doc.core_properties
    props.author = "Fixture"
    props.last_modified_by = "Fixture"
    props.created = fixed
    props.modified = fixed
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    add_custom_part(out_path)
    return out_path


def dump(path):
    """The text of every body paragraph (w:t only, as Word shows it with the
    field codes hidden), for tests. Standard library only."""
    from xml.etree import ElementTree as ET

    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return ["".join(t.text or "" for t in p.iter(ns + "t")) for p in root.iter(ns + "p")]


def dump_reviewed(path):
    """The text of every body paragraph as Range.getReviewedText gave it in the
    cite run of 2026-10-02: a field as U+0013 instruction U+0014 result U+0015
    (a field over several paragraphs with its U+0015 in the last one), a
    hyperlink as U+0013 HYPERLINK "url" U+0014 text U+0015. The U+0001 Word
    puts at the end of some instructions is left out. For tests."""
    from xml.etree import ElementTree as ET

    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    mark = {"begin": "\x13", "separate": "\x14", "end": "\x15"}

    def walk(el):
        tag = el.tag
        if tag == ns + "fldChar":
            return mark[el.get(ns + "fldCharType")]
        if tag in (ns + "t", ns + "instrText"):
            return el.text or ""
        inner = "".join(walk(child) for child in el)
        if tag == ns + "hyperlink":
            return f'\x13HYPERLINK "{LINK}"\x14{inner}\x15'
        return inner

    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return [walk(p) for p in root.iter(ns + "p")]


if __name__ == "__main__":
    if "--dump-reviewed" in sys.argv:
        print(json.dumps(dump_reviewed(sys.argv[sys.argv.index("--dump-reviewed") + 1])))
    elif "--cases" in sys.argv:
        print(json.dumps({"cases": CASES, "rewrite": REWRITE}))
    elif "--dump" in sys.argv:
        print(json.dumps(dump(sys.argv[sys.argv.index("--dump") + 1])))
    elif len(sys.argv) > 2:
        sys.exit(__doc__)
    else:
        print(build(*sys.argv[1:]))
