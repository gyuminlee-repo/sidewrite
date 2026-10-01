"""Build small docx files by hand (standard library only) for the anonymization tests.

Every name in here is invented. `build(path, ...)` writes a docx whose author
marks are given by keyword; the defaults are a "dirty" document that carries
every category the anonymizer knows. `clean_marks()` returns the keywords of a
document that has nothing left to anonymize. `EXPECTED_DIRTY` is the number of
marks per category in the dirty document, counted by construction.
"""

import struct
import zipfile
import zlib
from xml.sax.saxutils import escape, quoteattr

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"

FIXED_DATE = "1970-01-01T00:00:00Z"

DIRTY = dict(
    creator="Test Author",
    last_modified_by="Second Person",
    company="Example Lab",
    manager="Lab Head",
    comment_author="Second Person",
    comment_initials="SP",
    change_author="Test Author",
    note_change_author="Test Author",
    people=("Test Author", "Second Person"),
    presence=True,
    custom_properties=1,
    date="2026-01-02T03:04:00Z",
    wide=True,
)

# Marks per category in the dirty document.
EXPECTED_DIRTY = dict(
    names=2,  # creator, lastModifiedBy
    org=3,  # Company, Manager, HyperlinkBase
    custom_props=1,
    comment_authors=2,  # author, initials
    people=4,  # two persons, two presenceInfo
    change_authors=4,  # document: ins, del, rPrChange; footnotes: ins
    dates=6,  # those four, the comment, commentsExtensible dateUtc
    core_dates=4,  # created, modified, lastPrinted, revision
    core_text=4,  # keywords, description, subject, category
    title=1,
    rsids=7,  # document: three on the paragraph, one on a run; settings: rsidRoot and two rsid
    para_ids=2,  # paraId and textId on the body paragraph (the comment's paragraph is not counted)
    doc_vars=2,
    attached_template=2,  # the element and its relationship
    doc_ids=2,
    custom_xml=1,  # the cover-page item; the bibliography item is not counted
    thumbnail=1,
    image_meta=2,  # one JPEG with EXIF and a comment, one PNG with a text chunk
    hidden_text=1,
    protection=1,
    image_uninspected=1,  # a TIFF
    embedded=1,
)

BODY = "The culture grew on methane and reached a plateau."
# Text that looks like attributes: an edit of attributes must never touch it.
TRAP = 'The attributes w:rsidR="00AB12CD" w:date="2026-01-02T03:04:00Z" w:author="Somebody" are quoted here.'


def clean_marks(author="Anonymous", initials="A"):
    return dict(
        creator=author,
        last_modified_by=author,
        company="",
        manager="",
        comment_author=author,
        comment_initials=initials,
        change_author=author,
        note_change_author=author,
        people=(author,),
        presence=False,
        custom_properties=0,
        date=FIXED_DATE,
        wide=False,
    )


def _ns(prefix="w"):
    return (
        'xmlns:%s="%s" xmlns:w14="%s" xmlns:w15="%s" xmlns:mc="%s" xmlns:r="%s" mc:Ignorable="w14 w15"'
        % (prefix, W, W14, W15, MC, R)
    )


# ---------------------------------------------------------------- pictures
def _seg(marker, payload):
    return b"\xff" + bytes([marker]) + struct.pack(">H", len(payload) + 2) + payload


def exif(orientation=1, order="<"):
    """An Exif APP1 payload with an orientation and an artist ('<' little-endian, '>' big-endian)."""
    artist = b"Test Author\x00"
    ifd = struct.pack(order + "H", 2)
    ifd += struct.pack(order + "HHIHH", 0x0112, 3, 1, orientation, 0)
    ifd += struct.pack(order + "HHII", 0x013B, 2, len(artist), 8 + 2 + 24 + 4)
    ifd += struct.pack(order + "I", 0)
    head = (b"II*\x00" if order == "<" else b"MM\x00*") + struct.pack(order + "I", 8)
    return b"Exif\x00\x00" + head + ifd + artist


SCAN = b"\xff\xda" + struct.pack(">H", 8) + b"\x01\x01\x00\x00\x3f\x00" + b"\x12\x34\xff\x00\x56\xff\xd9"
JFIF = _seg(0xE0, b"JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00")
ICC = _seg(0xE2, b"ICC_PROFILE\x00\x01\x01profile-bytes")


def jpeg(orientation=1, order="<", with_exif=True, comment=b"made by Test Author"):
    """Marker structure of a JPEG (not a decodable picture): JFIF, optional Exif, ICC, optional comment, scan."""
    out = b"\xff\xd8" + JFIF
    if with_exif:
        out += _seg(0xE1, exif(orientation, order))
    out += ICC
    if comment:
        out += _seg(0xFE, comment)
    return out + SCAN


def _chunk(kind, data):
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)


def png(text=True):
    out = b"\x89PNG\r\n\x1a\n" + _chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 0, 0, 0, 0))
    if text:
        out += _chunk(b"tEXt", b"Author\x00Test Author") + _chunk(b"tIME", b"\x07\xea\x01\x02\x03\x04\x05")
    return out + _chunk(b"IDAT", zlib.compress(b"\x00\x00")) + _chunk(b"IEND", b"")


# ---------------------------------------------------------------- parts
def document_xml(change_author, body=BODY, extra_ins=False, prefix="w", quote='"', date=DIRTY["date"], wide=True):
    p = prefix
    a = quote + escape(change_author, {'"': "&quot;"}) + quote
    more = (
        '<{p}:ins {p}:id="9" {p}:author={a} {p}:date="{d}"><{p}:r><{p}:t>!</{p}:t></{p}:r></{p}:ins>'
        if extra_ins
        else ""
    )
    rsid_p = ' {p}:rsidR="00AB12CD" {p}:rsidRDefault="00AB12CD" {p}:rsidP="00EF3456" w14:textId="77777777"' if wide else ""
    rsid_r = ' {p}:rsidRPr="00EF3456"' if wide else ""
    # In every document, so that a clean one has the same text as a dirty one.
    hidden = "<{p}:r><{p}:rPr><{p}:vanish/></{p}:rPr><{p}:t>a hidden note</{p}:t></{p}:r>"
    trap = "<{p}:r><{p}:t>" + escape(TRAP) + "</{p}:t></{p}:r>"
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        "<{p}:document {ns}><{p}:body>"
        '<{p}:p w14:paraId="0A1B2C3D"' + rsid_p + "><{p}:r" + rsid_r + '><{p}:t xml:space="preserve">{body} </{p}:t></{p}:r>'
        '<{p}:ins {p}:id="1" {p}:author={a} {p}:date="{d}"><{p}:r><{p}:t>Added words.</{p}:t></{p}:r></{p}:ins>'
        '<{p}:del {p}:id="2" {p}:author={a} {p}:date="{d}"><{p}:r><{p}:delText>Removed words.</{p}:delText></{p}:r></{p}:del>'
        "<{p}:r><{p}:rPr><{p}:i/>"
        '<{p}:rPrChange {p}:id="3" {p}:author={a} {p}:date="{d}"><{p}:rPr/></{p}:rPrChange>'
        "</{p}:rPr><{p}:t>Escherichia coli</{p}:t></{p}:r>"
        '<{p}:commentRangeStart {p}:id="0"/><{p}:r><{p}:t> grew.</{p}:t></{p}:r><{p}:commentRangeEnd {p}:id="0"/>'
        '<{p}:r><{p}:commentReference {p}:id="0"/></{p}:r>' + hidden + trap + more + "</{p}:p>"
        "<{p}:sectPr/></{p}:body></{p}:document>"
    )
    return xml.format(p=p, ns=_ns(p), a=a, body=escape(body), d=date)


HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _rels(*rows):
    return HEAD + '<Relationships xmlns="%s">%s</Relationships>' % (
        REL,
        "".join(
            '<Relationship Id="rId%d" Type="%s" Target=%s%s/>'
            % (i + 1, kind, quoteattr(target), ' TargetMode="External"' if external else "")
            for i, (kind, target, external) in enumerate(rows)
        ),
    )


def _wide_parts(files, m):
    """The parts and references only a dirty (wide) document has."""
    office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    files["_rels/.rels"] = _rels(
        (office + "officeDocument", "word/document.xml", False),
        ("http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail", "docProps/thumbnail.jpeg", False),
        ("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml", False),
    )
    files["docProps/thumbnail.jpeg"] = jpeg(with_exif=False, comment=b"")
    files["word/_rels/document.xml.rels"] = _rels(
        (office + "settings", "settings.xml", False),
        (office + "comments", "comments.xml", False),
        (office + "customXml", "../customXml/item1.xml", False),
        (office + "customXml", "../customXml/item2.xml", False),
        (office + "image", "media/image1.jpeg", False),
        (office + "image", "media/image2.png", False),
        (office + "image", "media/image3.tiff", False),
        (office + "oleObject", "embeddings/oleObject1.bin", False),
        (office + "hyperlink", "https://journal.example/guide", True),
    )
    files["word/settings.xml"] = (
        HEAD + "<w:settings %s>"
        '<w:attachedTemplate r:id="rId1"/>'
        '<w:documentProtection w:edit="readOnly" w:enforcement="0"/>'
        '<w:rsids><w:rsidRoot w:val="00AB12CD"/><w:rsid w:val="00AB12CD"/><w:rsid w:val="00EF3456"/></w:rsids>'
        '<w:docVars><w:docVar w:name="RefManagerUser" w:val="test.author"/><w:docVar w:name="LabCode" w:val="EX"/></w:docVars>'
        '<w14:docId w14:val="1A2B3C4D"/><w15:docId w15:val="{00000000-1111-2222-3333-444444444444}"/>'
        '<w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/></w:compat>'
        "</w:settings>" % _ns()
    )
    files["word/_rels/settings.xml.rels"] = _rels(
        (office + "attachedTemplate", "file:///C:\\Users\\test.author\\Templates\\Lab.dotm", True)
    )
    files["word/commentsExtended.xml"] = (
        HEAD + '<w15:commentsEx %s><w15:commentEx w15:paraId="5E6F7A8B" w15:done="0"/></w15:commentsEx>' % _ns()
    )
    files["word/commentsExtensible.xml"] = (
        HEAD + '<w16cex:commentsExtensible xmlns:w16cex="%s">'
        '<w16cex:commentExtensible w16cex:durableId="1234ABCD" w16cex:dateUtc="%s"/>'
        "</w16cex:commentsExtensible>" % (W16CEX, m["date"])
    )
    files["customXml/item1.xml"] = (
        HEAD + '<CoverPageProperties xmlns="http://schemas.microsoft.com/office/2006/coverPageProps">'
        "<CompanyEmail>test.author@example.org</CompanyEmail></CoverPageProperties>"
    )
    files["customXml/item2.xml"] = (
        HEAD + '<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle=""/>'
    )
    for n in (1, 2):
        files["customXml/itemProps%d.xml" % n] = (
            HEAD + '<ds:datastoreItem ds:itemID="{0000000%d-AAAA-BBBB-CCCC-DDDDDDDDDDDD}" '
            'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>' % n
        )
        files["customXml/_rels/item%d.xml.rels" % n] = _rels((office + "customXmlProps", "itemProps%d.xml" % n, False))
    files["word/media/image1.jpeg"] = jpeg()
    files["word/media/image2.png"] = png()
    files["word/media/image3.tiff"] = b"II*\x00\x08\x00\x00\x00\x00\x00"
    files["word/embeddings/oleObject1.bin"] = b"\xd0\xcf\x11\xe0 an embedded object"


def build(path, body=BODY, extra_ins=False, prefix="w", quote='"', parts=None, extra=None, **marks):
    """Write a docx. `parts`: names of optional parts to leave out, e.g. {"word/people.xml"}.
    `extra`: {name: bytes} of further zip members."""
    m = dict(DIRTY)
    m.update(marks)
    wide = m["wide"]
    skip = set(parts or ())
    files = {}
    overrides = "".join(
        '<Override PartName="/customXml/itemProps%d.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' % n
        for n in ((1, 2) if wide else ())
    )
    files["[Content_Types].xml"] = (
        HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + overrides
        + "</Types>"
    )
    files["_rels/.rels"] = _rels(
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "word/document.xml", False)
    )
    core_wide = (
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-02-01T00:00:00Z</dcterms:modified>'
        "<cp:lastPrinted>2026-01-15T00:00:00Z</cp:lastPrinted><cp:revision>42</cp:revision>"
        "<cp:keywords>methane; invented</cp:keywords><dc:description>Draft for the Example Lab</dc:description>"
        "<dc:subject>Internal</dc:subject><cp:category>Manuscripts</cp:category>"
        "<dc:title>An invented manuscript</dc:title>"
        if wide
        else '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created><cp:revision>1</cp:revision>'
        "<cp:keywords></cp:keywords>" % FIXED_DATE
    )
    files["docProps/core.xml"] = (
        HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>%s</dc:creator><cp:lastModifiedBy>%s</cp:lastModifiedBy>%s"
        "</cp:coreProperties>" % (escape(m["creator"]), escape(m["last_modified_by"]), core_wide)
    )
    files["docProps/app.xml"] = (
        HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        "<Template>Normal.dotm</Template><TotalTime>135</TotalTime><Manager>%s</Manager><Company>%s</Company>"
        "<Pages>1</Pages>%s<AppVersion>16.0000</AppVersion>"
        "</Properties>"
        % (
            escape(m["manager"]),
            escape(m["company"]),
            "<HyperlinkBase>C:\\Users\\test.author\\Papers</HyperlinkBase>" if wide else "",
        )
    )
    props = "".join(
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="Project code %d">'
        "<vt:lpwstr>EX-%d</vt:lpwstr></property>" % (i + 2, i, i)
        for i in range(m["custom_properties"])
    )
    files["docProps/custom.xml"] = (
        HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">%s</Properties>' % props
    )
    files["word/document.xml"] = document_xml(m["change_author"], body, extra_ins, prefix, quote, m["date"], wide)
    files["word/comments.xml"] = (
        HEAD + "<w:comments %s>"
        '<w:comment w:id="0" w:author=%s w:date="%s" w:initials=%s>'
        "<w:p%s><w:r><w:t>Please check this sentence.</w:t></w:r></w:p></w:comment></w:comments>"
        % (
            _ns(),
            quoteattr(m["comment_author"]),
            m["date"],
            quoteattr(m["comment_initials"]),
            ' w14:paraId="5E6F7A8B"' if wide else "",
        )
    )
    people = "".join(
        "<w15:person w15:author=%s>%s</w15:person>"
        % (
            quoteattr(name),
            '<w15:presenceInfo w15:providerId="AD" w15:userId="S::%d@example.org::0000"/>' % i if m["presence"] else "",
        )
        for i, name in enumerate(m["people"])
    )
    files["word/people.xml"] = HEAD + "<w15:people %s>%s</w15:people>" % (_ns(), people)
    files["word/footnotes.xml"] = (
        HEAD + '<w:footnotes %s><w:footnote w:id="1"><w:p><w:r><w:t>A note.</w:t></w:r>'
        '<w:ins w:id="7" w:author=%s w:date="%s"><w:r><w:t> More.</w:t></w:r></w:ins>'
        "</w:p></w:footnote></w:footnotes>" % (_ns(), quoteattr(m["note_change_author"]), m["date"])
    )
    files["word/styles.xml"] = HEAD + '<w:styles xmlns:w="%s"/>' % W
    if wide:
        _wide_parts(files, m)
    files.update(extra or {})
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            if name not in skip:
                z.writestr(name, data)
    return path
