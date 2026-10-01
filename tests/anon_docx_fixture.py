"""Build small docx files by hand (standard library only) for the anonymization tests.

Every name in here is invented. `build(path, ...)` writes a docx whose author
marks are given by keyword; the defaults are a "dirty" document. `clean_marks()`
returns the keywords of a document that has nothing left to anonymize.
"""

import zipfile
from xml.sax.saxutils import escape, quoteattr

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

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
)

BODY = "The culture grew on methane and reached a plateau."


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
    )


def _ns(prefix="w"):
    return (
        'xmlns:%s="%s" xmlns:w14="%s" xmlns:w15="%s" xmlns:mc="%s" mc:Ignorable="w14 w15"'
        % (prefix, W, W14, W15, MC)
    )


def document_xml(change_author, body=BODY, extra_ins=False, prefix="w", quote='"'):
    p = prefix
    a = quote + escape(change_author, {'"': "&quot;"}) + quote
    more = (
        '<{p}:ins {p}:id="9" {p}:author={a} {p}:date="2026-01-02T03:04:00Z"><{p}:r><{p}:t>!</{p}:t></{p}:r></{p}:ins>'
        if extra_ins
        else ""
    )
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        "<{p}:document {ns}><{p}:body>"
        '<{p}:p w14:paraId="0A1B2C3D"><{p}:r><{p}:t xml:space="preserve">{body} </{p}:t></{p}:r>'
        '<{p}:ins {p}:id="1" {p}:author={a} {p}:date="2026-01-02T03:04:00Z"><{p}:r><{p}:t>Added words.</{p}:t></{p}:r></{p}:ins>'
        '<{p}:del {p}:id="2" {p}:author={a} {p}:date="2026-01-02T03:05:00Z"><{p}:r><{p}:delText>Removed words.</{p}:delText></{p}:r></{p}:del>'
        "<{p}:r><{p}:rPr><{p}:i/>"
        '<{p}:rPrChange {p}:id="3" {p}:author={a} {p}:date="2026-01-02T03:06:00Z"><{p}:rPr/></{p}:rPrChange>'
        "</{p}:rPr><{p}:t>Escherichia coli</{p}:t></{p}:r>"
        '<{p}:commentRangeStart {p}:id="0"/><{p}:r><{p}:t> grew.</{p}:t></{p}:r><{p}:commentRangeEnd {p}:id="0"/>'
        '<{p}:r><{p}:commentReference {p}:id="0"/></{p}:r>' + more + "</{p}:p>"
        "<{p}:sectPr/></{p}:body></{p}:document>"
    )
    return xml.format(p=p, ns=_ns(p), a=a, body=escape(body))


def build(path, body=BODY, extra_ins=False, prefix="w", quote='"', parts=None, extra=None, **marks):
    """Write a docx. `parts`: names of optional parts to leave out, e.g. {"word/people.xml"}.
    `extra`: {name: bytes} of further zip members."""
    m = dict(DIRTY)
    m.update(marks)
    skip = set(parts or ())
    files = {}
    files["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    files["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    files["docProps/core.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title>An invented manuscript</dc:title>"
        "<dc:creator>%s</dc:creator><cp:lastModifiedBy>%s</cp:lastModifiedBy>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>'
        "</cp:coreProperties>" % (escape(m["creator"]), escape(m["last_modified_by"]))
    )
    files["docProps/app.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        "<Template>Normal.dotm</Template><Manager>%s</Manager><Company>%s</Company><Pages>1</Pages>"
        "</Properties>" % (escape(m["manager"]), escape(m["company"]))
    )
    props = "".join(
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="Project code %d">'
        "<vt:lpwstr>EX-%d</vt:lpwstr></property>" % (i + 2, i, i)
        for i in range(m["custom_properties"])
    )
    files["docProps/custom.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">%s</Properties>'
        % props
    )
    files["word/document.xml"] = document_xml(m["change_author"], body, extra_ins, prefix, quote)
    files["word/comments.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        "<w:comments %s>"
        '<w:comment w:id="0" w:author=%s w:date="2026-01-03T00:00:00Z" w:initials=%s>'
        "<w:p><w:r><w:t>Please check this sentence.</w:t></w:r></w:p></w:comment></w:comments>"
        % (_ns(), quoteattr(m["comment_author"]), quoteattr(m["comment_initials"]))
    )
    people = "".join(
        "<w15:person w15:author=%s>%s</w15:person>"
        % (
            quoteattr(name),
            '<w15:presenceInfo w15:providerId="AD" w15:userId="S::%d@example.org::0000"/>' % i
            if m["presence"]
            else "",
        )
        for i, name in enumerate(m["people"])
    )
    files["word/people.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w15:people %s>%s</w15:people>' % (_ns(), people)
    )
    files["word/footnotes.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:footnotes %s><w:footnote w:id="1"><w:p><w:r><w:t>A note.</w:t></w:r>'
        '<w:ins w:id="7" w:author=%s w:date="2026-01-04T00:00:00Z"><w:r><w:t> More.</w:t></w:r></w:ins>'
        "</w:p></w:footnote></w:footnotes>" % (_ns(), quoteattr(m["note_change_author"]))
    )
    files["word/styles.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="%s"/>' % W
    )
    files.update(extra or {})
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            if name not in skip:
                z.writestr(name, data)
    return path
