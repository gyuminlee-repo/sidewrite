#!/usr/bin/env python3
"""Copy a .docx and tag the copy so Word opens the add-in task pane with it.

Adds the two Open XML parts that Microsoft documents for "Automatically open a
task pane with a document": a webextension part whose property bag sets
Office.AutoShowTaskpaneWithDocument=true and a taskpane part (visibility=1).
The reference uses storeType="Registry", store="developer", which is the row
for a sideloaded add-in (manifest registered under WEF\\Developer).

The add-in id and version are read from the add-in manifest. The source file
is never written.

Usage: embed_taskpane.py --manifest manifests/word.xml --src in.docx --out out.docx
"""

import argparse
import re
import sys
import uuid
import zipfile

WE_PART = "word/webextensions/webextension1.xml"
TP_PART = "word/webextensions/taskpanes.xml"
TP_RELS = "word/webextensions/_rels/taskpanes.xml.rels"
REL_TASKPANES = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
REL_WEBEXT = "http://schemas.microsoft.com/office/2011/relationships/webextension"
CT_WEBEXT = "application/vnd.ms-office.webextension+xml"
CT_TASKPANES = "application/vnd.ms-office.webextensiontaskpanes+xml"


def manifest_id_version(text):
    """Return (id, version) of an add-in only (XML) manifest."""
    m_id = re.search(r"<Id>\s*([0-9A-Fa-f-]{36})\s*</Id>", text)
    m_ver = re.search(r"<Version>\s*([0-9.]+)\s*</Version>", text)
    if not m_id or not m_ver:
        raise ValueError("manifest has no <Id> GUID or <Version>")
    return m_id.group(1), m_ver.group(1)


def webextension_xml(addin_id, version, instance_id):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"'
        f' id="{{{instance_id}}}">'
        f'<we:reference id="{addin_id}" version="{version}" store="developer" storeType="Registry"/>'
        "<we:alternateReferences/>"
        '<we:properties><we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/></we:properties>'
        "<we:bindings/>"
        '<we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        "</we:webextension>"
    )


def taskpanes_xml(width=350):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">'
        f'<wetp:taskpane dockstate="right" visibility="1" width="{width}" row="4">'
        '<wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
        "</wetp:taskpane></wetp:taskpanes>"
    )


TP_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{REL_WEBEXT}" Target="webextension1.xml"/>'
    "</Relationships>"
)


def add_root_rel(rels):
    if REL_TASKPANES in rels:
        return rels
    ids = {int(n) for n in re.findall(r'Id="rId(\d+)"', rels)}
    new_id = max(ids, default=0) + 1
    rel = f'<Relationship Id="rId{new_id}" Type="{REL_TASKPANES}" Target="{TP_PART}"/>'
    return rels.replace("</Relationships>", rel + "</Relationships>")


def add_content_types(ct):
    for part, kind in ((WE_PART, CT_WEBEXT), (TP_PART, CT_TASKPANES)):
        if f'PartName="/{part}"' not in ct:
            ct = ct.replace("</Types>", f'<Override PartName="/{part}" ContentType="{kind}"/></Types>')
    return ct


def embed(src, out, addin_id, version):
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        if any(n.startswith("word/webextensions/") for n in names):
            raise ValueError(f"{src} already has webextension parts")
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                data = zin.read(n)
                if n == "_rels/.rels":
                    data = add_root_rel(data.decode("utf-8")).encode("utf-8")
                elif n == "[Content_Types].xml":
                    data = add_content_types(data.decode("utf-8")).encode("utf-8")
                zout.writestr(zin.getinfo(n), data)
            zout.writestr(WE_PART, webextension_xml(addin_id, version, str(uuid.uuid4()).upper()))
            zout.writestr(TP_PART, taskpanes_xml())
            zout.writestr(TP_RELS, TP_RELS_XML)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--manifest", required=True)
    ap.add_argument("--src", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    if a.src == a.out:
        ap.error("--out must differ from --src (the source is never written)")
    with open(a.manifest, encoding="utf-8") as f:
        addin_id, version = manifest_id_version(f.read())
    embed(a.src, a.out, addin_id, version)
    print(f"embedded add-in {addin_id} {version} -> {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
