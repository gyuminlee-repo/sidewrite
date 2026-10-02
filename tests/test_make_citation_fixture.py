"""The fake manuscript of the cite acceptance stage (scripts/e2e/make_citation_fixture.py)."""
import base64
import importlib.util
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


fx = load("make_citation_fixture", ROOT / "scripts" / "e2e" / "make_citation_fixture.py")
check = load("citation_fields_check", ROOT / "tools" / "citation_fields_check.py")
W = check.W

try:
    import docx  # noqa: F401

    HAVE_DOCX = True
except ImportError:
    HAVE_DOCX = False


@unittest.skipUnless(HAVE_DOCX, "python-docx is not installed")
class CitationFixture(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.path = fx.build(Path(cls.tmp.name) / "citation.docx")
        cls.doc = check.read(cls.path)
        with zipfile.ZipFile(cls.path) as z:
            cls.body = ET.fromstring(z.read("word/document.xml"))
        cls.paras = list(cls.body.iter(W + "p"))
        cls.texts = ["".join(t.text or "" for t in p.iter(W + "t")) for p in cls.paras]

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def fields(self, kind):
        return [f for f in self.doc["fields"] if f["kind"] == kind]

    def test_the_document_is_balanced_and_every_field_is_live(self):
        self.assertEqual(self.doc["errors"], [])
        self.assertEqual({f["state"] for f in self.doc["fields"]}, {"live"})

    def test_zotero_citations_have_the_code_the_plugin_writes(self):
        items = self.fields("ZOTERO_ITEM")
        self.assertEqual(len(items), 11)
        ids = set()
        for f in items:
            code = f["code"]
            self.assertTrue(code.startswith(" ADDIN ZOTERO_ITEM CSL_CITATION {"), code[:40])
            self.assertTrue(code.endswith("} "))
            data = json.loads(code[len(" ADDIN ZOTERO_ITEM CSL_CITATION ") :])
            self.assertEqual(
                list(data), ["citationID", "properties", "citationItems", "schema"]
            )
            self.assertEqual(data["schema"], fx.SCHEMA)
            self.assertEqual(data["properties"]["noteIndex"], 0)
            shown = "".join(t for t, _ in f["result"])
            self.assertEqual(data["properties"]["plainCitation"], shown)
            self.assertEqual(data["properties"]["formattedCitation"], shown)
            for item in data["citationItems"]:
                self.assertEqual(sorted(item), ["id", "itemData", "uris"])
                self.assertTrue(item["uris"][0].startswith("http://zotero.org/users/local/"))
            ids.add(data["citationID"])
        self.assertEqual(len(ids), 11)
        multi = [f for f in items if "; " in "".join(t for t, _ in f["result"])]
        self.assertEqual(len(multi), 2)

    def test_the_four_places_of_a_zotero_citation(self):
        def para_with(text):
            hits = [i for i, t in enumerate(self.texts) if text in t]
            self.assertEqual(len(hits), 1, text)
            return hits[0]

        # In the middle of a sentence.
        self.assertIn("culture (Arden, 2019) and are cheap", self.texts[para_with("are cheap to run")])
        # Several items in one field, and a field right before the full stop.
        p = self.texts[para_with("Oxygen transfer falls")]
        self.assertIn("rises (Dellmar, 2018; Eskridge, 2022). The", p)
        self.assertTrue(p.endswith("is filled (Arden, 2019)."))
        # Next to italic and superscript runs.
        k = para_with("Examplea fictiva")
        runs = list(self.paras[k].iter(W + "r"))
        self.assertTrue(any(r.find(f"{W}rPr/{W}i") is not None for r in runs))
        self.assertTrue(any(r.find(f"{W}rPr/{W}vertAlign") is not None for r in runs))
        self.assertTrue(any(f["paragraph"]["index"] == k for f in self.fields("ZOTERO_ITEM")))

    def test_some_instructions_are_split_over_several_runs(self):
        counts = []
        for p in self.paras:
            n = 0
            for el in p.iter():
                if el.tag == W + "fldChar" and el.get(W + "fldCharType") == "begin":
                    n = 0
                elif el.tag == W + "instrText":
                    n += 1
                elif el.tag == W + "fldChar" and el.get(W + "fldCharType") == "separate":
                    counts.append(n)
        self.assertIn(1, counts)
        self.assertTrue(max(counts) >= 4, counts)

    def test_the_zotero_bibliography_is_one_field_over_its_paragraphs(self):
        (bibl,) = self.fields("ZOTERO_BIBL")
        self.assertEqual(
            bibl["code"], ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '
        )
        first = bibl["paragraph"]["index"]
        rows = {p["index"]: p for p in self.doc["paragraphs"]}
        self.assertEqual(rows[first]["depth_end"], 1)
        self.assertEqual(rows[first + len(fx.REFS) - 1]["depth_end"], 0)

    def test_the_zotero_document_data_is_in_properties_of_255_characters(self):
        props = self.doc["props"]
        self.assertEqual(sorted(props), ["ZOTERO_PREF_1", "ZOTERO_PREF_2"])
        self.assertEqual(len(props["ZOTERO_PREF_1"]), 255)
        self.assertEqual(props["ZOTERO_PREF_1"] + props["ZOTERO_PREF_2"], fx.ZOTERO_DATA)
        ET.fromstring(fx.ZOTERO_DATA)
        with zipfile.ZipFile(self.path) as z:
            self.assertIn(b"/docProps/custom.xml", z.read("[Content_Types].xml"))
            self.assertIn(b"docProps/custom.xml", z.read("_rels/.rels"))

    def test_endnote_citations_in_both_forms(self):
        short, long = self.fields("EN.CITE")
        self.assertTrue(short["code"].startswith(" ADDIN EN.CITE <EndNote><Cite>"))
        self.assertTrue(short["code"].endswith("</Cite></EndNote>"))
        self.assertIsNone(short["data"])
        ET.fromstring(short["code"][len(" ADDIN EN.CITE ") :])
        self.assertEqual(long["code"], " ADDIN EN.CITE ")
        xml = base64.b64decode(long["data"]).decode("utf-8")
        self.assertEqual(ET.fromstring(xml).find("Cite/DisplayText").text, "(Hallivane, 2016)")
        (nested,) = self.fields("EN.CITE.DATA")
        self.assertEqual((nested["code"], nested["depth"]), (" ADDIN EN.CITE.DATA ", 1))
        self.assertEqual(nested["data"], long["data"])
        self.assertEqual(nested["paragraph"]["index"], long["paragraph"]["index"])
        (reflist,) = self.fields("EN.REFLIST")
        self.assertEqual(reflist["code"], " ADDIN EN.REFLIST ")
        self.assertEqual(sorted(self.doc["vars"]), ["EN.InstantFormat", "EN.Layout", "EN.Libraries"])

    def test_no_paragraph_mixes_the_two_systems(self):
        by_para = {}
        for f in self.doc["fields"]:
            system = "Z" if f["kind"].startswith("ZOTERO") else "E" if f["kind"].startswith("EN.") else None
            if system:
                by_para.setdefault(f["paragraph"]["index"], set()).add(system)
        self.assertTrue(all(len(s) == 1 for s in by_para.values()), by_para)

    def test_the_controls(self):
        (ref,) = self.fields("REF")
        self.assertEqual(ref["code"], f" REF {fx.BOOKMARK} \\h ")
        self.assertTrue(any(b.get(W + "name") == fx.BOOKMARK for b in self.body.iter(W + "bookmarkStart")))
        (link,) = self.fields("hyperlink")
        self.assertEqual(link["code"], "HYPERLINK " + fx.LINK)

    def test_every_text_the_stage_edits_is_in_the_document_once(self):
        whole = "\n".join(self.texts)
        for case in fx.CASES:
            self.assertEqual(whole.count(case["find"]), 1, case["name"])
            self.assertLessEqual(len(case["find"]), 255)
        target = [t for t in self.texts if fx.REWRITE["paragraph_has"] in t]
        self.assertEqual(len(target), 1)
        for old, _new in fx.REWRITE["changes"]:
            self.assertEqual(target[0].count(old), 1, old)
        self.assertEqual(
            [c["name"] for c in fx.CASES if c["replace"] == ""],
            ["cite-text-zotero", "cite-text-endnote", "delete-sentence", "delete-sentence-zotero"],
        )

    def test_what_each_case_of_the_stage_expects(self):
        shown = {f["text"] for f in self.shown_citations()}
        by = {c["name"]: c for c in fx.CASES}
        self.assertEqual({c["expect"] for c in fx.CASES}, {"applied", "refused", "removed"})
        # Refused: the find is the visible text of a citation, or cuts into one,
        # or is replaced over one.
        for name in ("cite-text-zotero", "cite-text-endnote", "cite-text-replace"):
            self.assertIn(by[name]["find"], shown, name)
            self.assertEqual(by[name]["expect"], "refused")
        self.assertEqual((by["spans"]["expect"], by["straddles"]["expect"]), ("refused",) * 2)
        # Removed: a whole sentence, with words before the citation and a
        # character after it, so the range holds the field from begin to end.
        for name in ("delete-sentence", "delete-sentence-zotero"):
            c = by[name]
            self.assertEqual((c["expect"], c["fields"], c["replace"]), ("removed", 1, ""))
            inside = [t for t in shown if t in c["find"]]
            self.assertEqual(len(inside), 1, name)
            at = c["find"].index(inside[0])
            self.assertTrue(at > 0 and at + len(inside[0]) < len(c["find"]), name)
        # Applied: no citation text in the find.
        for name in ("next-to", "formatted"):
            self.assertEqual(by[name]["expect"], "applied")
            self.assertFalse(any(t in by[name]["find"] for t in shown))

    def shown_citations(self):
        return [
            {"text": "".join(t for t, _ in f["result"])}
            for f in self.doc["fields"]
            if f["kind"] in ("ZOTERO_ITEM", "EN.CITE")
        ]

    def test_the_committed_fixture_is_what_the_script_builds(self):
        _lines, result = check.compare(check.read(fx.FIXTURE), self.doc)
        self.assertEqual(result["failures"], [])
        with zipfile.ZipFile(fx.FIXTURE) as a, zipfile.ZipFile(self.path) as b:
            self.assertEqual(a.read("word/document.xml"), b.read("word/document.xml"))


if __name__ == "__main__":
    unittest.main()
