"""Tests for scripts/make-test-response.py: every variant is a readable docx
that really holds the construct it is named after.

Run: python3 -m unittest tests/test_make_test_response.py
"""

import importlib.util
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location("mtr", ROOT / "scripts" / "make-test-response.py")
mtr = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mtr)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
# variant -> elements document.xml must hold at least once
MUST_HOLD = {
    "plain": ["color"],
    "table": ["tbl", "tc"],
    "tracked": ["ins", "del", "delText", "moveFrom"],
    "styled": ["rStyle", "pStyle"],
    "eastasian": ["rFonts", "lang"],
    "highlight": ["highlight", "shd"],
}


class VariantFilesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def build(self, variant, name=None):
        path = Path(self.tmp.name) / (name or f"{variant}.docx")
        mtr.build(path, variant)
        return path

    def test_six_variants(self):
        self.assertEqual(sorted(mtr.VARIANTS), sorted(MUST_HOLD))

    def test_each_variant_opens_and_holds_its_construct(self):
        for variant, tags in MUST_HOLD.items():
            with self.subTest(variant=variant):
                z = zipfile.ZipFile(self.build(variant))
                self.assertIsNone(z.testzip())
                for name in z.namelist():
                    ET.fromstring(z.read(name))
                root = ET.fromstring(z.read("word/document.xml"))
                for tag in tags:
                    self.assertGreaterEqual(len(list(root.iter(W + tag))), 1, tag)

    def test_styled_has_a_styles_part_with_a_based_on_step_and_a_theme_colour(self):
        z = zipfile.ZipFile(self.build("styled"))
        styles = ET.fromstring(z.read("word/styles.xml"))
        self.assertEqual(len(list(styles.iter(W + "basedOn"))), 1)
        self.assertIn("styles.xml", z.read("[Content_Types].xml").decode("utf8"))
        self.assertIn("styles.xml", z.read("word/_rels/document.xml.rels").decode("utf8"))
        doc = ET.fromstring(z.read("word/document.xml"))
        colours = list(doc.iter(W + "color"))
        self.assertTrue(any(c.get(W + "themeColor") == "accent1" for c in colours))
        self.assertTrue(any(c.get(W + "val") == "auto" for c in colours))

    def test_east_asian_runs_carry_the_east_asian_settings(self):
        doc = ET.fromstring(zipfile.ZipFile(self.build("eastasian")).read("word/document.xml"))
        self.assertTrue(any(f.get(W + "eastAsia") for f in doc.iter(W + "rFonts")))
        self.assertTrue(any(l.get(W + "eastAsia") == "ko-KR" for l in doc.iter(W + "lang")))

    def test_the_letter_says_it_is_made_up(self):
        doc = ET.fromstring(zipfile.ZipFile(self.build("plain")).read("word/document.xml"))
        first = "".join(t.text or "" for t in next(doc.iter(W + "p")).iter(W + "t"))
        self.assertIn("Fake response letter for software testing", first)

    def test_same_variant_same_bytes(self):
        for variant in mtr.VARIANTS:
            self.assertEqual(self.build(variant).read_bytes(), self.build(variant, "again.docx").read_bytes(), variant)

    def test_the_committed_fixture_is_the_plain_variant(self):
        # Part by part: the compressed bytes may differ between zlib builds.
        def parts(path):
            z = zipfile.ZipFile(path)
            return {name: z.read(name) for name in z.namelist()}

        fixture = ROOT / "tests" / "fixtures" / "prior_response.docx"
        self.assertEqual(parts(fixture), parts(self.build("plain")))

    def test_unknown_variant_is_refused(self):
        with self.assertRaises(ValueError):
            mtr.build(Path(self.tmp.name) / "x.docx", "nope")


if __name__ == "__main__":
    unittest.main()
