"""Tests for tools/response_docx.py: the letter writer.

Run: python3 -m unittest tests/test_response_docx.py
"""

import contextlib
import io
import json
import re
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import response_docx as rd  # noqa: E402

MODEL = {"paragraphs": [
    {"kind": "reviewer_head", "text": "Reviewer #1", "color": None},
    {"kind": "comment", "text": "Is 5 < 6 & why?", "color": None},
    {"kind": "response", "text": "We agree.\u0005", "color": "aa0000"},
    {"kind": "location", "text": "[Revised in 2. Methods, line 9]", "color": "00AA00"},
    {"kind": "quote", "text": "Held at 30 °C.", "color": "00AA00"}]}


def run_main(argv):
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        code = rd.main(argv)
    return code, json.loads(out.getvalue().strip().splitlines()[-1])


class BuildTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)

    def build(self, name="a.docx", model=MODEL):
        out = self.dir / name
        rd.write_docx(model, out)
        return out

    def test_parts_text_and_colours(self):
        z = zipfile.ZipFile(self.build())
        self.assertEqual(sorted(z.namelist()), ["[Content_Types].xml", "_rels/.rels", "word/document.xml"])
        x = z.read("word/document.xml").decode("utf8")
        paras = re.findall(r"<w:p>.*?</w:p>", x, re.S)
        self.assertEqual(len(paras), 5)
        self.assertIn("Is 5 &lt; 6 &amp; why?", paras[1])
        self.assertNotIn("w:color", paras[1])
        self.assertIn('<w:color w:val="AA0000"/>', paras[2])
        self.assertNotIn("\u0005", x)
        self.assertIn("<w:b/>", paras[0])
        self.assertNotIn("<w:b/>", paras[2])
        self.assertIn("Held at 30 °C.", paras[4])

    def test_every_part_is_well_formed_xml(self):
        z = zipfile.ZipFile(self.build())
        for name in z.namelist():
            ET.fromstring(z.read(name))

    def test_headings_are_bold(self):
        model = {"paragraphs": [{"kind": k, "text": k, "color": None}
                                for k in ("reviewer_head", "section_head", "comment_head", "comment", "declined")]}
        x = zipfile.ZipFile(self.build(model=model)).read("word/document.xml").decode("utf8")
        paras = re.findall(r"<w:p>.*?</w:p>", x, re.S)
        self.assertEqual(["<w:b/>" in p for p in paras], [True, True, True, False, False])

    def test_font_and_size_go_on_every_run(self):
        model = {"font": 'A "quoted" font', "size_pt": 11, "paragraphs": MODEL["paragraphs"]}
        x = zipfile.ZipFile(self.build(model=model)).read("word/document.xml").decode("utf8")
        self.assertEqual(x.count('<w:sz w:val="22"/>'), 5)
        root = ET.fromstring(x)
        w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        self.assertEqual({f.get(w + "ascii") for f in root.iter(w + "rFonts")}, {'A "quoted" font'})

    def test_same_model_same_bytes(self):
        a = self.build().read_bytes()
        b = self.build("b.docx").read_bytes()
        self.assertEqual(a, b)

    def test_bad_colour_is_refused_and_nothing_is_written(self):
        out = self.dir / "x.docx"
        with self.assertRaises(ValueError):
            rd.write_docx({"paragraphs": [{"kind": "response", "text": "x", "color": "red"}]}, out)
        self.assertEqual(list(self.dir.iterdir()), [])

    def test_cli_writes_and_reports(self):
        model = self.dir / "model.json"
        model.write_text(json.dumps(MODEL), encoding="utf-8")
        out = self.dir / "letter.docx"
        code, rep = run_main(["--model", str(model), "--out", str(out)])
        self.assertEqual((code, rep), (0, {"ok": True, "path": str(out), "paragraphs": 5}))
        self.assertTrue(zipfile.is_zipfile(out))

    def test_cli_errors_are_exit_1_with_a_reason(self):
        model = self.dir / "model.json"
        model.write_text("{not json", encoding="utf-8")
        code, rep = run_main(["--model", str(model), "--out", str(self.dir / "o.docx")])
        self.assertEqual((code, rep["ok"], rep["code"]), (1, False, "error"))
        code, rep = run_main(["--model", str(self.dir / "none.json"), "--out", str(self.dir / "o.docx")])
        self.assertEqual((code, rep["code"]), (1, "error"))
        model.write_text(json.dumps({"paragraphs": [{"kind": "response", "text": "x", "color": "red"}]}), encoding="utf-8")
        code, rep = run_main(["--model", str(model), "--out", str(self.dir / "o.docx")])
        self.assertEqual(code, 1)
        self.assertIn("RRGGBB", rep["error"])

    def test_locked_target_is_exit_3_and_leaves_no_temp_file(self):
        model = self.dir / "model.json"
        model.write_text(json.dumps(MODEL), encoding="utf-8")
        out = self.dir / "open.docx"
        with mock.patch.object(rd.os, "replace", side_effect=PermissionError("in use")):
            code, rep = run_main(["--model", str(model), "--out", str(out)])
        self.assertEqual((code, rep["code"]), (3, "locked"))
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["model.json"])


if __name__ == "__main__":
    unittest.main()
