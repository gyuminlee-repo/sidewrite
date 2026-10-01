"""Tests for tools/response_docx.py: the letter writer.

Run: python3 -m unittest tests/test_response_docx.py
"""

import contextlib
import io
import json
import re
import sys
import tempfile
import threading
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

    def test_two_writes_to_one_path_at_once_both_succeed(self):
        out = self.dir / "letter.docx"
        models = [{"paragraphs": [{"kind": "response", "text": t * 200, "color": None}]} for t in "AB"]
        # Both writers have written their temp file before either moves it.
        both_written = threading.Barrier(2, timeout=10)
        real = rd.write_zip

        def write_then_wait(*args, **kwargs):
            real(*args, **kwargs)
            both_written.wait()

        errors = []

        def work(model):
            try:
                rd.write_docx(model, out)
            except Exception as e:  # noqa: BLE001  (the test reports whatever went wrong)
                errors.append(repr(e))

        with mock.patch.object(rd, "write_zip", write_then_wait):
            threads = [threading.Thread(target=work, args=(m,)) for m in models]
            for t in threads:
                t.start()
            for t in threads:
                t.join()
        self.assertEqual(errors, [])
        z = zipfile.ZipFile(out)
        self.assertIsNone(z.testzip())
        self.assertIn(z.read("word/document.xml").decode("utf8"), [rd.document_xml(m) for m in models])
        self.assertEqual([p.name for p in self.dir.iterdir()], ["letter.docx"])

    def test_a_tmp_file_that_was_already_there_is_left_alone(self):
        out = self.dir / "letter.docx"
        other = self.dir / "letter.docx.tmp"
        other.write_text("not ours", encoding="utf-8")
        rd.write_docx(MODEL, out)
        self.assertEqual(other.read_text(encoding="utf-8"), "not ours")
        self.assertTrue(zipfile.is_zipfile(out))
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["letter.docx", "letter.docx.tmp"])
        # Also when the write fails.
        with mock.patch.object(rd.os, "replace", side_effect=PermissionError("in use")):
            with self.assertRaises(rd.Locked):
                rd.write_docx(MODEL, out)
        self.assertEqual(other.read_text(encoding="utf-8"), "not ours")

    def test_a_failed_write_leaves_no_temp_file_and_keeps_the_old_letter(self):
        out = self.dir / "letter.docx"
        rd.write_docx(MODEL, out)
        before = out.read_bytes()
        real = rd.write_zip

        def half_written(*args, **kwargs):
            real(*args, **kwargs)
            raise OSError("disk full")

        with mock.patch.object(rd, "write_zip", half_written):
            with self.assertRaises(OSError):
                rd.write_docx({"paragraphs": [{"kind": "response", "text": "new", "color": None}]}, out)
        self.assertEqual([p.name for p in self.dir.iterdir()], ["letter.docx"])
        self.assertEqual(out.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
