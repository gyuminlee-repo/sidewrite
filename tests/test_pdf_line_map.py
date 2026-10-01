"""Tests for scripts/pdf-line-map.py on synthetic PDFs with known answers.

The PDF is drawn with PyMuPDF: margin numbers at x=30, body at x=72. Only
synthetic geometry is covered here; Word's own PDF layout is not.
Run: python3 -m unittest tests/test_pdf_line_map.py
"""

import importlib.util
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "pdf-line-map.py"
_spec = importlib.util.spec_from_file_location("pdf_line_map", SCRIPT)
plm = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(plm)

# (line_no, body text). Line 5 starts with a digit (must stay body text);
# lines 6/7 break the real compound "cell-free" at its hyphen;
# lines 8/9 break "temperature" with a soft line-end hyphen.
PAGE1 = [
    (1, "Microbial growth on single carbon sources"),
    (2, "is a common starting point for studies."),
    (3, "The aim of this fictional study was to"),
    (4, "describe the growth curve in detail."),
    (5, "30 samples were drawn from each flask."),
    (6, "A cell-"),
    (7, "free extract was prepared by sonication."),
    (8, "The incubation tem-"),
    (9, "perature was held at 30 degrees."),
]
PAGE2 = [
    (10, "Growth reached a plateau after four days."),
    (11, "The spread at the plateau was small."),
]


def make_pdf(path, pages):
    doc = fitz.open()
    for lines in pages:
        page = doc.new_page(width=612, height=792)
        y = 100
        for no, text in lines:
            page.insert_text((30, y), str(no), fontsize=9)
            page.insert_text((72, y), text, fontsize=12)
            y += 24
        page.insert_text((300, 760), "7", fontsize=10)  # centred page number: not a line number
    doc.save(path)
    doc.close()


class PdfLineMapTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.pdf = str(Path(cls.tmp.name) / "synthetic.pdf")
        make_pdf(cls.pdf, [PAGE1, PAGE2])
        cls.lines = plm.extract_lines(cls.pdf)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_lines_paired_with_numbers(self):
        got = [(l["page"], l["line_no"], l["text"]) for l in self.lines]
        want = [(1, n, t) for n, t in PAGE1] + [(2, n, t) for n, t in PAGE2]
        self.assertEqual(got, want)

    def test_digit_at_body_start_is_not_a_line_number(self):
        self.assertEqual(self.lines[4]["text"], "30 samples were drawn from each flask.")

    def test_single_line_quote(self):
        hit = plm.find_quote(self.lines, "a common starting point")
        self.assertEqual(hit["range"], "2")
        self.assertEqual(hit["match_count"], 1)

    def test_two_line_quote(self):
        hit = plm.find_quote(self.lines, "study was to describe the growth")
        self.assertEqual(hit["range"], "3-4")

    def test_real_compound_broken_at_hyphen(self):
        self.assertEqual(plm.find_quote(self.lines, "A cell-free extract")["range"], "6-7")

    def test_soft_hyphen_line_break(self):
        self.assertEqual(plm.find_quote(self.lines, "incubation temperature was held")["range"], "8-9")

    def test_quote_across_pages(self):
        hit = plm.find_quote(self.lines, "30 degrees. Growth reached")
        self.assertEqual(hit["range"], "9-10")
        self.assertEqual((hit["start"]["page"], hit["end"]["page"]), (1, 2))

    def test_substring_trap_first_of_several(self):
        # "plateau" occurs on lines 10 and 11; the first is reported with the count.
        hit = plm.find_quote(self.lines, "plateau")
        self.assertEqual(hit["range"], "10")
        self.assertEqual(hit["match_count"], 2)

    def test_missing_quote(self):
        self.assertIsNone(plm.find_quote(self.lines, "enzyme kinetics were not measured"))

    def test_cli_exit_codes(self):
        ok = subprocess.run([sys.executable, str(SCRIPT), self.pdf, "--quote", "describe the growth"],
                            capture_output=True, text=True)
        self.assertEqual(ok.returncode, 0, ok.stderr)
        self.assertIn('"range": "4"', ok.stdout)
        miss = subprocess.run([sys.executable, str(SCRIPT), self.pdf, "--quote", "not in the pdf"],
                              capture_output=True, text=True)
        self.assertEqual(miss.returncode, 1)


if __name__ == "__main__":
    unittest.main()
