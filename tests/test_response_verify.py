"""Tests for tools/response_verify.py: the written letter against its model
and the manuscript paragraph list.

Run: python3 -m unittest tests/test_response_verify.py
"""

import contextlib
import copy
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import response_docx as rd  # noqa: E402
import response_verify as rv  # noqa: E402


def run_main(argv):
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        code = rv.main(argv)
    return code, json.loads(out.getvalue().strip().splitlines()[-1])


class VerifyTest(unittest.TestCase):
    MODEL = {"paragraphs": [
        {"kind": "reviewer_head", "text": "Reviewer #1", "color": None},
        {"kind": "comment", "text": "State the unit.", "color": None},
        {"kind": "response", "text": "We now give it.", "color": "C00000"},
        {"kind": "location", "text": "[Revised in 2. Methods, line 9]", "color": "0070C0"},
        {"kind": "quote", "text": "The temperature was held at 30 °C.", "color": "0070C0"}]}
    MANUSCRIPT = ["2. Methods", "Flasks were shaken.  The temperature was held at 30 °C. Samples were drawn."]

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)

    def run_cli(self, model=None, manuscript=None, mutate=None, docx_model=None):
        """Write model.json, manuscript.json and the docx into a temp dir, apply
        mutate(document.xml) -> xml to the docx, call rv.main."""
        model = self.MODEL if model is None else model
        manuscript = self.MANUSCRIPT if manuscript is None else manuscript
        (self.dir / "model.json").write_text(json.dumps(model), encoding="utf-8")
        (self.dir / "manuscript.json").write_text(json.dumps(manuscript), encoding="utf-8")
        docx = self.dir / "letter.docx"
        rd.write_docx(model if docx_model is None else docx_model, docx)
        if mutate:
            with zipfile.ZipFile(docx) as z:
                parts = [(n, z.read(n)) for n in z.namelist()]
            parts = [(n, mutate(d.decode("utf8")) if n == "word/document.xml" else d) for n, d in parts]
            rd.write_zip(str(docx), parts)
        return run_main(["--docx", str(docx), "--model", str(self.dir / "model.json"),
                         "--manuscript", str(self.dir / "manuscript.json")])

    def test_good_letter_passes(self):
        code, out = self.run_cli()
        self.assertEqual((code, out["ok"], out["failures"]), (0, True, []))
        self.assertEqual(out["checked"], {"paragraphs": 5, "quotes": 1, "locations": 1})

    def test_each_axis_fails_alone(self):
        cases = {"colour": lambda x: x.replace("C00000", "C00001"),
                 "text": lambda x: x.replace("held at 30", "kept at 30"),
                 "count": lambda x: x.replace("<w:p>", "<w:p><w:r><w:t>x</w:t></w:r></w:p><w:p>", 1)}
        for axis, m in cases.items():
            code, out = self.run_cli(mutate=m)
            self.assertEqual(code, 1)
            self.assertIn(axis, {f["axis"] for f in out["failures"]}, axis)
            self.assertFalse(out["ok"])

    def test_colour_and_count_failures_name_only_their_axis(self):
        code, out = self.run_cli(mutate=lambda x: x.replace("C00000", "C00001"))
        self.assertEqual([(f["axis"], f["index"]) for f in out["failures"]], [("colour", 2)])
        code, out = self.run_cli(mutate=lambda x: x.replace("<w:p>", "<w:p><w:r><w:t>x</w:t></w:r></w:p><w:p>", 1))
        self.assertEqual([f["axis"] for f in out["failures"]], ["count"])

    def test_a_run_that_lost_its_colour_or_got_a_second_one_fails(self):
        code, out = self.run_cli(mutate=lambda x: x.replace('<w:rPr><w:color w:val="C00000"/></w:rPr>', ""))
        self.assertEqual([(f["axis"], f["index"]) for f in out["failures"]], [("colour", 2)])
        extra = '<w:r><w:rPr><w:color w:val="00FF00"/></w:rPr><w:t>!</w:t></w:r></w:p>'
        code, out = self.run_cli(mutate=lambda x: x.replace("</w:p>", extra, 1))
        self.assertEqual({f["axis"] for f in out["failures"]}, {"colour", "text"})

    def test_quote_missing_from_the_manuscript(self):
        code, out = self.run_cli(manuscript=["2. Methods", "Something else."])
        self.assertEqual(code, 1)
        self.assertEqual([f["axis"] for f in out["failures"]], ["quote_in_manuscript"])

    def test_quote_split_over_two_manuscript_paragraphs_is_not_in_the_manuscript(self):
        code, out = self.run_cli(manuscript=["The temperature was held", "at 30 °C."])
        self.assertEqual([f["axis"] for f in out["failures"]], ["quote_in_manuscript"])

    def test_location_without_a_quote(self):
        m = {"paragraphs": self.MODEL["paragraphs"][:4]}
        self.assertEqual([f["axis"] for f in self.run_cli(model=m)[1]["failures"]], ["location_has_quote"])

    def test_a_removal_needs_no_quote(self):
        m = copy.deepcopy({"paragraphs": self.MODEL["paragraphs"][:4]})
        m["paragraphs"][3].update(text="[Removed from 2. Methods]", removed=True)
        code, out = self.run_cli(model=m)
        self.assertEqual((code, out["failures"]), (0, []))
        self.assertEqual(out["checked"], {"paragraphs": 4, "quotes": 0, "locations": 1})

    def test_the_docx_is_checked_not_the_model(self):
        # The model says one thing, the file on disk another.
        other = copy.deepcopy(self.MODEL)
        other["paragraphs"][2]["text"] = "We refuse."
        code, out = self.run_cli(docx_model=other)
        self.assertEqual([(f["axis"], f["index"]) for f in out["failures"]], [("text", 2)])

    def test_control_characters_the_writer_drops_are_no_difference(self):
        m = copy.deepcopy(self.MODEL)
        m["paragraphs"][2]["text"] = "We now give it.\u0005"
        self.assertEqual(self.run_cli(model=m)[0], 0)

    def test_missing_manuscript_file_is_exit_2(self):
        self.assertEqual(rv.main(["--docx", "a", "--model", "b", "--manuscript", "/no/such.json"]), 2)

    def test_inputs_that_are_not_what_they_should_be_are_exit_2(self):
        self.run_cli()
        docx, model, ms = (str(self.dir / n) for n in ("letter.docx", "model.json", "manuscript.json"))
        (self.dir / "list.json").write_text('{"a": 1}', encoding="utf-8")
        (self.dir / "bad.json").write_text("{", encoding="utf-8")
        for argv in (["--docx", docx, "--model", model, "--manuscript", str(self.dir / "list.json")],
                     ["--docx", docx, "--model", str(self.dir / "bad.json"), "--manuscript", ms],
                     ["--docx", model, "--model", model, "--manuscript", ms]):
            code, out = run_main(argv)
            self.assertEqual((code, out["ok"]), (2, False), argv)
            self.assertTrue(out["error"])

    def test_selftest(self):
        self.assertEqual(rv.main(["--selftest"]), 0)

    def test_selftest_fails_when_the_verifier_is_blind(self):
        real = rv.verify
        rv.verify = lambda *a, **k: {"ok": True, "checked": {}, "failures": []}
        try:
            code, out = run_main(["--selftest"])
        finally:
            rv.verify = real
        self.assertEqual(code, 1)
        self.assertFalse(out["ok"])


if __name__ == "__main__":
    unittest.main()
