"""Tests for tools/figure_render.py on the invented fixture figure
(scripts/e2e/make_figure_fixture.py).

The renders need PyMuPDF; where it is not installed those tests are skipped
with the reason and the rest still runs.
Run: python3 -m unittest tests/test_figure_render.py
"""

import json
import struct
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TOOL = ROOT / "tools" / "figure_render.py"
FIXTURE = ROOT / "scripts" / "e2e" / "make_figure_fixture.py"

try:
    import fitz  # noqa: F401  (PyMuPDF)
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False
NEEDS_FITZ = unittest.skipUnless(HAS_FITZ, "python3 cannot import fitz (PyMuPDF)")
PNG = b"\x89PNG\r\n\x1a\n"


def run(script, *args):
    p = subprocess.run([sys.executable, str(script), *map(str, args)], capture_output=True, text=True)
    return p.returncode, json.loads(p.stdout.strip().splitlines()[-1])


def png_size(path):
    data = Path(path).read_bytes()
    assert data[:8] == PNG, "not a PNG file"
    return struct.unpack(">II", data[16:24])


class FixtureTest(unittest.TestCase):
    def test_fixture_is_an_invented_growth_curve_with_an_od600_label(self):
        with tempfile.TemporaryDirectory() as d:
            code, out = run(FIXTURE, "--out", d)
            self.assertEqual(code, 0)
            self.assertEqual(out["pdf"], HAS_FITZ)
            before = (Path(d) / "fig2_before.svg").read_text(encoding="utf-8")
            after = (Path(d) / "fig2_after.svg").read_text(encoding="utf-8")
            self.assertIn(">OD600</text>", before)
            self.assertIn(">OD600</text>", after)
            self.assertNotEqual(before, after)
            self.assertIn("control", after)
            self.assertNotIn("control", before)
            want = ["fig2_before.svg", "fig2_after.svg"]
            if HAS_FITZ:
                want += ["fig2_before.pdf", "fig2_after.pdf"]
            self.assertEqual(out["files"], want)
            self.assertEqual(sorted(p.name for p in Path(d).iterdir()), sorted(want))


class NoPyMuPDFTest(unittest.TestCase):
    def test_the_tool_says_so_when_pymupdf_cannot_be_imported(self):
        """A stand-in package named fitz that fails to import, first on the path."""
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "fitz.py").write_text("raise ImportError('not installed')\n", encoding="utf-8")
            p = subprocess.run([sys.executable, str(TOOL), "--probe"], capture_output=True, text=True,
                               env={"PYTHONPATH": d, "PATH": ""})
            out = json.loads(p.stdout.strip().splitlines()[-1])
            self.assertEqual(p.returncode, 3)
            self.assertEqual(out["code"], "no_pymupdf")
            self.assertFalse(out["ok"])


@NEEDS_FITZ
class RenderTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        cls.dir = Path(cls._tmp.name)
        code, _ = run(FIXTURE, "--out", cls.dir)
        assert code == 0

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def render(self, name, out=None, *more):
        out = out or self.dir / ("%s.%s.png" % (name, self.id().rsplit(".", 1)[-1]))
        code, facts = run(TOOL, "--in", self.dir / name, "--out", out, *more)
        return code, facts, out

    def test_probe_says_the_renderer_can_run(self):
        code, out = run(TOOL, "--probe")
        self.assertEqual((code, out["ok"]), (0, True))

    def test_pdf_facts_and_picture(self):
        code, facts, out = self.render("fig2_before.pdf")
        self.assertEqual(code, 0, facts)
        self.assertEqual(facts["type"], "PDF")
        self.assertEqual(facts["pages"], 1)
        self.assertEqual((facts["width_pt"], facts["height_pt"]), (320.0, 240.0))
        # 320 x 240 pt at 150 dpi.
        self.assertEqual((facts["width_px"], facts["height_px"]), (667, 500))
        self.assertEqual(png_size(out), (667, 500))
        self.assertEqual(facts["words"], 2)  # OD600, Time
        self.assertGreaterEqual(facts["paths"], 2)
        self.assertEqual(facts["images"], 0)
        self.assertTrue(facts["fonts"])
        self.assertTrue(all(f["embedded"] for f in facts["fonts"]))

    def test_after_is_wider_and_has_more_words(self):
        _, before, _ = self.render("fig2_before.pdf")
        _, after, _ = self.render("fig2_after.pdf", self.dir / "after-wider.png")
        self.assertEqual((after["width_px"], after["height_px"]), (834, 500))
        self.assertGreater(after["words"], before["words"])

    def test_dpi_scales_the_picture(self):
        code, facts, out = self.render("fig2_before.pdf", None, "--dpi", "72")
        self.assertEqual(code, 0)
        self.assertEqual(png_size(out), (320, 240))

    def test_svg_has_live_text_and_names_no_font(self):
        code, facts, out = self.render("fig2_after.svg")
        self.assertEqual(code, 0, facts)
        self.assertEqual(facts["type"], "SVG")
        self.assertGreaterEqual(facts["words"], 4)
        self.assertEqual(facts["fonts"], [])
        self.assertEqual(png_size(out), (834, 500))

    def test_a_picture_file_keeps_its_own_pixel_size_and_has_no_text(self):
        _, _, png = self.render("fig2_before.pdf", self.dir / "raster-in.png", "--dpi", "100")
        code, facts = run(TOOL, "--in", png, "--out", self.dir / "raster-out.png", "--dpi", "300")
        self.assertEqual(code, 0, facts)
        self.assertEqual(facts["type"], "PNG")
        self.assertEqual((facts["width_px"], facts["height_px"]), png_size(png))
        self.assertEqual((facts["width_pt"], facts["words"], facts["fonts"]), (None, 0, []))

    def test_an_ai_file_is_opened_as_pdf(self):
        ai = self.dir / "fig2_after.ai"
        ai.write_bytes((self.dir / "fig2_after.pdf").read_bytes())
        code, facts = run(TOOL, "--in", ai, "--out", self.dir / "ai.png")
        self.assertEqual(code, 0, facts)
        self.assertEqual(facts["type"], "AI")
        self.assertEqual(facts["width_px"], 834)

    def test_a_file_at_out_is_not_replaced(self):
        out = self.dir / "taken.png"
        out.write_text("mine", encoding="utf-8")
        code, facts, _ = self.render("fig2_before.pdf", out)
        self.assertEqual(code, 1)
        self.assertEqual(facts["code"], "refused")
        self.assertEqual(out.read_text(encoding="utf-8"), "mine")

    def test_the_input_is_only_read(self):
        src = self.dir / "fig2_before.pdf"
        before = src.read_bytes()
        self.render("fig2_before.pdf")
        self.assertEqual(src.read_bytes(), before)

    def test_what_is_not_a_figure_file_is_refused(self):
        txt = self.dir / "notes.txt"
        txt.write_text("not a figure", encoding="utf-8")
        code, facts = run(TOOL, "--in", txt, "--out", self.dir / "notes.png")
        self.assertEqual((code, facts["code"]), (1, "refused"))
        self.assertFalse((self.dir / "notes.png").exists())
        broken = self.dir / "broken.pdf"
        broken.write_text("not a PDF", encoding="utf-8")
        code, facts = run(TOOL, "--in", broken, "--out", self.dir / "broken.png")
        self.assertEqual((code, facts["code"]), (1, "unreadable"))
        self.assertFalse((self.dir / "broken.png").exists())
        code, facts = run(TOOL, "--in", src_missing(self.dir), "--out", self.dir / "none.png")
        self.assertEqual(code, 1)

    def test_arguments_are_required(self):
        code, facts = run(TOOL, "--in", self.dir / "fig2_before.pdf")
        self.assertEqual((code, facts["code"]), (2, "arguments"))


def src_missing(folder):
    return folder / "not_there.pdf"


if __name__ == "__main__":
    unittest.main()
