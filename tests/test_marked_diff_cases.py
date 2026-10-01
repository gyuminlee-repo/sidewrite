"""Cases the marked-diff tool once got wrong, each with a small invented document.

Run: python3 -m unittest tests/test_marked_diff_cases.py
"""

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TOOL = ROOT / "tools" / "build_marked_diff.py"

P1 = "The first paragraph of this invented text stays as it is."
P2 = "The second paragraph mentions Figure 1 and nothing else."
P2_NEW = P2.replace("nothing", "little")


def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), v)
    return e


def doc_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


class Case(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.n = 0

    def tearDown(self):
        self.tmp.cleanup()

    def doc(self, paragraphs, edit=None):
        """Write a document with one run per paragraph; `edit(document)` may change it before it is saved."""
        self.n += 1
        path = self.dir / ("doc%d.docx" % self.n)
        d = Document()
        for text in paragraphs:
            d.add_paragraph().add_run(text)
        if edit:
            edit(d)
        d.save(path)
        return path

    def run_tool(self, base, clean, *args):
        self.n += 1
        out = self.dir / ("out%d.docx" % self.n)
        r = subprocess.run(
            [sys.executable, str(TOOL), "--baseline", str(base), "--clean", str(clean), "--out", str(out), "--json", *args],
            capture_output=True, text=True,
        )
        last = r.stdout.strip().splitlines()[-1] if r.stdout.strip() else ""
        return r.returncode, out, (json.loads(last) if last.startswith("{") else None), r.stdout + r.stderr

    def accept_reject(self, out):
        """(accept-all, reject-all) paragraph texts of a colour-mode marked file: struck = deleted, red = inserted."""
        accept, reject = [], []
        for p in Document(out).paragraphs:
            runs = [(r.text, bool(r.font.strike), str(r.font.color.rgb) if r.font.color.rgb else None) for r in p.runs]
            accept.append("".join(t for t, struck, _ in runs if not struck))
            reject.append("".join(t for t, struck, colour in runs if struck or colour != "FF0000"))
        return [t for t in accept if t], [t for t in reject if t]


class StrikeSwitchedOff(Case):
    """<w:strike w:val="0"/> says "not struck"; the element being there does not make the run struck."""

    def off(self, tag, val):
        def edit(d):
            rpr = d.paragraphs[0].runs[0]._r.get_or_add_rPr()
            rpr.append(el("w:" + tag, val=val))

        return edit

    def test_a_clean_run_with_strike_switched_off_is_ordinary_text(self):
        base = self.doc([P1, P2])
        for tag in ("strike", "dstrike"):
            for val in ("0", "false", "off"):
                with self.subTest(tag=tag, val=val):
                    clean = self.doc([P1, P2_NEW], self.off(tag, val))
                    self.assertIn('<w:%s w:val="%s"/>' % (tag, val), doc_xml(clean))
                    code, out, summary, log = self.run_tool(base, clean)
                    self.assertEqual(code, 0, log)
                    self.assertEqual(self.accept_reject(out), ([P1, P2_NEW], [P1, P2]))

    def test_a_baseline_run_with_strike_switched_off_is_read_as_text(self):
        base = self.doc([P1, P2], self.off("strike", "0"))
        clean = self.doc([P1, P2_NEW])
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(code, 0, log)
        # Only the changed word is marked: the first paragraph was read, so it is not "inserted".
        self.assertEqual(summary["paragraphs"]["n_ins"], 0)
        self.assertEqual(self.accept_reject(out), ([P1, P2_NEW], [P1, P2]))

    def test_strike_switched_on_still_means_struck(self):
        base = self.doc([P1, P2])
        for val in (None, "1", "true", "on"):
            with self.subTest(val=val):
                def edit(d, val=val):
                    rpr = d.paragraphs[0].runs[0]._r.get_or_add_rPr()
                    rpr.append(el("w:strike", **({"val": val} if val else {})))

                clean = self.doc([P1, P2_NEW], edit)
                code, out, summary, log = self.run_tool(base, clean)
                self.assertEqual(code, 2, log)
                self.assertIn("struck-through", log)
                self.assertFalse(out.exists())


if __name__ == "__main__":
    unittest.main()
