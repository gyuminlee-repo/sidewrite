"""Known-answer check for tools/anonymize_docx.py, judged by scripts/verify-anonymized.py.

Run: python3 -m unittest tests/test_anonymize_docx.py
"""

import hashlib
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import anon_docx_fixture as fx  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
TOOL = ROOT / "tools" / "anonymize_docx.py"


def load(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class AnonymizeDocx(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.verifier = load(ROOT / "scripts" / "verify-anonymized.py", "verify_anonymized")

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.src = fx.build(self.dir / "paper.docx")
        self.out = self.dir / "paper_anonymized.docx"

    def tearDown(self):
        self.tmp.cleanup()

    def run_tool(self, *args):
        r = subprocess.run([sys.executable, str(TOOL), *map(str, args)], capture_output=True, text=True)
        return r.returncode, r.stdout, r.stderr

    def anonymize(self, src=None, out=None, *args):
        return self.run_tool("--in", src or self.src, "--out", out or self.out, *args)

    def test_every_mark_is_cleared_and_the_text_is_unchanged(self):
        before = sha(self.src)
        code, out, err = self.anonymize()
        self.assertEqual(code, 0, err)
        rep = self.verifier.verify(self.src, self.out)
        self.assertEqual(rep["total"], 0, rep["remaining"])
        self.assertEqual([rep["body_equal"], rep["markup_equal"], rep["unreadable_parts"]], [True, True, []])
        self.assertEqual(sha(self.src), before)
        report = json.loads(out)
        self.assertEqual(report["output"], str(self.out))
        self.assertEqual(report["changed"]["word/document.xml"], {"author": 3})
        self.assertEqual(report["changed"]["word/comments.xml"], {"author": 1, "initials": 1})
        self.assertEqual(report["changed"]["docProps/custom.xml"], {"property": 1})
        self.assertEqual(report["changed"]["word/people.xml"], {"person": 2, "presenceInfo": 2})

    def test_the_source_names_are_gone_from_every_part(self):
        self.anonymize()
        with zipfile.ZipFile(self.out) as z:
            blob = b"".join(z.read(n) for n in z.namelist())
        for name in (b"Test Author", b"Second Person", b"Example Lab", b"Lab Head", b"example.org", b"EX-0"):
            self.assertNotIn(name, blob)
        self.assertIn(b"An invented manuscript", blob)  # the title is kept

    def test_other_parts_and_the_zip_layout_stay_as_they_are(self):
        self.anonymize()
        with zipfile.ZipFile(self.src) as a, zipfile.ZipFile(self.out) as b:
            self.assertEqual(a.namelist(), b.namelist())
            self.assertIsNone(b.testzip())
            for name in ("[Content_Types].xml", "_rels/.rels", "word/styles.xml"):
                self.assertEqual(a.read(name), b.read(name))
            for x, y in zip(a.infolist(), b.infolist()):
                self.assertEqual((x.date_time, x.compress_type), (y.date_time, y.compress_type))
            doc = b.read("word/document.xml").decode("utf-8")
            for name in b.namelist():
                if name.endswith(".xml") or name.endswith(".rels"):
                    ET.fromstring(b.read(name))
        # Namespace declarations that mc:Ignorable names must survive, or Word asks to repair the file.
        self.assertIn('xmlns:w14="%s"' % fx.W14, doc)
        self.assertIn('xmlns:w15="%s"' % fx.W15, doc)
        self.assertIn('mc:Ignorable="w14 w15"', doc)
        self.assertIn('w14:paraId="0A1B2C3D"', doc)
        self.assertTrue(doc.startswith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))

    def test_a_chosen_name_is_written_and_escaped(self):
        code, _, err = self.anonymize(None, None, "--author", 'R&D "Copy" <1>', "--initials", "R&D")
        self.assertEqual(code, 0, err)
        rep = self.verifier.verify(self.src, self.out, 'R&D "Copy" <1>', "R&D")
        self.assertEqual(rep["total"], 0, rep["remaining"])
        self.assertEqual(self.verifier.verify(self.src, self.out)["total"] > 0, True)

    def test_another_prefix_and_single_quotes_are_handled(self):
        for kw in (dict(prefix="x"), dict(quote="'")):
            with self.subTest(**kw):
                src = fx.build(self.dir / "odd.docx", **kw)
                out = self.dir / "odd_out.docx"
                code, _, err = self.anonymize(src, out)
                self.assertEqual(code, 0, err)
                rep = self.verifier.verify(src, out)
                self.assertEqual([rep["total"], rep["body_equal"], rep["markup_equal"]], [0, True, True], rep)

    def test_a_docx_without_the_optional_parts_works(self):
        gone = {"word/people.xml", "word/comments.xml", "docProps/custom.xml", "docProps/app.xml", "word/footnotes.xml"}
        src = fx.build(self.dir / "small.docx", parts=gone)
        code, _, err = self.anonymize(src)
        self.assertEqual(code, 0, err)
        self.assertEqual(self.verifier.verify(src, self.out)["total"], 0)

    def test_running_it_on_its_own_output_changes_nothing_more(self):
        self.anonymize()
        again = self.dir / "again.docx"
        code, out, _ = self.anonymize(self.out, again)
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(out)["changed"], {})

    def test_the_output_may_not_be_the_source_under_any_name(self):
        before = sha(self.src)
        link = self.dir / "link.docx"
        os.symlink(self.src, link)
        hard = self.dir / "hard.docx"
        os.link(self.src, hard)
        sub = self.dir / "sub"
        sub.mkdir()
        spellings = [self.src, link, hard, sub / ".." / "paper.docx", Path(str(self.src).replace("paper.docx", "./paper.docx"))]
        for out in spellings:
            with self.subTest(out=str(out)):
                code, _, err = self.anonymize(self.src, out)
                self.assertEqual(code, 2)
                self.assertIn("same file", err)
                self.assertEqual(sha(self.src), before)
        # And the other way round: the source given through a link.
        self.assertEqual(self.anonymize(link, self.src)[0], 2)
        self.assertEqual(sha(self.src), before)
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["hard.docx", "link.docx", "paper.docx", "sub"])

    def test_no_half_written_file_is_left_behind(self):
        bad = self.dir / "bad.docx"
        bad.write_text("not a zip")
        code, _, err = self.anonymize(bad)
        self.assertEqual(code, 2)
        self.assertFalse(self.out.exists())
        missing_dir = self.dir / "nowhere" / "out.docx"
        self.assertEqual(self.anonymize(self.src, missing_dir)[0], 2)
        self.anonymize()
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["bad.docx", "paper.docx", "paper_anonymized.docx"])

    def test_an_existing_output_is_replaced_whole(self):
        self.out.write_text("old")
        code, _, err = self.anonymize()
        self.assertEqual(code, 0, err)
        self.assertTrue(zipfile.is_zipfile(self.out))


if __name__ == "__main__":
    unittest.main()
