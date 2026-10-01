"""Known-answer check for scripts/verify-anonymized.py.

The verifier must tell a clean docx from one that keeps a single author mark,
target by target, and must see a changed body. Run:
python3 -m unittest tests/test_verify_anonymized.py
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import anon_docx_fixture as fx  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "verify-anonymized.py"

# One mark left in an otherwise clean docx -> the report key that must be 1.
LEFT_IN = {
    "docProps/core.xml:creator": dict(creator="Test Author"),
    "docProps/core.xml:lastModifiedBy": dict(last_modified_by="Second Person"),
    "docProps/app.xml:Company": dict(company="Example Lab"),
    "docProps/app.xml:Manager": dict(manager="Lab Head"),
    "docProps/custom.xml:property": dict(custom_properties=1),
    "word/comments.xml:author": dict(comment_author="Second Person"),
    "word/comments.xml:initials": dict(comment_initials="SP"),
    "word/people.xml:person": dict(people=("Anonymous", "Second Person")),
    "word/people.xml:presenceInfo": dict(people=("Anonymous",), presence=True),
    "word/footnotes.xml:change_author": dict(note_change_author="Test Author"),
}


class VerifyAnonymized(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.before = fx.build(self.dir / "before.docx")

    def tearDown(self):
        self.tmp.cleanup()

    def verify(self, after, *args, before=None):
        r = subprocess.run(
            [sys.executable, str(SCRIPT), "--before", str(before or self.before), "--after", str(after), "--json", *args],
            capture_output=True,
            text=True,
        )
        return r.returncode, (json.loads(r.stdout) if r.stdout.strip() else None), r.stderr

    def test_clean_docx_passes(self):
        after = fx.build(self.dir / "after.docx", **fx.clean_marks())
        code, rep, err = self.verify(after)
        self.assertEqual(code, 0, err)
        self.assertEqual(rep["total"], 0)
        self.assertTrue(all(v == 0 for v in rep["remaining"].values()), rep["remaining"])
        self.assertEqual([rep["body_equal"], rep["markup_equal"]], [True, True])
        self.assertGreaterEqual(len(rep["remaining"]), 10)

    def test_each_mark_left_in_is_counted_under_its_own_key(self):
        self.assertGreaterEqual(len(LEFT_IN), 10)
        for key, marks in LEFT_IN.items():
            with self.subTest(key=key):
                after = fx.build(self.dir / "left.docx", **{**fx.clean_marks(), **marks})
                code, rep, err = self.verify(after)
                self.assertEqual(code, 1, err)
                self.assertEqual(rep["remaining"][key], 1, rep["remaining"])
                self.assertEqual(rep["total"], 1, rep["remaining"])
                self.assertTrue(rep["body_equal"])

    def test_tracked_change_authors_are_counted_per_mark(self):
        # document.xml carries three marks with an author: w:ins, w:del, w:rPrChange.
        after = fx.build(self.dir / "left.docx", **{**fx.clean_marks(), "change_author": "Test Author"})
        code, rep, _ = self.verify(after)
        self.assertEqual(code, 1)
        self.assertEqual(rep["remaining"]["word/document.xml:change_author"], 3)

    def test_the_dirty_fixture_is_dirty_everywhere(self):
        code, rep, _ = self.verify(self.before)
        self.assertEqual(code, 1)
        zero = [k for k, v in rep["remaining"].items() if v == 0]
        self.assertEqual(zero, [])

    def test_a_changed_body_is_seen(self):
        after = fx.build(self.dir / "after.docx", body=fx.BODY.replace("plateau", "plateaux"), **fx.clean_marks())
        code, rep, _ = self.verify(after)
        self.assertEqual([code, rep["body_equal"], rep["total"]], [1, False, 0])

    def test_a_lost_revision_mark_is_seen(self):
        before = fx.build(self.dir / "b2.docx", extra_ins=True)
        after = fx.build(self.dir / "after.docx", **fx.clean_marks())
        code, rep, _ = self.verify(after, before=before)
        self.assertEqual([code, rep["markup_equal"]], [1, False])

    def test_another_anonymous_name_is_the_only_one_allowed(self):
        after = fx.build(self.dir / "after.docx", **fx.clean_marks("Reviewer Copy", "RC"))
        self.assertEqual(self.verify(after)[0], 1)
        code, rep, err = self.verify(after, "--author", "Reviewer Copy", "--initials", "RC")
        self.assertEqual([code, rep["total"]], [0, 0], err)

    def test_missing_optional_parts_count_as_clean(self):
        gone = {"word/people.xml", "word/comments.xml", "docProps/custom.xml", "docProps/app.xml", "word/footnotes.xml"}
        before = fx.build(self.dir / "b3.docx", parts=gone)
        after = fx.build(self.dir / "after.docx", parts=gone, **fx.clean_marks())
        code, rep, err = self.verify(after, before=before)
        self.assertEqual([code, rep["total"]], [0, 0], err)

    def test_another_prefix_for_the_word_namespace_is_still_read(self):
        after = fx.build(self.dir / "after.docx", prefix="x", **{**fx.clean_marks(), "change_author": "Test Author"})
        code, rep, _ = self.verify(after)
        self.assertEqual(rep["remaining"]["word/document.xml:change_author"], 3)
        self.assertTrue(rep["body_equal"])

    def test_a_part_that_is_not_xml_is_not_read_and_a_broken_xml_part_fails(self):
        # Word stores a thumbnail next to the properties; it is not an XML part.
        picture = {"docProps/thumbnail.jpeg": b"\xff\xd8\xff\xe0 not xml", "word/media/image1.png": b"\x89PNG"}
        before = fx.build(self.dir / "b4.docx", extra=picture)
        after = fx.build(self.dir / "after.docx", extra=picture, **fx.clean_marks())
        code, rep, err = self.verify(after, before=before)
        self.assertEqual([code, rep["unreadable_parts"]], [0, []], err)
        broken = fx.build(self.dir / "broken.docx", extra={"word/endnotes.xml": b"<w:endnotes>"}, **fx.clean_marks())
        code, rep, _ = self.verify(broken)
        self.assertEqual([code, rep["unreadable_parts"]], [1, ["word/endnotes.xml"]])

    def test_a_file_that_is_not_a_docx_is_an_error_not_a_pass(self):
        bad = self.dir / "bad.docx"
        bad.write_text("not a zip")
        code, rep, err = self.verify(bad)
        self.assertEqual(code, 2)
        self.assertIn("bad.docx", err)


if __name__ == "__main__":
    unittest.main()
