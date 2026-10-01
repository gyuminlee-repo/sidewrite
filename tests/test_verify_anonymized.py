"""Known-answer check for scripts/verify-anonymized.py.

The verifier must tell a clean docx from one that keeps a single author mark,
target by target, and must see a changed body. Run:
python3 -m unittest tests/test_verify_anonymized.py
"""

import json
import os
import struct
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
        self.assertEqual([rep["body_equal"], rep["markup_equal"], rep["structure_ok"]], [True, True, True])
        self.assertGreaterEqual(len(rep["remaining"]), 10)
        for cid, c in rep["categories"].items():
            if c["default"] == "clear":
                self.assertEqual((c["status"], c["count"]), ("cleared", 0), cid)
        # The clean fixture still has a hidden run and a paragraph id: reported, and no reason to fail.
        self.assertEqual(rep["categories"]["hidden_text"]["status"], "not cleared")
        self.assertEqual(rep["categories"]["para_ids"]["status"], "kept by option")
        self.assertEqual(rep["not_cleared"], ["hidden_text"])

    def test_the_dirty_fixture_has_the_known_count_in_every_category(self):
        code, rep, _ = self.verify(self.before)
        self.assertEqual(code, 1)
        self.assertEqual(set(rep["categories"]), set(fx.EXPECTED_DIRTY))
        self.assertEqual({c: v["count"] for c, v in rep["categories"].items()}, fx.EXPECTED_DIRTY)
        self.assertEqual(len(rep["categories"]), 23)
        for cid, c in rep["categories"].items():
            self.assertEqual(c["status"], "kept by option" if c["default"] == "keep" else "not cleared", cid)
        clear = [c for c, v in rep["categories"].items() if v["default"] == "clear"]
        self.assertEqual(rep["total"], sum(fx.EXPECTED_DIRTY[c] for c in clear))

    def test_keep_and_also_move_a_category_between_kept_and_expected(self):
        clear = [c for c, v in self.verify(self.before)[1]["categories"].items() if v["default"] == "clear"]
        code, rep, err = self.verify(self.before, "--keep", ",".join(clear))
        self.assertEqual([code, rep["total"]], [0, 0], err)
        for cid in clear:
            self.assertEqual(rep["categories"][cid]["status"], "kept by option", cid)
        # Reported-only categories stay "not cleared" whatever is kept.
        self.assertEqual(rep["not_cleared"], ["hidden_text", "protection", "image_uninspected", "embedded"])
        code, rep, _ = self.verify(self.before, "--keep", ",".join(clear), "--also", "title,para_ids")
        self.assertEqual([code, rep["total"]], [1, fx.EXPECTED_DIRTY["title"] + fx.EXPECTED_DIRTY["para_ids"]])
        self.assertEqual(rep["categories"]["title"]["status"], "not cleared")
        for bad in ("hidden_text", "no_such_category"):
            r = subprocess.run([sys.executable, str(SCRIPT), "--before", str(self.before), "--after", str(self.before), "--keep", bad],
                               capture_output=True, text=True)
            self.assertEqual(r.returncode, 2, bad)

    def test_zip_comments_extra_fields_and_member_times_are_leaks_of_their_own(self):
        clean = fx.build(self.dir / "clean.docx", **fx.clean_marks())
        code, rep, err = self.verify(clean)
        self.assertEqual([code, rep["categories"]["zip_meta"]["count"]], [0, 0], err)
        leaks = {
            "zip:archive_comment": dict(comment=fx.ZIP_COMMENT),
            "zip:entry_comment": dict(entry_comment=fx.ENTRY_COMMENT),
            "zip:extra_field": dict(extra=fx.ZIP_EXTRA),
            "zip:timestamp": dict(date_time=(2026, 1, 2, 3, 4, 6)),
        }
        for key, how in leaks.items():
            with self.subTest(key=key):
                after = fx.rezip(clean, self.dir / "leak.docx", **how)
                code, rep, err = self.verify(after)
                self.assertEqual(code, 1, err)
                self.assertEqual(rep["remaining"][key], 1)
                self.assertEqual([rep["total"], rep["body_equal"], rep["markup_equal"]], [1, True, True])
                self.assertEqual(rep["categories"]["zip_meta"]["status"], "not cleared")
                self.assertIn("zip_meta", rep["not_cleared"])
                kept = self.verify(after, "--keep", "zip_meta")
                self.assertEqual([kept[0], kept[1]["categories"]["zip_meta"]["status"]], [0, "kept by option"])
        # Both comments at once, as in a file an archiver wrote: two leaks.
        both = fx.rezip(clean, self.dir / "both.docx", comment=fx.ZIP_COMMENT, entry_comment=fx.ENTRY_COMMENT)
        self.assertEqual(self.verify(both)[1]["categories"]["zip_meta"]["count"], 2)
        # The zip64 size field is not metadata about a person.
        zip64 = struct.pack("<HHQQ", 0x0001, 16, 10, 10)
        self.assertEqual(self.verify(fx.rezip(clean, self.dir / "z64.docx", extra=zip64))[1]["categories"]["zip_meta"]["count"], 0)

    def test_a_picture_that_cannot_be_read_counts_as_not_cleared(self):
        marks = fx.clean_marks()
        for name, data, n in (
            ("word/media/a.jpeg", fx.jpeg(with_exif=False, comment=b""), 0),
            ("word/media/a.jpeg", fx.jpeg(orientation=6), 1),
            ("word/media/a.JPG", fx.jpeg(with_exif=False), 1),  # a comment only
            ("word/media/a.jpeg", b"\xff\xd8 broken", 1),
            ("word/media/a.png", fx.png(text=False), 0),
            ("word/media/a.png", fx.png(), 1),
            ("word/media/a.png", b"\x89PNG\r\n\x1a\n cut", 1),
        ):
            with self.subTest(name=name, n=n):
                after = fx.build(self.dir / "pic.docx", extra={name: data}, **marks)
                code, rep, _ = self.verify(after)
                self.assertEqual([rep["categories"]["image_meta"]["count"], code], [n, 1 if n else 0])

    def test_a_reference_to_a_part_that_is_gone_fails_the_check(self):
        before = fx.build(self.dir / "b5.docx", **fx.clean_marks())
        rels = (
            b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Type="t" Target="gone.xml"/>'
            b'<Relationship Id="rId2" Type="t" Target="https://journal.example/" TargetMode="External"/></Relationships>'
        )
        after = fx.build(self.dir / "after.docx", extra={"word/_rels/document.xml.rels": rels}, **fx.clean_marks())
        code, rep, _ = self.verify(after, before=before)
        self.assertEqual([code, rep["structure_ok"], rep["total"]], [1, False, 0])
        self.assertEqual(rep["dangling"], ["word/_rels/document.xml.rels -> word/gone.xml"])
        # A reference that was already broken in the source is not held against the copy.
        self.assertEqual(self.verify(after, before=after)[0], 0)

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
        picture = {"docProps/thumbnail.jpeg": b"\xff\xd8\xff\xe0 not xml", "word/media/image9.png": fx.png(text=False)}
        before = fx.build(self.dir / "b4.docx", extra=picture)
        after = fx.build(self.dir / "after.docx", extra=picture, **fx.clean_marks())
        code, rep, err = self.verify(after, "--keep", "thumbnail", before=before)
        self.assertEqual([code, rep["unreadable_parts"]], [0, []], err)
        self.assertEqual(rep["categories"]["thumbnail"], {**rep["categories"]["thumbnail"], "status": "kept by option", "count": 1})
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
