"""Known-answer check for scripts/e2e/embed_taskpane.py.

Run: python3 -m unittest tests/test_embed_taskpane.py
"""

import contextlib
import hashlib
import importlib.util
import io
import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location("embed", ROOT / "scripts" / "e2e" / "embed_taskpane.py")
embed = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(embed)

MANIFEST = ROOT / "manifests" / "word.xml"
SRC = ROOT / "tests" / "fixtures" / "review_manuscript.docx"


class Embed(unittest.TestCase):
    def test_manifest_id_version(self):
        self.assertEqual(
            embed.manifest_id_version(MANIFEST.read_text(encoding="utf-8")),
            ("27ee7761-18ca-4fae-b482-4171d5afa6cd", "0.1.2.0"),
        )
        with self.assertRaises(ValueError):
            embed.manifest_id_version("<OfficeApp><Id>not-a-guid</Id></OfficeApp>")

    def test_root_rel_added_once_with_a_new_id(self):
        rels = '<Relationships><Relationship Id="rId1"/><Relationship Id="rId7"/></Relationships>'
        once = embed.add_root_rel(rels)
        self.assertIn(f'Id="rId8" Type="{embed.REL_TASKPANES}"', once)
        self.assertEqual(embed.add_root_rel(once), once)

    def test_copy_has_parts_and_source_is_untouched(self):
        digest = hashlib.sha256(SRC.read_bytes()).hexdigest()
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "tagged.docx"
            embed.embed(SRC, out, "27ee7761-18ca-4fae-b482-4171d5afa6cd", "0.1.2.0")
            with zipfile.ZipFile(out) as z:
                we = z.read(embed.WE_PART).decode()
                tp = z.read(embed.TP_PART).decode()
                ct = z.read("[Content_Types].xml").decode()
                rels = z.read("_rels/.rels").decode()
                self.assertIn('store="developer" storeType="Registry"', we)
                self.assertIn('name="Office.AutoShowTaskpaneWithDocument" value="true"', we)
                self.assertIn('visibility="1"', tp)
                self.assertIn(embed.CT_WEBEXT, ct)
                self.assertIn(embed.CT_TASKPANES, ct)
                self.assertIn(embed.TP_PART, rels)
                with zipfile.ZipFile(SRC) as s:
                    self.assertEqual(len(z.namelist()), len(s.namelist()) + 3)
            with self.assertRaises(ValueError):
                embed.embed(out, Path(tmp) / "again.docx", "x", "1")
        self.assertEqual(hashlib.sha256(SRC.read_bytes()).hexdigest(), digest)


class SourceIsNeverWritten(unittest.TestCase):
    """Codex review 16: --out /path/./original.docx passed a string compare
    with --src, the output zip was opened for writing and truncated the source
    (38,766 bytes became 22). The output is refused when it is the source by
    any name, and it is written to a temp file first."""

    ARGS = ("27ee7761-18ca-4fae-b482-4171d5afa6cd", "0.1.2.0")

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.src = self.dir / "original.docx"
        shutil.copyfile(SRC, self.src)
        self.digest = hashlib.sha256(self.src.read_bytes()).hexdigest()

    def tearDown(self):
        self.tmp.cleanup()

    def assert_source_intact(self):
        self.assertEqual(hashlib.sha256(self.src.read_bytes()).hexdigest(), self.digest)
        self.assertEqual([p.name for p in self.dir.glob("*.tmp")], [])

    def refused(self, out):
        with self.assertRaises(ValueError):
            embed.embed(self.src, out, *self.ARGS)
        self.assert_source_intact()
        # The command line refuses it too, with exit status 2.
        with self.assertRaises(SystemExit) as cm, contextlib.redirect_stderr(io.StringIO()):
            embed.main(["--manifest", str(MANIFEST), "--src", str(self.src), "--out", str(out)])
        self.assertEqual(cm.exception.code, 2)
        self.assert_source_intact()

    def test_same_file_by_another_spelling(self):
        self.refused(f"{self.dir}/./original.docx")
        self.refused(f"{self.dir}/sub/../original.docx")
        self.refused(os.path.relpath(self.src))

    def test_symlink_to_the_source(self):
        link = self.dir / "link.docx"
        os.symlink(self.src, link)
        self.refused(link)

    def test_hardlink_to_the_source(self):
        link = self.dir / "hard.docx"
        os.link(self.src, link)
        self.refused(link)

    def test_a_failing_write_leaves_the_previous_output(self):
        out = self.dir / "tagged.docx"
        embed.embed(self.src, out, *self.ARGS)
        before = hashlib.sha256(out.read_bytes()).hexdigest()
        # The output already has the parts, so making it again from itself
        # fails; and a write that fails half way.
        with self.assertRaises(ValueError):
            embed.embed(out, self.dir / "again.docx", *self.ARGS)
        self.assertFalse((self.dir / "again.docx").exists())
        original = embed.taskpanes_xml
        embed.taskpanes_xml = lambda *a, **k: (_ for _ in ()).throw(OSError("disk full"))
        try:
            with self.assertRaises(OSError):
                embed.embed(self.src, out, *self.ARGS)
        finally:
            embed.taskpanes_xml = original
        self.assertEqual(hashlib.sha256(out.read_bytes()).hexdigest(), before)
        self.assert_source_intact()
        # And it still works afterwards.
        embed.embed(self.src, out, *self.ARGS)
        with zipfile.ZipFile(out) as z:
            self.assertIn(embed.TP_PART, z.namelist())


if __name__ == "__main__":
    unittest.main()
