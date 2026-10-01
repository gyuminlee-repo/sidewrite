"""Known-answer check for scripts/e2e/embed_taskpane.py.

Run: python3 -m unittest tests/test_embed_taskpane.py
"""

import hashlib
import importlib.util
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


if __name__ == "__main__":
    unittest.main()
