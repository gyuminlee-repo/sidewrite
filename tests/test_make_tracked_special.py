"""Known-answer check for scripts/e2e/make_tracked_special.py.

Run: python3 -m unittest tests/test_make_tracked_special.py
"""

import importlib.util
import re
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location(
    "special", ROOT / "scripts" / "e2e" / "make_tracked_special.py"
)
special = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(special)


class TrackedSpecial(unittest.TestCase):
    def test_one_tracked_run_per_case_with_the_special_elements(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = special.build(Path(tmp) / "s.docx")
            with zipfile.ZipFile(out) as z:
                xml = z.read("word/document.xml").decode("utf-8")
                settings = z.read("word/settings.xml").decode("utf-8")
        marks = len(re.findall(r"<w:ins ", xml)) + len(re.findall(r"<w:del ", xml))
        self.assertEqual(marks, len(special.CASES))
        self.assertGreaterEqual(len(special.CASES), 12)
        for tag in ("w:noBreakHyphen", "w:softHyphen", "w:tab", "w:br", "w:cr", "w:sym"):
            self.assertIn(f"<{tag}", xml, tag)
        self.assertIn("<w:delText", xml)
        self.assertRegex(settings, r'w:name="compatibilityMode"[^>]*w:val="15"')


if __name__ == "__main__":
    unittest.main()
