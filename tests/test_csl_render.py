"""Check tools/csl_render.py: the optional reference rendering with citeproc-py.

The tool is only active when the `citeproc` package can be imported. Without
it the tool must say so with exit code 3; that path is tested in every
environment by hiding the package. The rendering itself is tested only where
citeproc-py is installed. Run: python3 -m unittest tests/test_csl_render.py
"""

import contextlib
import importlib.util
import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
TOOL = ROOT / "tools" / "csl_render.py"
STYLE = ROOT / "tests" / "fixtures" / "csl_independent.csl"
DEPENDENT = ROOT / "tests" / "fixtures" / "csl_dependent.csl"
HAVE_CITEPROC = importlib.util.find_spec("citeproc") is not None

ITEMS = [{"id": "kim2019", "type": "article-journal", "title": "Growth of invented cells on methane",
          "author": [{"family": "Kim", "given": "A."}], "issued": {"date-parts": [[2019]]}}]


def load_tool():
    spec = importlib.util.spec_from_file_location("csl_render", TOOL)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def call(tool, argv):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = tool.main(argv)
    return code, out.getvalue(), err.getvalue()


class CslRender(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.items = Path(self.tmp.name) / "items.json"
        self.items.write_text(json.dumps(ITEMS), encoding="utf-8")

    def tearDown(self):
        self.tmp.cleanup()

    def test_without_citeproc_the_check_says_how_to_get_it_and_exits_3(self):
        tool = load_tool()
        with mock.patch.dict(sys.modules, {"citeproc": None}):  # makes `import citeproc` fail
            code, out, err = call(tool, ["--check"])
            self.assertEqual(code, 3)
            self.assertIn("pip install citeproc-py", err)
            self.assertEqual(out, "")
            # Rendering is refused the same way; it never half-works.
            code, out, err = call(tool, ["--style", str(STYLE), "--items", str(self.items)])
            self.assertEqual(code, 3)
            self.assertIn("pip install citeproc-py", err)

    def test_the_command_line_reports_what_this_machine_has(self):
        r = subprocess.run([sys.executable, str(TOOL), "--check"], capture_output=True, text=True)
        self.assertEqual(r.returncode, 0 if HAVE_CITEPROC else 3, r.stderr)
        if HAVE_CITEPROC:
            self.assertIn("citeproc-py is available", r.stdout)
        else:
            self.assertIn("pip install citeproc-py", r.stderr)

    def test_bad_arguments_are_errors_before_citeproc_is_needed(self):
        tool = load_tool()
        code, _, err = call(tool, ["--style", str(STYLE)])
        self.assertEqual(code, 2)
        self.assertIn("--items", err)
        code, _, err = call(tool, ["--style", str(Path(self.tmp.name) / "none.csl"), "--items", str(self.items)])
        self.assertEqual(code, 2)
        bad = Path(self.tmp.name) / "bad.json"
        bad.write_text("{not json", encoding="utf-8")
        self.assertEqual(call(tool, ["--style", str(STYLE), "--items", str(bad)])[0], 2)

    def test_a_dependent_style_cannot_be_rendered_and_names_its_parent(self):
        tool = load_tool()
        code, _, err = call(tool, ["--style", str(DEPENDENT), "--items", str(self.items)])
        self.assertEqual(code, 2)
        self.assertIn("invented-parent-style", err)

    @unittest.skipUnless(HAVE_CITEPROC, "citeproc-py is not installed")
    def test_one_reference_is_rendered_with_a_minimal_style(self):
        tool = load_tool()
        code, out, err = call(tool, ["--style", str(STYLE), "--items", str(self.items)])
        self.assertEqual(code, 0, err)
        rows = json.loads(out)["entries"]
        self.assertEqual([r["id"] for r in rows], ["kim2019"])
        self.assertIn("Growth of invented cells on methane", rows[0]["text"])


if __name__ == "__main__":
    unittest.main()
