"""The marked-diff tool keeps the structure of both documents, not only their text.

For every invented baseline/clean pair of tests/marked_diff_corpus.py and both modes: the tool writes a file,
and that file, read with the corpus module's own reader, gives the clean structure when every mark is accepted
and the baseline structure when every mark is rejected. A pair the tool cannot handle must be refused with a
reason (exit 2); a failed check (exit 1) or a wrong file is never accepted.

Run: python3 -m unittest tests/test_marked_diff_roundtrip.py
     python3 tests/test_marked_diff_roundtrip.py --list     (one line per case and mode, no assertion)
"""

import contextlib
import importlib.util
import io
import re
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import marked_diff_corpus as corpus  # noqa: E402

TOOL = Path(__file__).resolve().parent.parent / "tools" / "build_marked_diff.py"
MODES = ("color", "track_changes")
MIXED = 60


def load_tool():
    spec = importlib.util.spec_from_file_location("build_marked_diff", TOOL)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def all_cases():
    cases = dict(corpus.CASES)
    cases.update(corpus.mixed_cases(MIXED))
    return cases


def first_difference(got, want):
    if len(got) != len(want):
        kinds = lambda xs: "".join(x[0] for x in xs)  # noqa: E731
        return "%d items (%s) where %d (%s) are expected" % (len(got), kinds(got), len(want), kinds(want))
    for i, (g, w) in enumerate(zip(got, want)):
        if g != w:
            return "item %d: %r where %r is expected" % (i, g, w)
    return None


def judge(tool, case, mode, folder):
    """None when the pair passes, else one line that says what is wrong."""
    base, clean, refusal = corpus.models(case)
    base_path = corpus.write(base, folder / "base.docx", header="An invented running head")
    clean_path = corpus.write(clean, folder / "clean.docx", header="An invented running head")
    out = folder / ("out_%s.docx" % mode)
    if out.exists():
        out.unlink()
    log = io.StringIO()
    with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
        code = tool.main(["--baseline", str(base_path), "--clean", str(clean_path), "--out", str(out), "--mode", mode])
    last = (log.getvalue().strip().splitlines() or [""])[-1]
    if refusal:
        if code == 2 and re.search(refusal, log.getvalue()) and not out.exists():
            return None
        return "expected a refusal (%s), got exit %d: %s" % (refusal, code, last[:160])
    if code != 0:
        fails = [line for line in log.getvalue().splitlines() if line.startswith("FAIL")]
        return "exit %d: %s" % (code, (fails or [last])[0][:200])
    want_clean, marks_clean = corpus.structure_of(clean)
    want_base, _ = corpus.structure_of(base)
    got, marks = corpus.structure_read(corpus.view(out, mode, "accept"))
    diff = first_difference(got, want_clean)
    if diff:
        return "accept-all is not the clean document: " + diff
    if marks != marks_clean:
        return "accept-all bookmarks: %r where %r is expected" % (marks, marks_clean)
    got, _ = corpus.structure_read(corpus.view(out, mode, "reject"))
    if mode == "color":
        got, want_base = corpus.without_page_breaks(got), corpus.without_page_breaks(want_base)
    for name, g, w in zip(("paragraphs", "tables"), corpus.split_reject(got), corpus.split_reject(want_base)):
        diff = first_difference(g, w)
        if diff:
            return "reject-all is not the baseline (%s): %s" % (name, diff)
    return None


class RoundTrip(unittest.TestCase):
    def test_every_pair_keeps_both_structures_in_both_modes(self):
        tool = load_tool()
        cases = all_cases()
        self.assertGreaterEqual(len(cases), 150)
        with tempfile.TemporaryDirectory() as tmp:
            for name, case in cases.items():
                for mode in MODES:
                    with self.subTest(case=name, mode=mode):
                        self.assertIsNone(judge(tool, case, mode, Path(tmp)))

    def test_the_reader_tells_the_two_sides_apart(self):
        # The reader is not allowed to agree with everything: a plain copy of the clean file read as "reject-all"
        # must differ from a baseline that differs.
        base, clean, _ = corpus.models(corpus.CASES["a paragraph added in the middle"])
        with tempfile.TemporaryDirectory() as tmp:
            path = corpus.write(clean, Path(tmp) / "clean.docx")
            got, _ = corpus.structure_read(corpus.view(path, "track_changes", "reject"))
        self.assertEqual(got, corpus.structure_of(clean)[0])
        self.assertIsNotNone(first_difference(got, corpus.structure_of(base)[0]))


if __name__ == "__main__":
    if "--list" in sys.argv:
        tool = load_tool()
        bad = 0
        with tempfile.TemporaryDirectory() as tmp:
            for name, case in all_cases().items():
                for mode in MODES:
                    verdict = judge(tool, case, mode, Path(tmp))
                    bad += verdict is not None
                    print("%-4s %-13s %s%s" % ("ok" if verdict is None else "BAD", mode, name, "" if verdict is None else " :: " + verdict))
        print("%d of %d bad" % (bad, 2 * len(all_cases())))
    else:
        unittest.main()
