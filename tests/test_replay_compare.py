"""scripts/e2e/replay-compare.py on invented manuscripts built with the standard library."""
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "e2e" / "replay-compare.py"
spec = importlib.util.spec_from_file_location("replay_compare", SCRIPT)
rc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rc)

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)


def p(text):
    return '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text


def docx(path, body):
    xml = '<?xml version="1.0" encoding="UTF-8"?><w:document %s><w:body>%s<w:sectPr/></w:body></w:document>' % (NS, body)
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", xml)
    return path


INTRO = "The invented flasks were shaken at a fixed speed for two days in this placeholder study."
METHOD = "Samples were drawn every six hours and the optical density was read in a plate reader."
METHOD_NEW = "Samples were drawn every four hours and the optical density was read in a plate reader at 600 nm."
RESULT = "Growth began after a lag of one day and reached a plateau after four days in every repeat."
GONE = "This sentence about an invented side experiment was removed by the authors in the revision."
ADDED = "A new invented paragraph on the limits of the placeholder design was added by the authors."
OLD_LIMITS = "Only one flask was run in the invented design and the plate reader was not calibrated before the run."
NEW_LIMITS = ("Limits: one flask, and the plate reader was not calibrated. Every optical density of the "
              "invented design is therefore indicative, should not be compared between days, and says "
              "nothing about a second flask that was never run.")
CLOSING = "No real data are in this document, which exists for a software test only."
CLOSING_TOOL = "No real data are in this document, which exists only for a software test."

BEFORE = [INTRO, "", METHOD, GONE, RESULT, CLOSING]
AFTER = [INTRO, "", METHOD_NEW, RESULT, ADDED, CLOSING]


class Paragraphs(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_text_of_a_paragraph(self):
        body = (
            "<w:p><w:r><w:t>One</w:t><w:tab/><w:t>two</w:t><w:br/><w:t>three  </w:t></w:r>"
            '<w:del w:id="1"><w:r><w:delText>gone</w:delText></w:r></w:del>'
            '<w:ins w:id="2"><w:r><w:t> four</w:t></w:r></w:ins></w:p>'
            "<w:p/>"
            "<w:tbl><w:tr><w:tc>" + p("in a cell") + "</w:tc></w:tr></w:tbl>"
        )
        self.assertEqual(rc.paragraphs(docx(self.dir / "a.docx", body)), ["One two three four", "", "in a cell"])

    def test_a_text_box_is_read_once_and_is_not_part_of_its_paragraph(self):
        box = "<w:txbxContent>%s</w:txbxContent>" % p("in the box")
        body = (
            "<w:p><w:r><w:t>around</w:t></w:r><w:r><mc:AlternateContent>"
            "<mc:Choice><w:drawing>%s</w:drawing></mc:Choice>"
            "<mc:Fallback><w:pict>%s</w:pict></mc:Fallback>"
            "</mc:AlternateContent></w:r></w:p>" % (box, box)
        )
        self.assertEqual(rc.paragraphs(docx(self.dir / "b.docx", body)), ["around", "in the box"])


class Reference(unittest.TestCase):
    def test_modified_deleted_inserted_with_their_paragraph_numbers(self):
        rows = rc.reference_changes(BEFORE, AFTER)
        self.assertEqual(
            [(r["kind"], r["paragraph_index_before"], r["paragraph_index_after"]) for r in rows],
            [("modified", 2, 2), ("deleted", 3, None), ("inserted", None, 4)],
        )
        self.assertEqual((rows[0]["before"], rows[0]["after"]), (METHOD, METHOD_NEW))
        self.assertGreater(rows[0]["similarity"], 0.8)
        self.assertEqual((rows[1]["before"], rows[1]["after"]), (GONE, ""))
        self.assertEqual((rows[2]["before"], rows[2]["after"]), ("", ADDED))

    def test_equal_documents_have_no_change_and_empty_paragraphs_do_not_count(self):
        self.assertEqual(rc.reference_changes(BEFORE, list(BEFORE)), [])
        self.assertEqual(rc.reference_changes([INTRO, "", RESULT], [INTRO, RESULT, ""]), [])

    def test_a_paragraph_written_anew_is_one_row_of_its_own_kind(self):
        self.assertLess(rc.ratio(OLD_LIMITS, NEW_LIMITS), rc.PAIR_MIN)
        self.assertGreaterEqual(rc.word_ratio(OLD_LIMITS, NEW_LIMITS), rc.REWRITE_MIN)
        rows = rc.reference_changes([INTRO, OLD_LIMITS, GONE, RESULT], [INTRO, ADDED, NEW_LIMITS, RESULT])
        self.assertEqual(
            [(r["kind"], r["paragraph_index_before"], r["paragraph_index_after"]) for r in rows],
            [("rewritten", 1, 2), ("deleted", 2, None), ("inserted", None, 1)],
        )
        self.assertEqual((rows[0]["before"], rows[0]["after"]), (OLD_LIMITS, NEW_LIMITS))
        self.assertGreaterEqual(rows[0]["similarity"], rc.REWRITE_MIN)
        # Two paragraphs with nothing in common stay a deletion and an insertion.
        self.assertEqual([r["kind"] for r in rc.reference_changes([INTRO, GONE], [INTRO, ADDED])],
                         ["deleted", "inserted"])

    def test_a_rewritten_paragraph_next_to_a_new_one_is_paired_with_its_own(self):
        rows = rc.reference_changes([INTRO, METHOD, RESULT], [INTRO, ADDED, METHOD_NEW, RESULT])
        self.assertEqual(
            [(r["kind"], r["paragraph_index_before"], r["paragraph_index_after"]) for r in rows],
            [("modified", 1, 2), ("inserted", None, 1)],
        )


def change(cid, before, after, kind="replace"):
    return {"change_id": cid, "kind": kind, "state": "pending", "before": before, "after": after}


class Tool(unittest.TestCase):
    def test_changes_are_placed_by_text_and_chained_within_a_paragraph(self):
        step1 = METHOD.replace("six", "four")
        report = [
            {"ref": "R1.1", "changes": [change("R1.1-1", METHOD, step1)]},
            {"ref": "R2.1", "changes": [
                change("R2.1-1", step1, step1 + " Again."),
                change("R2.1-2", "  " + CLOSING.replace("document,", "document ,"), CLOSING_TOOL),
                change("R2.1-3", "A paragraph that is in neither file at all, about something else.", "x"),
                change("R2.1-4", "", "An inserted paragraph.", "insert"),
            ]},
        ]
        placed, unplaced = rc.place_tool_changes(report, BEFORE)
        self.assertEqual(sorted(placed), [2, 5])
        self.assertEqual(placed[2], {"before": METHOD, "after": step1 + " Again.",
                                     "refs": ["R1.1", "R2.1"], "change_ids": ["R1.1-1", "R2.1-1"]})
        # A text that is close to one paragraph is placed on it.
        self.assertEqual(placed[5]["after"], CLOSING_TOOL)
        self.assertEqual([u["change_id"] for u in unplaced], ["R2.1-3", "R2.1-4"])


class EndToEnd(unittest.TestCase):
    def test_the_two_files_are_written_and_the_screen_shows_counts_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            before = docx(d / "before.docx", "".join(p(t) if t else "<w:p/>" for t in BEFORE))
            after = docx(d / "after.docx", "".join(p(t) if t else "<w:p/>" for t in AFTER))
            tool = d / "changes.json"
            tool.write_text(json.dumps([
                {"ref": "R1.1", "changes": [change("R1.1-1", METHOD, METHOD.replace("six", "four"))]},
                {"ref": "R1.2", "changes": [change("R1.2-1", CLOSING, CLOSING_TOOL)]},
            ]), encoding="utf-8")
            out = d / "out"
            r = subprocess.run(
                [sys.executable, str(SCRIPT), "--before", str(before), "--after", str(after),
                 "--tool", str(tool), "--out", str(out)],
                capture_output=True, text=True)
            self.assertEqual(r.returncode, 0, r.stderr)
            counts = json.loads(r.stdout)
            self.assertEqual(
                {k: counts[k] for k in ("paragraphs_before", "paragraphs_after", "modified", "rewritten",
                                         "deleted", "inserted", "tool_paragraphs", "tool_unplaced", "both")},
                {"paragraphs_before": 5, "paragraphs_after": 5, "modified": 1, "rewritten": 0, "deleted": 1,
                 "inserted": 1, "tool_paragraphs": 2, "tool_unplaced": 0, "both": 1},
            )
            self.assertNotIn("flasks", r.stdout)
            ref = json.loads((out / "reference_changes.json").read_text(encoding="utf-8"))
            self.assertEqual([x["kind"] for x in ref], ["modified", "deleted", "inserted"])
            md = (out / "compare.md").read_text(encoding="utf-8")
            self.assertIn("- Changed by the authors: 1 modified, 0 rewritten, 1 deleted, 1 inserted", md)
            self.assertIn(
                "- Of the paragraphs the authors modified, rewrote or deleted, also changed by the tool: 1 of 2 (50%)", md)
            self.assertIn(
                "- The same, by what the authors did: modified 1 of 1 (100%), rewritten n/a, deleted 0 of 1 (0%)", md)
            self.assertIn("- Of the paragraphs the tool changed, also changed by the authors: 1 of 2 (50%)", md)
            both = md.split("## Changed by both")[1].split("## ")[0]
            self.assertIn("| 2 | R1.1 | modified | %s | %s | %s |"
                          % (METHOD, METHOD_NEW, METHOD.replace("six", "four")), both)
            only_ref = md.split("## Changed by the authors only")[1].split("## ")[0]
            self.assertIn("| 3 | deleted | %s | (empty) |" % GONE, only_ref)
            self.assertIn("| after 4 | inserted | (empty) | %s |" % ADDED, only_ref)
            only_tool = md.split("## Changed by the tool only")[1].split("## ")[0]
            self.assertIn("| 5 | R1.2 | %s | %s |" % (CLOSING, CLOSING_TOOL), only_tool)
            self.assertNotIn("—", md)
            # The inputs are as they were, and nothing is written next to them.
            self.assertEqual(sorted(x.name for x in d.iterdir()),
                             ["after.docx", "before.docx", "changes.json", "out"])

    def test_without_tool_changes_only_the_reference_is_written(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            before = docx(d / "before.docx", "".join(p(t) for t in BEFORE if t))
            after = docx(d / "after.docx", "".join(p(t) for t in AFTER if t))
            self.assertEqual(rc.main(["--before", str(before), "--after", str(after), "--out", str(d / "o")]), 0)
            md = (d / "o" / "compare.md").read_text(encoding="utf-8")
            self.assertIn("## Changed by the authors\n", md)
            self.assertNotIn("tool", md.split("## Counts")[1].lower())

    def test_the_script_imports_the_standard_library_only(self):
        src = SCRIPT.read_text(encoding="utf-8")
        for name in ("docx", "lxml"):
            self.assertNotRegex(src, r"(?m)^\s*(?:import|from) %s\b" % name)


if __name__ == "__main__":
    unittest.main()
