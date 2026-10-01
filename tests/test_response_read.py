"""Tests for tools/response_read.py: read an existing response letter and
propose profile overrides from what it shows.

Run: python3 -m unittest tests/test_response_read.py
"""

import contextlib
import importlib.util
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import response_docx as rd  # noqa: E402
import response_read as rr  # noqa: E402

_spec = importlib.util.spec_from_file_location("mtr", ROOT / "scripts" / "make-test-response.py")
mtr = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mtr)

BODY, REVISED = "C00000", "0070C0"
ANSWERS = {
    "R1.1": "We now give the unit as degrees Celsius.",
    "R1.2": "We replaced the term in that paragraph.",
    "R1.3": "We state how growth was followed.",
    "R1.4": "We deleted the sentence.",
    "R1.5": "No antifoam agent was used in this study.",
}


def run_main(argv):
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        code = rr.main(argv)
    return code, json.loads(out.getvalue().strip().splitlines()[-1])


class VariantTest(unittest.TestCase):
    def read(self, variant):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        path = Path(self.tmp.name) / f"{variant}.docx"
        mtr.build(path, variant)
        ps = rr.paragraphs(path)
        return ps, rr.propose(ps, {})

    def colours(self, out):
        return (out["overrides"]["colors"]["response_body"], out["overrides"]["colors"]["response_revised"])

    def test_plain(self):
        ps, out = self.read("plain")
        self.assertEqual(self.colours(out), (BODY, REVISED))
        self.assertEqual(sorted(out["previous"]), ["R1.1", "R1.2", "R1.3", "R1.4", "R1.5"])
        self.assertEqual(out["previous"], ANSWERS)
        self.assertEqual(out["overrides"]["response_letter"]["location_line"], "[Revised in {section}, {lines}]")
        self.assertNotIn("response_kr", out["overrides"]["colors"])
        self.assertNotIn("highlight", out["overrides"]["colors"])
        self.assertFalse(any(p["in_table"] or p["tracked"] for p in ps))
        self.assertEqual(out["evidence"]["tracked_paragraphs"], 0)
        self.assertEqual(out["evidence"]["colors"][BODY]["role"], "response_body")
        self.assertEqual(out["evidence"]["colors"][REVISED]["role"], "response_revised")

    def test_table_cells_are_read_in_order(self):
        ps, out = self.read("table")
        self.assertTrue(any(p["in_table"] for p in ps))
        self.assertEqual(self.colours(out), (BODY, REVISED))
        self.assertEqual(len(out["previous"]), 5)
        self.assertEqual(out["previous"], ANSWERS)
        texts = [p["text"] for p in ps]
        first = texts.index("Comment 1: Please state the unit of the incubation temperature as degrees Celsius.")
        self.assertEqual(texts[first + 1:first + 4], [
            "We now give the unit as degrees Celsius.",
            "[Revised in 2. Materials and Methods, line 9]",
            "The incubation temperature was held at 30 °C.",
        ])
        # An empty paragraph in a cell stays, with no text and no colour.
        self.assertTrue(any(p["in_table"] and p["text"] == "" and p["color"] == "NONE" for p in ps))
        self.assertFalse(ps[0]["in_table"])

    def test_tracked_runs_read_as_accepted(self):
        ps, out = self.read("tracked")
        hit = [p for p in ps if p["tracked"]]
        self.assertEqual(len(hit), 1)
        self.assertEqual(hit[0]["text"], "We added the unit as degrees Celsius.")
        self.assertIn("added ", hit[0]["text"])
        self.assertNotIn("removed ", hit[0]["text"])
        self.assertNotIn("moved here ", hit[0]["text"])
        self.assertEqual(hit[0]["color"], BODY)
        self.assertEqual(out["evidence"]["tracked_paragraphs"], 1)
        self.assertEqual(out["previous"]["R1.1"], "We added the unit as degrees Celsius.")

    def test_colour_from_character_style_paragraph_style_and_theme(self):
        ps, out = self.read("styled")
        self.assertEqual(self.colours(out), (BODY, REVISED))
        self.assertEqual(out["evidence"]["colors"][REVISED]["theme"], ["accent1"])
        self.assertTrue(any("theme colour" in n for n in out["evidence"]["notes"]))
        self.assertTrue(any(p["color"] == "NONE" and p["text"] for p in ps))   # the auto paragraph
        by_text = {p["text"]: p for p in ps}
        # A character style that only has a parent takes the parent's colour.
        self.assertEqual(by_text["We deleted the sentence."]["color"], BODY)
        # A paragraph style colours runs that carry none of their own.
        self.assertEqual(by_text["[Removed from 1. Introduction]"]["color"], REVISED)
        # A direct "auto" wins over the character style of the same run.
        self.assertEqual(by_text["[No change made]"]["color"], "NONE")
        self.assertEqual(by_text["Response to Reviewers"]["color"], "NONE")
        self.assertEqual(out["previous"], ANSWERS)

    def test_east_asian_runs(self):
        ps, out = self.read("eastasian")
        self.assertTrue(any("단위를 추가했다." in p["text"] for p in ps))
        self.assertEqual(out["overrides"]["colors"]["response_kr"], "7F7F7F")
        self.assertEqual(out["overrides"]["response_letter"]["gloss_prefix"], "[T] ")
        gloss = next(p for p in ps if p["text"] == "[T] 단위를 추가했다.")
        self.assertEqual(gloss["colors"], {"7F7F7F": len("[T] 단위를 추가했다.")})   # code points
        mixed = next(p for p in ps if "광학 밀도" in p["text"])
        self.assertEqual(mixed["text"], "We state how growth was followed (광학 밀도).")
        self.assertEqual(mixed["color"], BODY)
        # The translation is not part of the earlier answer.
        self.assertEqual(out["previous"]["R1.1"], ANSWERS["R1.1"])

    def test_highlight_and_shading_are_not_text_colours(self):
        ps, out = self.read("highlight")
        self.assertEqual(self.colours(out), (BODY, REVISED))
        self.assertEqual(out["overrides"]["colors"]["highlight"], "yellow")
        self.assertIn("FFFF00", out["evidence"]["shades"])
        self.assertNotIn("FFFF00", out["evidence"]["colors"])
        quote = next(p for p in ps if "Fig. 2" in p["text"])
        self.assertEqual(quote["highlights"], {"yellow": 6})
        self.assertEqual(quote["color"], REVISED)
        shaded = next(p for p in ps if p["shades"])
        self.assertEqual(shaded["shades"], {"FFFF00": len("the term")})

    def test_not_a_docx_is_an_error_not_a_crash(self):
        bad = Path(tempfile.gettempdir()) / "not.docx"
        bad.write_text("x")
        self.addCleanup(bad.unlink)
        code, out = run_main(["--docx", str(bad)])
        self.assertEqual((code, out["ok"]), (1, False))
        self.assertTrue(out["error"])

    def test_zip_without_a_document_part_and_broken_xml_are_errors(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        empty = Path(tmp.name) / "empty.docx"
        with zipfile.ZipFile(empty, "w") as z:
            z.writestr("other.xml", "<a/>")
        broken = Path(tmp.name) / "broken.docx"
        with zipfile.ZipFile(broken, "w") as z:
            z.writestr("word/document.xml", "<w:document")
        entity = Path(tmp.name) / "entity.docx"
        with zipfile.ZipFile(entity, "w") as z:
            z.writestr("word/document.xml", '<!DOCTYPE d [<!ENTITY a "aaaa">]><d>&a;</d>')
        for path in (empty, broken, entity, Path(tmp.name) / "missing.docx"):
            code, out = run_main(["--docx", str(path)])
            self.assertEqual((code, out["ok"]), (1, False), path.name)

    def test_cli_prints_the_proposal(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "plain.docx"
        mtr.build(path, "plain")
        code, out = run_main(["--docx", str(path)])
        self.assertEqual((code, out["ok"]), (0, True))
        self.assertEqual(out["overrides"]["colors"], {"response_body": BODY, "response_revised": REVISED})
        self.assertEqual(out["previous"], ANSWERS)


PRIOR = {"paragraphs": [
    {"kind": "x", "text": "Response to Reviewers", "color": None},
    {"kind": "x", "text": "Reviewer #2", "color": None},
    {"kind": "x", "text": "Comment 1: The unit is missing.", "color": None},
    {"kind": "x", "text": "We added the unit.", "color": "C00000"},
    {"kind": "x", "text": "[Changed in Section 2.1, lines 40-41]", "color": "0070C0"},
    {"kind": "x", "text": "The temperature was 30 °C.", "color": "0070C0"},
    {"kind": "x", "text": "[T] 단위를 추가했다.", "color": "7F7F7F"},
    {"kind": "x", "text": "Comment 2: Explain the control.", "color": None},
    {"kind": "x", "text": "The control is described now.", "color": "C00000"}]}


class ReadTest(unittest.TestCase):
    """Round trip through the letter writer."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / "prior.docx"
        rd.write_docx(PRIOR, self.path)

    def test_paragraph_colours_by_characters(self):
        ps = rr.paragraphs(self.path)
        self.assertEqual([p["color"] for p in ps],
                         ["NONE", "NONE", "NONE", "C00000", "0070C0", "0070C0", "7F7F7F", "NONE", "C00000"])
        self.assertEqual([p["text"] for p in ps], [p["text"] for p in PRIOR["paragraphs"]])

    def test_run_with_attributes_keeps_its_colour(self):
        x = zipfile.ZipFile(self.path).read("word/document.xml").decode("utf8").replace("<w:r>", '<w:r w:rsidRPr="00AB12CD">')
        self.assertEqual(rr.paragraphs_from_xml(x)[3]["color"], "C00000")

    def test_the_colour_with_most_characters_names_the_paragraph(self):
        x = ('<w:document xmlns:w="%s"><w:body><w:p>'
             '<w:r><w:rPr><w:color w:val="111111"/></w:rPr><w:t>ab</w:t></w:r>'
             '<w:r><w:rPr><w:color w:val="222222"/></w:rPr><w:t>cdefg</w:t></w:r>'
             '<w:r><w:t>hi</w:t></w:r>'
             '</w:p></w:body></w:document>' % rd.W)
        p = rr.paragraphs_from_xml(x)[0]
        self.assertEqual((p["text"], p["color"], p["colors"]), ("abcdefghi", "222222", {"111111": 2, "222222": 5, "NONE": 2}))

    def test_proposal_is_an_override_with_evidence(self):
        out = rr.propose(rr.paragraphs(self.path), {})
        self.assertEqual(out["overrides"]["colors"], {"response_body": "C00000", "response_revised": "0070C0", "response_kr": "7F7F7F"})
        self.assertEqual(out["overrides"]["response_letter"]["location_line"], "[Changed in {section}, {lines}]")
        self.assertEqual(out["overrides"]["response_letter"]["gloss_prefix"], "[T] ")
        self.assertEqual(out["evidence"]["colors"]["C00000"]["paragraphs"], 2)
        self.assertEqual(out["previous"], {"R2.1": "We added the unit.", "R2.2": "The control is described now."})

    def test_letter_without_colours_proposes_none(self):
        plain = {"paragraphs": [dict(p, color=None) for p in PRIOR["paragraphs"]]}
        rd.write_docx(plain, self.path)
        out = rr.propose(rr.paragraphs(self.path), {})
        self.assertEqual(out["overrides"].get("colors"), None)
        self.assertEqual(out["previous"], {})
        self.assertTrue(any("no coloured" in n for n in out["evidence"]["notes"]))

    def test_an_editor_block_is_read_as_e_not_as_the_first_reviewer(self):
        row = lambda text, color=None: {"kind": "x", "text": text, "color": color}  # noqa: E731
        letter = {"paragraphs": [
            row("Editor"), row("Comment 1"), row("Shorten the title."), row("The title is at the limit.", "C00000"),
            row("Reviewer #1"), row("Comment 1"), row("State the unit."), row("We now state the unit.", "C00000")]}
        rd.write_docx(letter, self.path)
        ps = rr.paragraphs(self.path)
        out = rr.propose(ps, {})
        self.assertEqual(out["previous"], {"E.1": "The title is at the limit.", "R1.1": "We now state the unit."})
        self.assertEqual(out["evidence"]["notes"], [])
        # The label of the profile, when the letter uses another word.
        letter["paragraphs"][0]["text"] = "Handling editor:"
        rd.write_docx(letter, self.path)
        ps = rr.paragraphs(self.path)
        self.assertIn("R1.1 occurs twice; the first answer is kept", rr.propose(ps, {})["evidence"]["notes"])
        out = rr.propose(ps, {"response_letter": {"editor_label": "Handling Editor"}})
        self.assertEqual(sorted(out["previous"]), ["E.1", "R1.1"])
        # A sentence that starts with the word is no heading.
        letter["paragraphs"][0]["text"] = "Reviewer #2"
        letter["paragraphs"][6]["text"] = "Editor comments on the unit were the same."
        rd.write_docx(letter, self.path)
        self.assertEqual(sorted(rr.propose(rr.paragraphs(self.path), {})["previous"]), ["R1.1", "R2.1"])

    def test_head_patterns_of_the_profile_are_used_when_they_fit_the_letter(self):
        letter = {"paragraphs": [
            {"kind": "x", "text": "Referee report 3", "color": None},
            {"kind": "x", "text": "Issue 4. Too long.", "color": None},
            {"kind": "x", "text": "Shortened.", "color": "C00000"}]}
        rd.write_docx(letter, self.path)
        ps = rr.paragraphs(self.path)
        self.assertEqual(rr.propose(ps, {})["previous"], {})
        profile = {"reviewer_head": r"^Referee report (\d+)", "comment_head": r"^Issue (\d+)\."}
        self.assertEqual(rr.propose(ps, profile)["previous"], {"R3.4": "Shortened."})
        # A profile pattern that fits nothing falls back to the built-in candidates.
        out = rr.propose(rr.paragraphs_from_xml(zipfile.ZipFile(self.path).read("word/document.xml").decode("utf8")),
                         {"comment_head": "^Nothing like this (\\d+)", "reviewer_head": "("})
        self.assertEqual(out["previous"], {})


if __name__ == "__main__":
    unittest.main()
