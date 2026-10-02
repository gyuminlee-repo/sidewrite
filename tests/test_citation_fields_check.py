"""tools/citation_fields_check.py on the citation fixture and on copies of it
that are broken by hand. Standard library only: the fixture is committed."""
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOL = ROOT / "tools" / "citation_fields_check.py"
FIXTURE = ROOT / "tests" / "fixtures" / "citation_manuscript.docx"
EMBED = ROOT / "scripts" / "e2e" / "embed_taskpane.py"
DOC = "word/document.xml"
END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
BEGIN = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
SEPARATE = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
DEL = '<w:del w:id="900" w:author="Hand" w:date="2026-01-01T00:00:00Z">'

spec = importlib.util.spec_from_file_location("citation_fields_check", TOOL)
tool = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tool)


def deleted(xml):
    """Runs as Word writes them inside a tracked deletion."""
    return (
        DEL
        + xml.replace("w:instrText", "w:delInstrText").replace("<w:t ", "<w:delText ").replace(
            "</w:t>", "</w:delText>"
        )
        + "</w:del>"
    )


class Check(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.dir = Path(cls.tmp.name)
        with zipfile.ZipFile(FIXTURE) as z:
            cls.xml = z.read(DOC).decode("utf-8")
        # The first field of the document: the citation (Arden, 2019).
        a = cls.xml.index(BEGIN)
        b = cls.xml.index(SEPARATE, a)
        c = cls.xml.index(END, b) + len(END)
        cls.head, cls.code, cls.tail = cls.xml[:a], cls.xml[a:b], cls.xml[b:c]
        cls.rest = cls.xml[c:]

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def copy(self, name, edit=None, part=DOC):
        out = self.dir / name
        with zipfile.ZipFile(FIXTURE) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if edit and item.filename == part:
                    text = edit(data.decode("utf-8"))
                    self.assertNotEqual(text, data.decode("utf-8"), "the edit changed nothing")
                    data = text.encode("utf-8")
                z.writestr(item, data)
        return out

    def run_tool(self, after, *more, before=FIXTURE):
        p = subprocess.run(
            [sys.executable, str(TOOL), "--before", str(before), "--after", str(after)]
            + ["--detail", *more],
            capture_output=True,
            text=True,
        )
        lines = p.stdout.strip().split("\n")
        return p.returncode, json.loads(lines[-1]), lines[:-1]

    def test_without_detail_the_json_line_has_the_counts_only(self):
        p = subprocess.run(
            [sys.executable, str(TOOL), "--before", str(FIXTURE), "--after", str(FIXTURE)],
            capture_output=True,
            text=True,
        )
        out = json.loads(p.stdout.strip())
        self.assertEqual((p.returncode, out["ok"], out["after"]["fields"]), (0, True, 18))
        self.assertNotIn("paragraphs", out)
        self.assertNotIn("list", out["after"])

    def test_the_fixture_against_an_untouched_copy_passes(self):
        code, out, lines = self.run_tool(self.copy("same.docx"))
        self.assertEqual((code, out["ok"], out["failures"], lines), (0, True, [], []))
        self.assertEqual(out["before"]["fields"], out["after"]["fields"])
        self.assertEqual(out["after"]["by_state"], {"live": out["after"]["fields"]})

    def test_what_the_fixture_holds(self):
        _code, out, _ = self.run_tool(FIXTURE)
        self.assertEqual(
            out["before"]["by_kind"],
            {
                "EN.CITE": 2,
                "EN.CITE.DATA": 1,
                "EN.REFLIST": 1,
                "REF": 1,
                "ZOTERO_BIBL": 1,
                "ZOTERO_ITEM": 11,
                "hyperlink": 1,
            },
        )
        self.assertEqual(out["fld_data"]["before"], 2)
        self.assertEqual(out["properties"]["before"], 2)
        self.assertEqual(out["doc_vars"]["before"], 3)
        # The two lists are one field each over several paragraphs: open at
        # the end of their first paragraph, closed in their last.
        spans = [p for p in out["paragraphs"] if p["depth_start"] or p["depth_end"]]
        self.assertEqual(len(spans), 8)
        self.assertEqual(out["text_changed"], [])

    def test_a_dropped_end_fails_and_names_the_paragraph(self):
        code, out, lines = self.run_tool(self.copy("no-end.docx", lambda x: x.replace(END, "", 1)))
        self.assertEqual(code, 1)
        self.assertIn("unbalanced-after", out["failures"])
        self.assertTrue(out["unbalanced"])
        self.assertTrue(any(l.startswith("FAIL unbalanced (after)") for l in lines), lines)
        self.assertTrue(any("kind=ZOTERO_ITEM" in l for l in lines), lines)

    def test_a_dropped_begin_fails(self):
        code, out, _ = self.run_tool(self.copy("no-begin.docx", lambda x: x.replace(BEGIN, "", 1)))
        self.assertEqual(code, 1)
        self.assertIn("unbalanced-after", out["failures"])

    def test_one_changed_byte_of_an_instruction_fails(self):
        code, out, lines = self.run_tool(
            self.copy("byte.docx", lambda x: x.replace("cFx00001", "cFx0000X", 1))
        )
        self.assertEqual(code, 1)
        self.assertEqual(
            out["failures"], ["instruction-added-or-changed", "instruction-missing-or-changed"]
        )
        self.assertEqual((len(out["missing"]), len(out["added"])), (1, 1))
        row = next(l for l in lines if l.startswith("FAIL instruction-missing"))
        self.assertIn('"Shaken flasks remain the usual first ves"', row)
        self.assertIn("kind=ZOTERO_ITEM", row)
        self.assertIn("before: begin@- instr@- separate@- result@- end@-", row)

    def test_a_field_half_inside_a_deletion_fails_with_both_layouts(self):
        half = self.head + deleted(self.code) + self.tail + self.rest
        code, out, lines = self.run_tool(self.copy("half.docx", lambda _x: half))
        self.assertEqual(code, 1)
        self.assertEqual(out["failures"], ["split-field"])
        self.assertEqual(out["split"][0]["kind"], "ZOTERO_ITEM")
        self.assertEqual(
            lines,
            [
                'FAIL split-field | word/document.xml p3 "Shaken flasks remain the usual first ves" '
                "| kind=ZOTERO_ITEM | before: begin@- instr@- separate@- result@- end@- "
                "| after: begin@del instr@del separate@- result@- end@-"
            ],
        )

    def test_a_field_whole_inside_a_deletion_passes_and_counts_as_deleted(self):
        whole = self.head + deleted(self.code + self.tail) + self.rest
        code, out, _ = self.run_tool(self.copy("whole.docx", lambda _x: whole))
        self.assertEqual((code, out["failures"]), (0, []))
        self.assertEqual(out["after"]["by_state"]["deleted"], 1)
        self.assertEqual(out["after"]["list"][0]["state"], "deleted")
        self.assertEqual(out["flattened"], [])

    def test_a_deleted_field_whose_text_was_typed_again_is_listed_as_flattened(self):
        typed = (
            '<w:ins w:id="901" w:author="Hand" w:date="2026-01-01T00:00:00Z"><w:r>'
            '<w:t xml:space="preserve">inexpensive culture (Arden, 2019) and</w:t></w:r></w:ins>'
        )
        flat = self.head + deleted(self.code + self.tail) + typed + self.rest
        code, out, lines = self.run_tool(self.copy("flat.docx", lambda _x: flat))
        self.assertEqual((code, out["ok"], lines), (0, True, []))
        self.assertEqual(
            [(f["kind"], f["shown"], f["state"]) for f in out["flattened"]],
            [("ZOTERO_ITEM", "(Arden, 2019)", "deleted")],
        )

    def test_a_removed_field_needs_expect_removed_with_its_count(self):
        gone = self.copy("gone.docx", lambda _x: self.head + self.rest)
        self.assertEqual(self.run_tool(gone)[0], 1)
        code, out, lines = self.run_tool(gone, "--expect-removed", "1")
        self.assertEqual((code, out["failures"], lines), (0, [], []))
        self.assertEqual([m["kind"] for m in out["missing"]], ["ZOTERO_ITEM"])
        code, out, lines = self.run_tool(gone, "--expect-removed", "2")
        self.assertEqual((code, out["failures"]), (1, ["removed-count"]))
        self.assertTrue(any("1 instruction(s) gone, expected 2" in l for l in lines), lines)

    def test_expect_removed_does_not_allow_a_changed_instruction(self):
        code, out, _ = self.run_tool(
            self.copy("byte2.docx", lambda x: x.replace("cFx00001", "cFx0000X", 1)),
            "--expect-removed",
            "1",
        )
        self.assertEqual((code, out["failures"]), (1, ["instruction-added-or-changed"]))

    def test_a_lost_custom_property_fails(self):
        def drop(x):
            a = x.index('<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3"')
            return x[:a] + x[x.index("</property>", a) + len("</property>") :]

        code, out, lines = self.run_tool(self.copy("prop.docx", drop, "docProps/custom.xml"))
        self.assertEqual((code, out["failures"]), (1, ["property"]))
        self.assertEqual(out["properties"]["differ"], ["ZOTERO_PREF_2"])
        self.assertEqual(lines, ["FAIL props | ZOTERO_PREF_2 lost"])

    def test_a_changed_document_variable_fails(self):
        code, out, _ = self.run_tool(
            self.copy(
                "var.docx",
                lambda x: x.replace("&lt;Enabled&gt;0", "&lt;Enabled&gt;1", 1),
                "word/settings.xml",
            )
        )
        self.assertEqual((code, out["failures"]), (1, ["document-variable"]))

    def test_changed_field_data_fails(self):
        def edit(x):
            a = x.index("<w:fldData")
            b = x.index(">", a) + 1
            return x[:b] + "QUJD" + x[b:]

        code, out, _ = self.run_tool(self.copy("data.docx", edit))
        self.assertEqual((code, out["failures"]), (1, ["fldData"]))

    def test_changed_citation_text_is_listed_and_does_not_fail(self):
        code, out, _ = self.run_tool(
            self.copy("text.docx", lambda x: x.replace("(Arden, 2019)</w:t>", "(Arden 2019)</w:t>", 1))
        )
        self.assertEqual((code, out["ok"]), (0, True))
        self.assertEqual(len(out["text_changed"]), 1)
        self.assertEqual(out["text_changed"][0]["recorded"], "(Arden, 2019)")
        self.assertEqual(out["text_changed"][0]["accepted"], "(Arden 2019)")

    def test_simple_fields_and_other_stories_are_read(self):
        notes = (
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:footnote w:id="2"><w:p><w:fldSimple w:instr=" REF _RefX \\h "><w:r><w:t>x</w:t></w:r>'
            "</w:fldSimple></w:p></w:footnote></w:footnotes>"
        )
        out_path = self.dir / "notes.docx"
        with zipfile.ZipFile(FIXTURE) as zin, zipfile.ZipFile(out_path, "w") as z:
            for item in zin.infolist():
                z.writestr(item, zin.read(item.filename))
            z.writestr("word/footnotes.xml", notes)
        doc = tool.read(out_path)
        extra = [f for f in doc["fields"] if f["story"] == "word/footnotes.xml"]
        self.assertEqual([(f["form"], f["kind"], f["state"]) for f in extra], [("simple", "REF", "live")])
        self.assertEqual(self.run_tool(out_path)[1]["failures"], ["instruction-added-or-changed"])

    def test_the_copy_with_the_task_pane_embedded_keeps_everything(self):
        out_path = self.dir / "embedded.docx"
        subprocess.run(
            [sys.executable, str(EMBED), "--manifest", str(ROOT / "manifests" / "word.xml")]
            + ["--src", str(FIXTURE), "--out", str(out_path)],
            check=True,
            capture_output=True,
        )
        code, out, _ = self.run_tool(out_path)
        self.assertEqual((code, out["failures"]), (0, []))

    def test_a_file_that_is_no_docx_exits_2(self):
        bad = self.dir / "bad.docx"
        bad.write_text("not a zip")
        p = subprocess.run(
            [sys.executable, str(TOOL), "--before", str(FIXTURE), "--after", str(bad)],
            capture_output=True,
            text=True,
        )
        self.assertEqual(p.returncode, 2)
        self.assertEqual(json.loads(p.stdout.strip())["failures"], ["unreadable"])


if __name__ == "__main__":
    unittest.main()
