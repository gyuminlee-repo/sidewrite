"""Cases the marked-diff tool once got wrong, each with a small invented document.

Run: python3 -m unittest tests/test_marked_diff_cases.py
"""

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TOOL = ROOT / "tools" / "build_marked_diff.py"

P1 = "The first paragraph of this invented text stays as it is."
P2 = "The second paragraph mentions Figure 1 and nothing else."
P2_NEW = P2.replace("nothing", "little")


def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), v)
    return e


def doc_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


class Case(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.n = 0

    def tearDown(self):
        self.tmp.cleanup()

    def doc(self, paragraphs, edit=None):
        """Write a document with one run per paragraph; `edit(document)` may change it before it is saved."""
        self.n += 1
        path = self.dir / ("doc%d.docx" % self.n)
        d = Document()
        for text in paragraphs:
            d.add_paragraph().add_run(text)
        if edit:
            edit(d)
        d.save(path)
        return path

    def run_tool(self, base, clean, *args):
        self.n += 1
        out = self.dir / ("out%d.docx" % self.n)
        r = subprocess.run(
            [sys.executable, str(TOOL), "--baseline", str(base), "--clean", str(clean), "--out", str(out), "--json", *args],
            capture_output=True, text=True,
        )
        last = r.stdout.strip().splitlines()[-1] if r.stdout.strip() else ""
        return r.returncode, out, (json.loads(last) if last.startswith("{") else None), r.stdout + r.stderr

    def accept_reject(self, out):
        """(accept-all, reject-all) paragraph texts of a marked file, table cells included, in document order.
        Colour mode: struck = deleted, red = inserted. Track changes: inside w:del / w:ins."""
        accept, reject = [], []
        for p in Document(out).element.body.iter(qn("w:p")):
            a = r = ""
            for run in p.iter(qn("w:r")):
                text = "".join(t.text or "" for t in run if t.tag in (qn("w:t"), qn("w:delText")))
                rpr = run.find(qn("w:rPr"))
                strike = rpr.find(qn("w:strike")) if rpr is not None else None
                struck = strike is not None and strike.get(qn("w:val")) not in ("0", "false", "off")
                colour = rpr.find(qn("w:color")) if rpr is not None else None
                red = colour is not None and colour.get(qn("w:val")) == "FF0000"
                marks = {x.tag for x in run.iterancestors()}
                deleted = struck or qn("w:del") in marks
                inserted = (red and not struck) or qn("w:ins") in marks
                a += "" if deleted else text
                r += "" if inserted else text
            accept.append(a)
            reject.append(r)
        return [t for t in accept if t], [t for t in reject if t]


class StrikeSwitchedOff(Case):
    """<w:strike w:val="0"/> says "not struck"; the element being there does not make the run struck."""

    def off(self, tag, val):
        def edit(d):
            rpr = d.paragraphs[0].runs[0]._r.get_or_add_rPr()
            rpr.append(el("w:" + tag, val=val))

        return edit

    def test_a_clean_run_with_strike_switched_off_is_ordinary_text(self):
        base = self.doc([P1, P2])
        for tag in ("strike", "dstrike"):
            for val in ("0", "false", "off"):
                with self.subTest(tag=tag, val=val):
                    clean = self.doc([P1, P2_NEW], self.off(tag, val))
                    self.assertIn('<w:%s w:val="%s"/>' % (tag, val), doc_xml(clean))
                    code, out, summary, log = self.run_tool(base, clean)
                    self.assertEqual(code, 0, log)
                    self.assertEqual(self.accept_reject(out), ([P1, P2_NEW], [P1, P2]))

    def test_a_baseline_run_with_strike_switched_off_is_read_as_text(self):
        base = self.doc([P1, P2], self.off("strike", "0"))
        clean = self.doc([P1, P2_NEW])
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(code, 0, log)
        # Only the changed word is marked: the first paragraph was read, so it is not "inserted".
        self.assertEqual(summary["paragraphs"]["n_ins"], 0)
        self.assertEqual(self.accept_reject(out), ([P1, P2_NEW], [P1, P2]))

    def test_strike_switched_on_still_means_struck(self):
        base = self.doc([P1, P2])
        for val in (None, "1", "true", "on"):
            with self.subTest(val=val):
                def edit(d, val=val):
                    rpr = d.paragraphs[0].runs[0]._r.get_or_add_rPr()
                    rpr.append(el("w:strike", **({"val": val} if val else {})))

                clean = self.doc([P1, P2_NEW], edit)
                code, out, summary, log = self.run_tool(base, clean)
                self.assertEqual(code, 2, log)
                self.assertIn("struck-through", log)
                self.assertFalse(out.exists())


def table(rows):
    """edit(document): append a table with the given cell texts."""

    def edit(d):
        t = d.add_table(rows=len(rows), cols=len(rows[0]))
        for i, row in enumerate(rows):
            for j, text in enumerate(row):
                t.cell(i, j).text = text

    return edit


def both(*edits):
    def edit(d):
        for e in edits:
            e(d)

    return edit


class TableText(Case):
    """Text of a table that is new, laid out differently, or gone is marked and passes the tool's own check."""

    MODES = ("color", "track_changes")

    def check(self, base, clean, old, new):
        for mode in self.MODES:
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertTrue(out.exists())
                self.assertEqual(self.accept_reject(out), (new, old))

    def test_a_table_added_in_the_clean_document(self):
        base = self.doc([P1, P2])
        clean = self.doc([P1, P2], table([["A new cell.", "Another new cell."]]))
        self.check(base, clean, [P1, P2], [P1, P2, "A new cell.", "Another new cell."])

    def test_a_table_whose_layout_changed(self):
        base = self.doc([P1, P2], table([["An old cell."]]))
        clean = self.doc([P1, P2], table([["An old cell.", "A second column."]]))
        self.check(base, clean, [P1, P2, "An old cell."], [P1, P2, "An old cell.", "A second column."])

    def test_a_cell_with_a_line_break_and_a_tab(self):
        def edit(d):
            run = d.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0].add_run("Line one")
            run.add_break()
            run.add_tab()
            run.add_text("line two.")

        base = self.doc([P1, P2])
        clean = self.doc([P1, P2], edit)
        for mode in self.MODES:
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)

    def test_the_second_of_two_tables_removed(self):
        base = self.doc([P1, P2], both(table([["First table."]]), table([["Second table, removed."]])))
        clean = self.doc([P1, P2], table([["First table."]]))
        self.check(base, clean, [P1, P2, "First table.", "Second table, removed."], [P1, P2, "First table."])

    def test_a_table_with_fewer_rows(self):
        base = self.doc([P1, P2], table([["Row one."], ["Row two."]]))
        clean = self.doc([P1, P2], table([["Row one."]]))
        self.check(base, clean, [P1, P2, "Row one.", "Row two."], [P1, P2, "Row one."])

    def test_every_table_removed_is_refused_with_a_reason(self):
        # Known limit: a deleted table is placed next to a clean table; with none left there is no place for it.
        base = self.doc([P1, P2], table([["Only table."]]))
        clean = self.doc([P1, P2])
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(code, 2, log)
        self.assertIn("no anchor", log)
        self.assertFalse(out.exists())


def tool_module():
    spec = importlib.util.spec_from_file_location("build_marked_diff", TOOL)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def mark(tag, n=[900]):
    """An empty revision element with an id, an author and a date."""
    n[0] += 1
    return el(tag, id=str(n[0]), author="Somebody", date="2031-01-02T03:04:05Z")


def wrap(run, tag):
    """Put a run inside a w:ins / w:del / w:moveFrom / w:moveTo; deleted text is w:delText, as Word writes it."""
    w = mark(tag)
    run._r.addprevious(w)
    w.append(run._r)
    if tag == "w:del":
        for t in run._r.findall(qn("w:t")):
            t.tag = qn("w:delText")
    return w


def para_mark(paragraph, tag):
    """Mark the paragraph mark itself as inserted or deleted."""
    ppr = paragraph._p.get_or_add_pPr()
    rpr = el("w:rPr")
    ppr.append(rpr)
    rpr.append(mark(tag))


class PendingRevisions(Case):
    """Tracked changes that are still pending in an input are read as accepted, and do not end up among the tool's marks."""

    MODES = ("color", "track_changes")
    LEFTOVERS = ("Somebody", "w:moveFrom", "w:moveTo", "w:rPrChange", "w:pPrChange", "RangeStart", "RangeEnd")

    def check(self, base, clean, old, new, changed=None):
        for mode in self.MODES:
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                accept, reject = self.accept_reject(out)
                self.assertEqual(accept, new)
                # A deleted sentence may be shown struck in the neighbouring paragraph (the tool keeps the old
                # paragraph boundary as a struck line break), so the old text is compared as a whole.
                self.assertEqual("".join(reject), "".join(old))
                xml = doc_xml(out)
                for word in self.LEFTOVERS:
                    self.assertNotIn(word, xml)
                if changed is not None:
                    self.assertEqual(summary["insert_runs"] + summary["delete_runs"] > 0, changed)

    def test_a_pending_insertion_whose_text_equals_the_baseline(self):
        base = self.doc([P1, P2])
        clean = self.doc([P1, P2], lambda d: wrap(d.paragraphs[1].runs[0], "w:ins"))
        self.check(base, clean, [P1, P2], [P1, P2], changed=False)

    def test_a_pending_deletion_is_text_that_is_gone(self):
        def edit(d):
            wrap(d.paragraphs[1].add_run(" This sentence was deleted with tracking on."), "w:del")

        base = self.doc([P1, P2 + " This sentence was deleted with tracking on."])
        clean = self.doc([P1, P2], edit)
        self.check(base, clean, [P1, P2 + " This sentence was deleted with tracking on."], [P1, P2], changed=True)
        # ... and when the baseline never had it, nothing is marked.
        self.check(self.doc([P1, P2]), clean, [P1, P2], [P1, P2], changed=False)

    def test_a_pending_move_is_read_at_its_new_place(self):
        moved = " A sentence that was moved."

        def edit(d):
            p1, p2 = d.paragraphs
            source = wrap(p1.add_run(moved), "w:moveFrom")
            source.addprevious(el("w:moveFromRangeStart", id="50", name="move1"))
            source.addnext(el("w:moveFromRangeEnd", id="50"))
            target = wrap(p2.add_run(moved), "w:moveTo")
            target.addprevious(el("w:moveToRangeStart", id="51", name="move1"))
            target.addnext(el("w:moveToRangeEnd", id="51"))

        base = self.doc([P1 + moved, P2])
        clean = self.doc([P1, P2], edit)
        self.check(base, clean, [P1 + moved, P2], [P1, P2 + moved], changed=True)
        self.check(self.doc([P1, P2 + moved]), clean, [P1, P2 + moved], [P1, P2 + moved], changed=False)

    def test_pending_format_changes_are_dropped(self):
        def edit(d):
            p = d.paragraphs[1]
            p.runs[0].bold = True
            change = mark("w:rPrChange")
            change.append(el("w:rPr"))
            p.runs[0]._r.rPr.append(change)
            ppr = p._p.get_or_add_pPr()
            change = mark("w:pPrChange")
            change.append(el("w:pPr"))
            ppr.append(change)

        base = self.doc([P1, P2])
        self.check(base, self.doc([P1, P2_NEW], edit), [P1, P2], [P1, P2_NEW], changed=True)
        # The baseline's formatting is copied into deleted runs: its pending changes must not travel along.
        self.check(self.doc([P1, P2], edit), self.doc([P1]), [P1, P2], [P1], changed=True)

    def test_a_deleted_paragraph_mark_joins_two_paragraphs(self):
        base = self.doc([P1 + P2, "A last paragraph."])
        clean = self.doc([P1, P2, "A last paragraph."], lambda d: para_mark(d.paragraphs[0], "w:del"))
        self.check(base, clean, [P1 + P2, "A last paragraph."], [P1 + P2, "A last paragraph."], changed=False)

    def test_a_wholly_deleted_paragraph_is_gone(self):
        def edit(d):
            wrap(d.paragraphs[1].runs[0], "w:del")
            para_mark(d.paragraphs[1], "w:del")
            para_mark(d.paragraphs[0], "w:ins")

        base = self.doc([P1, "A last paragraph."])
        clean = self.doc([P1, P2, "A last paragraph."], edit)
        self.check(base, clean, [P1, "A last paragraph."], [P1, "A last paragraph."], changed=False)

    def test_a_deleted_table_row_is_gone(self):
        def edit(d):
            table([["Row one."], ["Row two, deleted."]])(d)
            row = d.tables[0].rows[1]._tr
            trpr = el("w:trPr")
            trpr.append(mark("w:del"))
            row.insert(0, trpr)
            wrap(d.tables[0].cell(1, 0).paragraphs[0].runs[0], "w:del")

        base = self.doc([P1, P2], table([["Row one."]]))
        clean = self.doc([P1, P2], edit)
        self.check(base, clean, [P1, P2, "Row one."], [P1, P2, "Row one."], changed=False)

    def test_a_pending_cell_change_is_refused_with_a_reason(self):
        def edit(d):
            table([["One.", "Two."]])(d)
            d.tables[0].cell(0, 1)._tc.get_or_add_tcPr().append(mark("w:cellDel"))

        code, out, summary, log = self.run_tool(self.doc([P1, P2]), self.doc([P1, P2], edit))
        self.assertEqual(code, 2, log)
        self.assertIn("cellDel", log)
        self.assertFalse(out.exists())

    def test_a_document_without_pending_changes_is_not_rewritten(self):
        xml = doc_xml(self.doc([P1, P2])).encode("utf-8")
        got, n = tool_module().accept_revisions(xml)
        self.assertIs(got, xml)
        self.assertEqual(n, 0)

    def test_the_summary_says_how_many_pending_changes_were_accepted(self):
        base = self.doc([P1, P2])
        clean = self.doc([P1, P2], lambda d: wrap(d.paragraphs[1].runs[0], "w:ins"))
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(summary["accepted_revisions"], {"baseline": 0, "clean": 1})
        self.assertIn("pending tracked changes read as accepted: baseline 0, clean 1", log)


if __name__ == "__main__":
    unittest.main()
