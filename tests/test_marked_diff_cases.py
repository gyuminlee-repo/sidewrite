"""Cases the marked-diff tool once got wrong, each with a small invented document.

Run: python3 -m unittest tests/test_marked_diff_cases.py
"""

import hashlib
import importlib.util
import json
import re
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

    def test_the_only_table_removed(self):
        # The deleted table is put behind the paragraph it followed in the baseline; it needs no other table.
        base = self.doc([P1, P2, "A last paragraph."], lambda d: d.paragraphs[1]._p.addnext(d.add_table(1, 1)._tbl))
        for cell in Document(base).tables[0].rows[0].cells:
            self.assertEqual(cell.text, "")
        clean = self.doc([P1, P2, "A last paragraph."])
        for mode in self.MODES:
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                body = Document(out).element.body
                self.assertEqual([x.tag.split("}")[1] for x in body][:4], ["p", "p", "tbl", "p"])

    def test_text_behind_a_last_table_where_the_clean_document_has_no_paragraph_is_refused(self):
        # python-docx leaves no paragraph behind a last table. The clean document then has a paragraph there
        # (the last one) where the baseline has none after its table: no paragraph mark can show that.
        base = self.doc([P1, P2], table([["Only table."]]))
        clean = self.doc([P1, P2, "A last paragraph."], table([["Only table."]]))
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(code, 0, log)
        clean = self.doc([P1], table([["Only table."]]))
        base = self.doc([P1], both(table([["Only table."]]), lambda d: d.add_paragraph().add_run("A last paragraph.")))
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(code, 2, log)
        self.assertIn("a paragraph mark cannot show that", log)
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
                # A deleted sentence may be shown in a paragraph of its own or in the neighbouring one, and this
                # helper does not join paragraphs at their marks, so the old text is compared as a whole.
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


def START(name, n):
    return el("w:bookmarkStart", id=str(n), name=name)


def END(n):
    return el("w:bookmarkEnd", id=str(n))


def pieces(index, *parts):
    """edit(document): rewrite paragraph `index` as the given parts: a string is a run, anything else an element."""

    def edit(d):
        p = d.paragraphs[index]
        for r in p.runs:
            p._p.remove(r._r)
        for part in parts:
            if isinstance(part, str):
                p.add_run(part)
            else:
                p._p.append(part)

    return edit


def is_deleted(run):
    rpr = run.find(qn("w:rPr"))
    strike = rpr.find(qn("w:strike")) if rpr is not None else None
    return strike is not None or any(a.tag == qn("w:del") for a in run.iterancestors())


def bookmark_text(path, name, accepted=False):
    """Text between the start and the end of a bookmark; with `accepted`, without the text shown as deleted."""
    body = Document(path).element.body
    text, wanted, inside = "", None, False
    for e in body.iter(qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:r")):
        if e.tag == qn("w:bookmarkStart") and e.get(qn("w:name")) == name:
            wanted, inside = e.get(qn("w:id")), True
        elif e.tag == qn("w:bookmarkEnd") and inside and e.get(qn("w:id")) == wanted:
            return text
        elif e.tag == qn("w:r") and inside and not (accepted and is_deleted(e)):
            text += "".join(x.text or "" for x in e if x.tag in (qn("w:t"), qn("w:delText")))
    raise AssertionError("bookmark %s has no start and end in this order" % name)


class Bookmarks(Case):
    """A bookmark in a changed paragraph keeps the text it had: its ends stay at their places in the new text."""

    MODES = ("color", "track_changes")
    HEAD, MID, TAIL = "The second paragraph mentions ", "Figure 1", " and little else."

    def marked(self, base_text, *parts):
        """Run both modes on [P1, base_text] against [P1, parts]: [(mode, output, summary, log)]."""
        base = self.doc([P1, base_text])
        clean = self.doc([P1, "x"], pieces(1, *parts))
        runs = []
        for mode in self.MODES:
            code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
            self.assertEqual(code, 0, log)
            runs.append((mode, out, summary, log))
        return runs

    def test_a_bookmark_in_the_middle_of_a_changed_paragraph(self):
        for mode, out, summary, log in self.marked(P2, self.HEAD, START("fig1", 7), self.MID, END(7), self.TAIL):
            with self.subTest(mode=mode):
                self.assertEqual(bookmark_text(out, "fig1"), "Figure 1")
                self.assertEqual(summary["warnings"], [])

    def test_text_deleted_right_before_and_right_after_stays_outside(self):
        base = "The second paragraph mentions the old Figure 1 indeed and little else."
        for mode, out, summary, log in self.marked(base, self.HEAD, START("fig1", 7), self.MID, END(7), self.TAIL):
            with self.subTest(mode=mode):
                self.assertEqual(bookmark_text(out, "fig1"), "Figure 1")
                self.assertEqual(summary["warnings"], [])

    def test_a_bookmark_over_the_whole_paragraph_keeps_deleted_edges_outside(self):
        base = "Formerly first. The second paragraph mentions Figure 1 and little else. Formerly last."
        text = self.HEAD + self.MID + self.TAIL
        for mode, out, summary, log in self.marked(base, START("all", 7), text, END(7)):
            with self.subTest(mode=mode):
                self.assertEqual(bookmark_text(out, "all"), text)
                self.assertEqual(summary["warnings"], [])

    def test_text_deleted_inside_a_bookmark_is_reported(self):
        base = "The second paragraph mentions Figure 1 of the old set and little else."
        parts = (self.HEAD, START("fig1", 7), "Figure 1 of the set", END(7), self.TAIL)
        for mode, out, summary, log in self.marked(base, *parts):
            with self.subTest(mode=mode):
                self.assertEqual(bookmark_text(out, "fig1", accepted=True), "Figure 1 of the set")
                if mode == "color":
                    # The struck words sit between the two ends: nothing can keep them out, so the run says so.
                    self.assertEqual(bookmark_text(out, "fig1"), "Figure 1 of the old set")
                    self.assertEqual(len(summary["warnings"]), 1)
                    self.assertIn('"fig1"', summary["warnings"][0])
                    self.assertIn("WARNING", log)
                else:
                    self.assertEqual(summary["warnings"], [])

    def test_an_empty_bookmark_stays_empty_and_in_order(self):
        for mode, out, summary, log in self.marked(P2, self.HEAD, START("point", 7), END(7), self.MID + self.TAIL):
            with self.subTest(mode=mode):
                self.assertEqual(bookmark_text(out, "point"), "")
                xml = doc_xml(out)
                self.assertLess(xml.index("<w:bookmarkStart"), xml.index("<w:bookmarkEnd"))
                self.assertIn(self.HEAD, xml[: xml.index('w:name="point"')])

    def test_two_bookmarks_that_meet_where_text_was_deleted(self):
        base = "The second paragraph mentions Figure 1 or so and little else."
        parts = (START("a", 7), self.HEAD + self.MID, END(7), START("b", 8), self.TAIL, END(8))
        for mode, out, summary, log in self.marked(base, *parts):
            with self.subTest(mode=mode):
                self.assertEqual(bookmark_text(out, "a"), self.HEAD + self.MID)
                self.assertEqual(bookmark_text(out, "b"), self.TAIL)
                self.assertEqual(summary["warnings"], [])

    def test_a_bookmark_inside_a_link(self):
        link = el("w:hyperlink", anchor="somewhere")
        for part in (START("fig1", 7), None, END(7)):
            if part is None:
                run = el("w:r")
                text = el("w:t")
                text.text = self.MID
                run.append(text)
                part = run
            link.append(part)
        for mode, out, summary, log in self.marked(P2, self.HEAD, link, self.TAIL):
            with self.subTest(mode=mode):
                self.assertEqual(bookmark_text(out, "fig1"), "Figure 1")
                xml = doc_xml(out)
                inside = xml[xml.index("<w:hyperlink"): xml.index("</w:hyperlink>")]
                self.assertIn("Figure 1", inside)
                self.assertIn("w:bookmarkStart", inside)

    def test_a_bookmark_across_paragraphs_with_a_deletion_between_its_ends(self):
        def ends(d):
            d.paragraphs[0]._p.insert(0, START("span", 9))
            d.add_paragraph().add_run("A last paragraph.")
            d.paragraphs[2]._p.append(END(9))

        base = self.doc([P1, P2, "A last paragraph."])
        clean = self.doc([P1, P2_NEW], ends)
        for mode in self.MODES:
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertEqual(len(summary["warnings"]), 1 if mode == "color" else 0)

    def test_a_document_without_bookmarks_has_no_warnings(self):
        code, out, summary, log = self.run_tool(self.doc([P1, P2]), self.doc([P1, P2_NEW]))
        self.assertEqual(summary["warnings"], [])
        self.assertNotIn("WARNING", log)


def with_break(attrs):
    """edit(document): put a break with these attributes between two runs of paragraph 1."""

    def edit(d):
        p = d.paragraphs[1]
        run = el("w:r")
        run.append(el("w:br", **attrs))
        p._p.append(run)
        p.add_run("More text after the break.")

    return edit


class BreakKinds(Case):
    """A break in a rewritten paragraph stays the break it was: line, page, column, with its clear attribute."""

    KINDS = (
        ({}, "<w:br/>"),
        ({"type": "page"}, '<w:br w:type="page"/>'),
        ({"type": "column"}, '<w:br w:type="column"/>'),
        ({"clear": "all"}, '<w:br w:clear="all"/>'),
        ({"clear": "left"}, '<w:br w:clear="left"/>'),
        ({"type": "textWrapping", "clear": "right"}, '<w:br w:clear="right"/>'),
    )

    def test_a_break_keeps_its_kind_when_a_word_of_the_paragraph_changes(self):
        for attrs, written in self.KINDS:
            base = self.doc([P1, P2], with_break(attrs))
            clean = self.doc([P1, P2_NEW], with_break(attrs))
            for mode in ("color", "track_changes"):
                with self.subTest(attrs=attrs, mode=mode):
                    code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                    self.assertEqual(code, 0, log)
                    xml = doc_xml(out)
                    self.assertEqual(xml.count("<w:br"), 1, xml[xml.index("<w:body"):][:900])
                    self.assertIn(written, xml)

    def test_a_break_of_another_kind_is_a_change(self):
        base = self.doc([P1, P2], with_break({"type": "page"}))
        clean = self.doc([P1, P2], with_break({"type": "column"}))
        code, out, summary, log = self.run_tool(base, clean, "--mode", "track_changes")
        self.assertEqual(code, 0, log)
        self.assertIn('<w:br w:type="column"/>', doc_xml(out))
        self.assertGreater(summary["insert_runs"], 0)


sys.path.insert(0, str(Path(__file__).resolve().parent))
import marked_diff_corpus as corpus  # noqa: E402


class ParagraphMarks(Case):
    """A paragraph that is new or gone is marked in its paragraph mark, so the paragraph itself comes and goes."""

    MODES = ("color", "track_changes")
    NEW = "A paragraph that is new in the clean document."

    def views(self, out, mode):
        """(accept-all, reject-all) as lists of paragraph texts, empty paragraphs included, by the corpus reader."""
        return tuple([x[1] for x in corpus.structure_read(corpus.view(out, mode, side))[0] if x[0] == "p"]
                     for side in ("accept", "reject"))

    def test_a_new_paragraph_between_two_has_an_inserted_paragraph_mark(self):
        base = self.doc([P1, P2])
        clean = self.doc([P1, self.NEW, P2])
        for mode in self.MODES:
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertEqual(self.views(out, mode), ([P1, self.NEW, P2], [P1, P2]))
                self.assertEqual(summary["paragraph_marks"], {"ins": 1})
                mark = Document(out).paragraphs[1]._p.find(qn("w:pPr") + "/" + qn("w:rPr"))
                if mode == "track_changes":
                    self.assertIsNotNone(mark.find(qn("w:ins")))
                else:
                    self.assertEqual(mark.find(qn("w:color")).get(qn("w:val")), "FF0000")

    def test_a_deleted_paragraph_has_a_deleted_paragraph_mark(self):
        base = self.doc([P1, self.NEW, P2])
        clean = self.doc([P1, P2])
        for mode in self.MODES:
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertEqual(self.views(out, mode), ([P1, P2], [P1, self.NEW, P2]))
                self.assertEqual(summary["paragraph_marks"], {"del": 1})

    def test_paragraphs_at_the_end_come_and_go_too(self):
        for old, new in (([P1, P2], [P1, P2, self.NEW]), ([P1, P2, self.NEW], [P1, P2])):
            base, clean = self.doc(old), self.doc(new)
            for mode in self.MODES:
                with self.subTest(mode=mode, added=len(new) > len(old)):
                    code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                    self.assertEqual(code, 0, log)
                    self.assertEqual(self.views(out, mode), (new, old))

    def test_empty_paragraphs_come_and_go(self):
        for old, new in (([P1, P2], [P1, "", P2]), ([P1, "", "", P2], [P1, P2]), ([P1, "", P2], [P1, "", P2])):
            base, clean = self.doc(old), self.doc(new)
            for mode in self.MODES:
                with self.subTest(mode=mode, old=len(old), new=len(new)):
                    code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                    self.assertEqual(code, 0, log)
                    self.assertEqual(self.views(out, mode), (new, old))

    def test_a_split_and_a_join_are_a_paragraph_mark(self):
        for old, new in (([P1 + " " + P2, "A last paragraph."], [P1, P2, "A last paragraph."]),
                         ([P1, P2, "A last paragraph."], [P1 + " " + P2, "A last paragraph."])):
            base, clean = self.doc(old), self.doc(new)
            for mode in self.MODES:
                with self.subTest(mode=mode, split=len(new) > len(old)):
                    code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                    self.assertEqual(code, 0, log)
                    self.assertEqual(self.views(out, mode), (new, old))
                    self.assertNotIn('w:clear="none"', doc_xml(out))

    def test_a_line_break_written_with_clear_none_is_a_line_break(self):
        base = self.doc([P1, P2], with_break({"clear": "none"}))
        for clean in (self.doc([P1, P2], with_break({"clear": "none"})), self.doc([P1, P2_NEW], with_break({"clear": "none"}))):
            for mode in self.MODES:
                with self.subTest(mode=mode):
                    code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                    self.assertEqual(code, 0, log)
                    self.assertEqual(doc_xml(out).count("<w:br"), 1)

    def test_a_clean_file_with_a_paragraph_mark_in_the_insert_colour_is_refused_in_colour_mode(self):
        def red_mark(d):
            rpr = el("w:rPr")
            rpr.append(el("w:color", val="FF0000"))
            d.paragraphs[0]._p.get_or_add_pPr().append(rpr)

        base, clean = self.doc([P1, P2]), self.doc([P1, P2_NEW], red_mark)
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(code, 2, log)
        self.assertIn("paragraph mark", log)
        code, out, summary, log = self.run_tool(base, clean, "--insert-rgb", "0000FF")
        self.assertEqual(code, 0, log)
        code, out, summary, log = self.run_tool(base, clean, "--mode", "track_changes")
        self.assertEqual(code, 0, log)


class TableRows(Case):
    """A table whose rows changed is a deleted table and an inserted table, with the rows marked."""

    def test_a_table_that_lost_a_row_round_trips_in_track_changes_mode(self):
        base = self.doc([P1, P2], both(table([["Row one."], ["Row two."]]), lambda d: d.add_paragraph().add_run("End.")))
        clean = self.doc([P1, P2], both(table([["Row one."]]), lambda d: d.add_paragraph().add_run("End.")))
        code, out, summary, log = self.run_tool(base, clean, "--mode", "track_changes")
        self.assertEqual(code, 0, log)
        tool = tool_module()
        xml = doc_xml(out).encode("utf-8")
        for side, rows in (("accept", 1), ("reject", 2)):
            resolved = tool.resolve_revisions(xml, side)[0].decode("utf-8")
            self.assertEqual((resolved.count("<w:tbl>"), resolved.count("<w:tr>") + resolved.count("<w:tr ")), (1, rows), side)
        # ... and by the corpus module's own reader.
        tables = lambda side: [x[1] for x in corpus.structure_read(corpus.view(out, "track_changes", side))[0] if x[0] == "t"]
        self.assertEqual(tables("accept"), [((("Row one.",),),)])
        self.assertEqual(tables("reject"), [((("Row one.",),), (("Row two.",),))])

    def test_the_same_in_colour_mode(self):
        base = self.doc([P1, P2], both(table([["Row one."], ["Row two."]]), lambda d: d.add_paragraph().add_run("End.")))
        clean = self.doc([P1, P2], both(table([["Row one."]]), lambda d: d.add_paragraph().add_run("End.")))
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(code, 0, log)
        tables = lambda side: [x[1] for x in corpus.structure_read(corpus.view(out, "color", side))[0] if x[0] == "t"]
        self.assertEqual(tables("accept"), [((("Row one.",),),)])
        self.assertEqual(tables("reject"), [((("Row one.",),), (("Row two.",),))])


GRID = [["One.", "Two."], ["Three.", "Four."]]


def merged(first, last, drop):
    """edit(document): the 2x2 table with two cells merged; the text of the cell that is swallowed is dropped."""

    def edit(d):
        table(GRID)(d)
        t = d.tables[-1]
        t.cell(*drop).paragraphs[0].runs[0].text = ""
        t.cell(*first).merge(t.cell(*last))
        d.add_paragraph().add_run("End.")

    return edit


def picture(text=""):
    """edit(document): one more paragraph, with a picture of one dot after its text."""

    def edit(d):
        import io

        p = d.add_paragraph()
        if text:
            p.add_run(text)
        p.add_run().add_picture(io.BytesIO(corpus._png()))

    return edit


class MergedCells(Case):
    """Cells merged in the clean document only: rejecting gives the four cells of the baseline back."""

    PLAIN = staticmethod(both(table(GRID), lambda d: d.add_paragraph().add_run("End.")))

    def layouts(self, out, mode):
        return tuple([x[1] for x in corpus.structure_read(corpus.view(out, mode, side))[0] if x[0] == "g"]
                     for side in ("accept", "reject"))

    def check(self, first, last, drop, merged_layout):
        four = (2, (((1, None), (1, None)), ((1, None), (1, None))))
        for mode in ("color", "track_changes"):
            with self.subTest(mode=mode):
                base, clean = self.doc([P1, P2], self.PLAIN), self.doc([P1, P2], merged(first, last, drop))
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertEqual(self.layouts(out, mode), ([merged_layout], [four]))
                # ... and the other way round: the merge is in the baseline.
                code, out, summary, log = self.run_tool(clean, base, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertEqual(self.layouts(out, mode), ([four], [merged_layout]))

    def test_a_merge_down_a_column(self):
        self.check((0, 0), (1, 0), (1, 0), (2, (((1, "restart"), (1, None)), ((1, "continue"), (1, None)))))

    def test_a_merge_across_a_row(self):
        self.check((0, 0), (0, 1), (0, 1), (2, (((2, None),), ((1, None), (1, None)))))

    def test_the_tool_s_own_check_sees_a_merge_that_is_left_behind(self):
        """What the check is for: an output that keeps the clean table unmarked fails on the reject side."""
        base, clean = self.doc([P1, P2], self.PLAIN), self.doc([P1, P2], merged((0, 0), (1, 0), (1, 0)))
        tool = tool_module()
        accept, reject = tool.structure_failures(doc_xml(clean).encode(), doc_xml(base).encode(), doc_xml(clean).encode())
        self.assertEqual(accept, [])
        self.assertRegex(" ".join(reject), r"reject-all, tables: #1 .*restart")


class ElementsThatAreNotText(Case):
    """A picture, a field, a note reference or a sign that is no character is content. What is the same in both
    documents stays as the clean document has it; what is new or gone is marked with its own elements."""

    def pair(self, name):
        """The two files of a named pair of the corpus."""
        self.n += 1
        base, clean, _ = corpus.models(corpus.CASES[name])
        return (corpus.write(base, self.dir / ("b%d.docx" % self.n)), corpus.write(clean, self.dir / ("c%d.docx" % self.n)))

    def contents(self, out, mode):
        """(accept-all, reject-all): what every paragraph holds besides text, by the corpus reader."""
        return tuple([x[1] for x in corpus.structure_read(corpus.view(out, mode, side))[0] if x[0] == "x"]
                     for side in ("accept", "reject"))

    def test_a_new_paragraph_with_a_picture_is_inserted_with_its_picture(self):
        base, clean = self.doc([P1, P2]), self.doc([P1, P2], picture("A caption."))
        tool = tool_module()
        shown = ("drawing", hashlib.sha256(corpus._png()).hexdigest())
        for mode in ("color", "track_changes"):
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertEqual(self.contents(out, mode), ([(shown,)], []))
        xml = doc_xml(out).encode("utf-8")  # track changes: the run of the picture stands inside w:ins
        self.assertEqual(xml.count(b"<w:drawing>"), 1)
        self.assertNotIn(b"<w:drawing>", tool.resolve_revisions(xml, "reject")[0])
        self.assertIn(b"<w:drawing>", tool.resolve_revisions(xml, "accept")[0])

    def test_a_new_paragraph_that_is_only_a_picture(self):
        base, clean = self.doc([P1, P2]), self.doc([P1, P2], picture())
        code, out, summary, log = self.run_tool(base, clean, "--mode", "track_changes")
        self.assertEqual(code, 0, log)
        tool = tool_module()
        self.assertNotIn(b"<w:drawing>", tool.resolve_revisions(doc_xml(out).encode("utf-8"), "reject")[0])

    def test_the_tool_s_own_check_sees_a_picture_that_is_left_behind(self):
        base, clean = self.doc([P1, P2]), self.doc([P1, P2], picture("A caption."))
        tool = tool_module()
        accept, reject = tool.structure_failures(doc_xml(clean).encode(), doc_xml(base).encode(), doc_xml(clean).encode())
        self.assertEqual(accept, [])
        self.assertTrue(reject)

    def test_a_deleted_picture_is_shown_with_the_baseline_s_picture(self):
        """The picture of the baseline is not in the clean file: the part comes along, under a name of its own."""
        base, clean = self.pair("a picture swapped for another")
        old, new = hashlib.sha256(corpus._png()).hexdigest(), hashlib.sha256(corpus._png(200)).hexdigest()
        for mode in ("color", "track_changes"):
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertEqual(self.contents(out, mode), ([(("drawing", new),)], [(("drawing", old),)]))
                with zipfile.ZipFile(out) as z, zipfile.ZipFile(clean) as c:
                    self.assertEqual(z.namelist(), c.namelist() + ["word/media/bmd_1.png"])
                    self.assertEqual(z.read("word/media/bmd_1.png"), corpus._png())
                    self.assertIn(b'Target="media/bmd_1.png"', z.read("word/_rels/document.xml.rels"))
                    self.assertIn(b'PartName="/word/media/bmd_1.png" ContentType="image/png"', z.read("[Content_Types].xml"))
        code, out2, summary, log = self.run_tool(base, clean, "--mode", mode)
        self.assertEqual(Path(out).read_bytes(), Path(out2).read_bytes(), "the same inputs give the same file")

    def test_a_pair_without_deleted_parts_leaves_the_package_alone(self):
        base, clean = self.pair("a picture added to a paragraph")
        code, out, summary, log = self.run_tool(base, clean, "--mode", "track_changes")
        self.assertEqual(code, 0, log)
        with zipfile.ZipFile(out) as z, zipfile.ZipFile(clean) as c:
            self.assertEqual(z.namelist(), c.namelist())
            self.assertEqual(z.read("word/_rels/document.xml.rels"), c.read("word/_rels/document.xml.rels"))

    def test_a_citation_between_two_changed_words_stays_unmarked(self):
        base, clean = self.pair("words changed on both sides of a citation")
        code, out, summary, log = self.run_tool(base, clean, "--mode", "track_changes")
        self.assertEqual(code, 0, log)
        xml = doc_xml(out)
        marked = re.findall(r"<w:(?:ins|del) [^>]*>.*?</w:(?:ins|del)>", xml)
        self.assertTrue(marked)
        self.assertFalse([m for m in marked if "fldChar" in m or "nstrText" in m])
        self.assertEqual(xml.count("<w:fldChar "), 6)
        self.assertEqual(doc_xml(out).count("ZOTERO_ITEM"), doc_xml(clean).count("ZOTERO_ITEM"))

    def test_a_citation_whose_shown_text_changed_is_the_old_field_deleted_and_the_new_one_inserted(self):
        base, clean = self.pair("the shown text of a citation changed")
        tool = tool_module()
        code, out, summary, log = self.run_tool(base, clean, "--mode", "track_changes")
        self.assertEqual(code, 0, log)
        xml = doc_xml(out)
        gone = re.search(r"<w:del [^>]*>(?:(?!</w:del>).)*delInstrText.*?</w:del>", xml).group()
        self.assertEqual(gone.count("<w:fldChar "), 3)
        self.assertIn(">[2]</w:delText>", gone)
        self.assertNotIn("<w:instrText", gone)
        new = re.search(r"<w:ins [^>]*>(?:(?!</w:ins>).)*<w:instrText.*?</w:ins>", xml).group()
        self.assertIn(">[2, 3]</w:t>", new)
        # rejecting gives the code back as code
        back = tool.resolve_revisions(xml.encode("utf-8"), "reject")[0].decode("utf-8")
        self.assertNotIn("delInstrText", back)
        self.assertEqual(back.count("ZOTERO_ITEM"), doc_xml(base).count("ZOTERO_ITEM"))
        model = corpus.models(corpus.CASES["the shown text of a citation changed"])[0]
        self.assertEqual(self.contents(out, "track_changes")[1], [x[1] for x in corpus.structure_of(model)[0] if x[0] == "x"])

    def test_the_tool_s_own_check_sees_a_field_whose_result_differs(self):
        base, clean = self.pair("the result of a cross-reference changed")
        tool = tool_module()
        accept, reject = tool.structure_failures(doc_xml(clean).encode(), doc_xml(base).encode(), doc_xml(clean).encode())
        self.assertEqual(accept, [])
        self.assertRegex(" ".join(reject), r"reject-all, body paragraphs")

    def test_a_field_code_cut_differently_is_the_same_field(self):
        tool = tool_module()
        w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        mark = '<w:r><w:fldChar w:fldCharType="%s"/></w:r>'
        code = lambda *parts: "".join('<w:r><w:instrText xml:space="preserve">%s</w:instrText></w:r>' % x for x in parts)
        para = lambda inner: tool.etree.fromstring(
            "<w:p %s>%s%s%s<w:r><w:t>7</w:t></w:r>%s</w:p>" % (w, mark % "begin", inner, mark % "separate", mark % "end"))
        one, cut, other = (tool.read_para(para(code(*x))) for x in ((" PAGE \\* MERGEFORMAT ",), (" PA", "GE \\* MERGEFORMAT "), (" NUMPAGES ",)))
        self.assertEqual(one.text, cut.text)
        self.assertNotEqual(one.text, other.text)
        self.assertEqual((len(one.text), one.inline), (1, {"field": 1}))
        simple = tool.read_para(tool.etree.fromstring(
            '<w:p %s><w:fldSimple w:instr=" PAGE \\* MERGEFORMAT "><w:r><w:t>7</w:t></w:r></w:fldSimple></w:p>' % w))
        self.assertEqual(simple.text, one.text)

    def test_a_field_over_several_paragraphs_is_read_by_its_ends(self):
        base, clean = self.pair("the first entry of a bibliography reworded")
        for mode in ("color", "track_changes"):
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                xml = doc_xml(out)
                self.assertEqual(xml.count("ZOTERO_BIBL"), 1)       # the begin of the field is kept, not deleted and inserted
                self.assertEqual(xml.count("<w:fldChar "), 3)

    def test_a_paragraph_in_a_content_control_is_shown_whole_with_a_note(self):
        base, clean = self.pair("a word changed in a paragraph with a content control")
        for mode in ("color", "track_changes"):
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                self.assertRegex(" ".join(summary["warnings"]), r"clean paragraph 5 of the body .* holds sdt and is not rewritten")
                self.assertEqual(doc_xml(out).count("<w:sdt>"), 1)

    def test_a_deleted_note_reference_brings_its_note(self):
        tool = tool_module()
        for name, part in (("a footnote in a deleted paragraph", "word/bmd_footnotes.xml"),
                           ("a footnote swapped for another", "word/footnotes.xml")):
            base, clean = self.pair(name)
            for mode in ("color", "track_changes"):
                with self.subTest(case=name, mode=mode):
                    code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                    self.assertEqual(code, 0, log)
                    with zipfile.ZipFile(out) as z:
                        notes = tool.etree.fromstring(z.read(part))
                        refs = set(re.findall(r'<w:footnoteReference w:id="(\d+)"', z.read("word/document.xml").decode()))
                    ids = [n.get(tool.q("id")) for n in notes]
                    self.assertEqual(len(ids), len(set(ids)))
                    self.assertLessEqual(refs, set(ids))
                    self.assertIn(corpus.NOTE_TEXT["fn"], "".join(notes.itertext()))

    def test_an_equation_that_changes_is_refused_with_its_place(self):
        from lxml import etree

        m = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
        math = lambda text: (lambda d: d.paragraphs[-1]._p.append(etree.fromstring("<m:oMath %s><m:r><m:t>%s</m:t></m:r></m:oMath>" % (m, text))))
        plain = self.doc([P1, P2])
        with_x = self.doc([P1, P2, "A new paragraph. "], math("x"))
        with_y = self.doc([P1, P2, "A new paragraph. "], math("y"))
        reworded = self.doc([P1, P2, "A changed paragraph. "], math("x"))
        for name, base, clean, message in (
            ("a new paragraph", plain, with_x, r"clean paragraph 3 of the body \('A new paragraph. <oMath>'\) holds an element \(oMath\) that stands outside a run and cannot be shown as inserted"),
            ("a deleted paragraph", with_x, plain, r"baseline paragraph 3 of the body .* \(oMath\) that stands outside a run and cannot be shown as deleted"),
            ("another equation", with_x, with_y, r"clean paragraph 3 of the body .* \(oMath\) that stands outside a run"),
        ):
            for mode in ("color", "track_changes"):
                with self.subTest(case=name, mode=mode):
                    code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                    self.assertEqual(code, 2, log)
                    self.assertRegex(log, message)
                    self.assertFalse(out.exists())
        code, out, summary, log = self.run_tool(with_x, reworded, "--mode", "track_changes")   # the equation stays: written
        self.assertEqual(code, 0, log)
        self.assertEqual(doc_xml(out).count("<m:oMath"), 1)

    def test_a_file_with_characters_of_the_private_plane_is_refused(self):
        base, clean = self.doc([P1, P2]), self.doc([P1, P2 + " \U000F0001"])
        code, out, summary, log = self.run_tool(base, clean)
        self.assertEqual(code, 2, log)
        self.assertRegex(log, r"private plane")

    def test_a_comment_is_no_change(self):
        """The two invented manuscripts differ in a comment on a paragraph with the same text: both ways are written."""
        fx = ROOT / "tests" / "fixtures"
        a, b = fx / "spike_manuscript.docx", fx / "review_manuscript.docx"
        for base, clean in ((a, b), (b, a)):
            for mode in ("color", "track_changes"):
                with self.subTest(clean=clean.name, mode=mode):
                    code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                    self.assertEqual(code, 0, log)
                    self.assertEqual(doc_xml(out).count("<w:commentReference"), doc_xml(clean).count("<w:commentReference"))

    def test_a_comment_keeps_its_place_in_a_changed_paragraph(self):
        base, clean = self.pair("a word changed under a comment")
        for mode in ("color", "track_changes"):
            with self.subTest(mode=mode):
                code, out, summary, log = self.run_tool(base, clean, "--mode", mode)
                self.assertEqual(code, 0, log)
                xml = doc_xml(out)
                start, end, ref = (xml.index(x) for x in ("<w:commentRangeStart", "<w:commentRangeEnd", "<w:commentReference"))
                self.assertLess(start, xml.index("reworded"))
                self.assertLess(xml.index("phrase"), end)
                self.assertLess(end, ref)


if __name__ == "__main__":
    unittest.main()
