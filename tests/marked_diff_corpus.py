"""Invented baseline/clean pairs for the marked-diff tool, and an own reading of its output.

Nothing here uses the tool's code. A document is a model (a list of blocks); the same model gives the docx file
and the structure the tool's output must show: after every mark is accepted the structure of the clean model,
after every mark is rejected the structure of the baseline model. Structure means every paragraph (empty ones
too) with its text, tabs and typed breaks, every table with its rows, cells and cell paragraphs, and the text
inside every bookmark.

A paragraph is {"k": "p", "id": ..., "pieces": [...]}. A piece is a string or a tuple:
("tab",), ("br", "line" | "page" | "column" | "all" | "none"), ("link", text), ("bm", name, text), ("b", text),
("fn",), ("ins", text) and ("del", text) (a change that is still pending in the input file).
A table is {"k": "t", "id": ..., "rows": [[cell, ...], ...]}; a cell is a list of piece lists (its paragraphs).
{"k": "bs", "name": ..., "n": ...} and {"k": "be", "n": ...} are the start and the end of a bookmark that stand
between paragraphs, as Word writes them when a bookmark covers whole paragraphs.
"""

import copy
import random
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

INSERT_RGB = "FF0000"
PAGE, COLUMN, LINE_ALL = "<page>", "<column>", "<line:all>"

_A = ["quiet", "narrow", "distant", "early", "plain", "heavy", "bright", "silent"]
_N = ["harbour", "orchard", "bridge", "lantern", "meadow", "ferry", "quarry", "garden"]
_V = ["shelters", "follows", "crosses", "carries", "watches", "divides", "gathers", "borders"]
_O = ["river", "valley", "station", "market", "tower", "village", "forest", "island"]
_T = ["before the rain", "in the long winter", "near the old road", "without any hurry", "during the harvest"]


def sentence(i):
    """A fixed invented sentence; different numbers share few words."""
    return "The %s %s %s the %s %s number %d." % (
        _A[i % 8], _N[(i * 3) % 8], _V[(i * 5) % 8], _O[(i * 7 + i // 8) % 8], _T[i % 5], i)


def text(*numbers):
    return " ".join(sentence(i) for i in numbers)


def P(ident, *pieces):
    return {"k": "p", "id": ident, "pieces": list(pieces)}


def T(ident, rows):
    return {"k": "t", "id": ident, "rows": rows}


def base_model():
    return [
        P("h1", "Opening notes"),
        P("a", text(0, 1, 2)),
        P("b", text(3, 4)),
        P("e", ),
        P("c", text(5, 6, 7)),
        T("t1", [[[[text(30)]], [[text(31)]]], [[[text(32)]], [[text(33)], [text(34)]]]]),
        P("d", text(8, 9)),
        P("h2", "Closing notes"),
        P("f", text(10, 11, 12)),
        P("g", text(13, 14)),
    ]


# ---------------------------------------------------------------- the model as a docx file
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), v)
    return e


def _run(text_):
    r = _el("w:r")
    t = _el("w:t")
    t.text = text_
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


_IDS = iter(range(500, 100000))


def _fill(p, pieces):
    for piece in pieces:
        if isinstance(piece, str):
            p.append(_run(piece))
            continue
        kind = piece[0]
        if kind == "tab":
            r = _el("w:r")
            r.append(_el("w:tab"))
            p.append(r)
        elif kind == "br":
            r = _el("w:r")
            attrs = {"line": {}, "page": {"type": "page"}, "column": {"type": "column"},
                     "all": {"clear": "all"}, "none": {"clear": "none"}}[piece[1]]
            r.append(_el("w:br", **attrs))
            p.append(r)
        elif kind == "link":
            link = _el("w:hyperlink", anchor="somewhere")
            link.append(_run(piece[1]))
            p.append(link)
        elif kind == "bm":
            n = str(next(_IDS))
            p.append(_el("w:bookmarkStart", id=n, name=piece[1]))
            p.append(_run(piece[2]))
            p.append(_el("w:bookmarkEnd", id=n))
        elif kind == "b":
            r = _run(piece[1])
            rpr = _el("w:rPr")
            rpr.append(_el("w:b"))
            r.insert(0, rpr)
            p.append(r)
        elif kind == "fn":
            r = _el("w:r")
            r.append(_el("w:footnoteReference", id="2"))
            p.append(r)
        elif kind in ("ins", "del"):
            w = _el("w:" + kind, id=str(next(_IDS)), author="Somebody", date="2031-01-02T03:04:05Z")
            r = _run(piece[1])
            if kind == "del":
                r[0].tag = qn("w:delText")
            w.append(r)
            p.append(w)
        else:
            raise ValueError(kind)


FOOTNOTES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:footnote w:id="2"><w:p><w:r><w:t>An invented note at the foot of the page.</w:t></w:r></w:p></w:footnote>'
    "</w:footnotes>"
)


def write(model, path, header=None):
    """Save the model as a docx file. With a ("fn",) piece a footnotes part is added."""
    d = Document()
    for block in model:
        if block["k"] == "p":
            _fill(d.add_paragraph()._p, block["pieces"])
        elif block["k"] in ("bs", "be"):
            attrs = {"id": str(block["n"]), "name": block["name"]} if block["k"] == "bs" else {"id": str(block["n"])}
            d.element.body.find(qn("w:sectPr")).addprevious(_el("w:bookmarkStart" if block["k"] == "bs" else "w:bookmarkEnd", **attrs))
        else:
            rows = block["rows"]
            table = d.add_table(rows=len(rows), cols=len(rows[0]))
            for i, row in enumerate(rows):
                for j, cell in enumerate(row):
                    tc = table.cell(i, j)
                    _fill(tc.paragraphs[0]._p, cell[0])
                    for pieces in cell[1:]:
                        _fill(tc.add_paragraph()._p, pieces)
    if header:
        d.sections[0].header.paragraphs[0].text = header
    d.save(path)
    if any(piece == ("fn",) for b in model if b["k"] == "p" for piece in b["pieces"]):
        _add_footnotes(path)
    return path


def _add_footnotes(path):
    with zipfile.ZipFile(path) as z:
        parts = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in parts:
            if info.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-'
                    b'officedocument.wordprocessingml.footnotes+xml"/></Types>')
            if info.filename == "word/_rels/document.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    b'relationships/footnotes" Target="footnotes.xml"/></Relationships>')
            z.writestr(info.filename, data)
        z.writestr("word/footnotes.xml", FOOTNOTES)


# ---------------------------------------------------------------- the structure a model stands for
def _text(pieces):
    out = ""
    for piece in pieces:
        if isinstance(piece, str):
            out += piece
        elif piece[0] == "tab":
            out += "\t"
        elif piece[0] == "br":
            out += {"line": "\n", "none": "\n", "page": PAGE, "column": COLUMN, "all": LINE_ALL}[piece[1]]
        elif piece[0] in ("link", "b", "ins"):
            out += piece[1]
        elif piece[0] == "bm":
            out += piece[2]
    return out


def structure_of(model):
    """(items in order, {bookmark name: text}); an item is ("p", text) or ("t", rows of cells of paragraph texts)."""
    items, marks, opened = [], {}, {}
    for block in model:
        if block["k"] == "p":
            items.append(("p", _text(block["pieces"])))
            marks.update({piece[1]: piece[2] for piece in block["pieces"] if not isinstance(piece, str) and piece[0] == "bm"})
            for name in opened.values():
                marks[name] += _text(block["pieces"])
        elif block["k"] == "bs":
            opened[block["n"]] = block["name"]
            marks[block["name"]] = ""
        elif block["k"] == "be":
            opened.pop(block["n"])
        else:
            items.append(("t", tuple(tuple(tuple(_text(p) for p in cell) for cell in row) for row in block["rows"])))
    return items, marks


# ---------------------------------------------------------------- an own reading of a marked file
def _on(el):
    return el is not None and el.get(qn("w:val")) not in ("0", "false", "off")


def _rpr_state(rpr):
    if rpr is None:
        return None
    if _on(rpr.find(qn("w:strike"))):
        return "del"
    colour = rpr.find(qn("w:color"))
    return "ins" if colour is not None and (colour.get(qn("w:val")) or "").upper() == INSERT_RGB else None


def _mark_state(rpr):
    if rpr is None:
        return None
    return "ins" if rpr.find(qn("w:ins")) is not None else "del" if rpr.find(qn("w:del")) is not None else None


def view(path, mode, side):
    """The body of a marked file with every mark accepted (side "accept") or rejected ("reject").

    Track changes: runs in w:ins / w:del, rows with w:trPr/w:ins|w:del, paragraph marks with w:pPr/w:rPr/w:ins|w:del.
    Colour: struck runs are deletions and runs in the insert colour insertions; a paragraph mark formatted that
    way is a deleted or inserted paragraph mark; a row whose paragraph marks are all deleted (or all inserted)
    is a deleted (inserted) row. Dropping a paragraph mark joins the paragraph with the one after it."""
    body = Document(path).element.body
    drop = "del" if side == "accept" else "ins"
    track = mode == "track_changes"

    def para_state(p):
        rpr = p.find(qn("w:pPr") + "/" + qn("w:rPr"))
        return _mark_state(rpr) if track else _rpr_state(rpr)

    def row_state(tr):
        if track:
            return _mark_state(tr.find(qn("w:trPr")))
        states = {para_state(p) for p in tr.iter(qn("w:p"))}
        return states.pop() if len(states) == 1 else None

    def run_state(r):
        if track:
            tags = [a.tag for a in r.iterancestors()]
            return "del" if qn("w:del") in tags else "ins" if qn("w:ins") in tags else None
        return _rpr_state(r.find(qn("w:rPr")))

    for tr in list(body.iter(qn("w:tr"))):
        if row_state(tr) == drop:
            tr.getparent().remove(tr)
    for tbl in list(body.iter(qn("w:tbl"))):
        if tbl.find(qn("w:tr")) is None:
            tbl.getparent().remove(tbl)
    for r in list(body.iter(qn("w:r"))):
        if run_state(r) == drop:
            r.getparent().remove(r)
    for p in list(body.iter(qn("w:p"))):
        if para_state(p) != drop:
            continue
        nxt, between = p.getnext(), []
        while nxt is not None and nxt.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
            between.append(nxt)  # a bookmark that starts or ends between two paragraphs does not keep them apart
            nxt = nxt.getnext()
        if nxt is None or nxt.tag != qn("w:p"):
            continue  # stays: the comparison then shows a paragraph too many
        at = 1 if len(nxt) and nxt[0].tag == qn("w:pPr") else 0
        for child in [c for c in p if c.tag != qn("w:pPr")] + between:  # the markers keep their place in the text
            nxt.insert(at, child)
            at += 1
        p.getparent().remove(p)
    return body


def _para_text(p):
    return "".join(_run_text(r) for r in p.iter(qn("w:r")))


def _run_text(r):
    out = ""
    for ch in r:
        if ch.tag in (qn("w:t"), qn("w:delText")):
            out += ch.text or ""
        elif ch.tag == qn("w:tab"):
            out += "\t"
        elif ch.tag == qn("w:br"):
            kind, clear = ch.get(qn("w:type")), ch.get(qn("w:clear"))
            if kind == "page":
                out += PAGE
            elif kind == "column":
                out += COLUMN
            elif clear in (None, "none"):
                out += "\n"
            else:
                out += "<line:%s>" % clear
    return out


def structure_read(body):
    """The same two values as structure_of, read from a body element."""
    items, marks, opened = [], {}, {}
    for child in body:
        if child.tag == qn("w:p"):
            items.append(("p", _para_text(child)))
        elif child.tag == qn("w:tbl"):
            items.append(("t", tuple(
                tuple(tuple(_para_text(p) for p in tc.findall(qn("w:p"))) for tc in tr.findall(qn("w:tc")))
                for tr in child.findall(qn("w:tr")))))
    for e in body.iter(qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:r")):
        if e.tag == qn("w:bookmarkStart"):
            opened[e.get(qn("w:id"))] = e.get(qn("w:name"))
            marks.setdefault(e.get(qn("w:name")), "")
        elif e.tag == qn("w:bookmarkEnd"):
            opened.pop(e.get(qn("w:id")), None)
        else:
            piece = _run_text(e)
            for name in opened.values():
                marks[name] += piece
    marks.pop("_GoBack", None)
    return items, marks


def without_page_breaks(items):
    """Colour mode shows a deleted page or column break as nothing (a struck break would still break the page)."""
    strip = lambda s: s.replace(PAGE, "").replace(COLUMN, "")
    return [("p", strip(x[1])) if x[0] == "p" else
            ("t", tuple(tuple(tuple(strip(p) for p in cell) for cell in row) for row in x[1])) for x in items]


def split_reject(items):
    """(body paragraphs, tables): where a table stands among the paragraphs is not compared on the reject side."""
    return [x for x in items if x[0] == "p"], [x for x in items if x[0] == "t"]


# ---------------------------------------------------------------- changes
def _at(model, ident):
    return next(i for i, b in enumerate(model) if b["id"] == ident)


def _reword(pieces, n=3, new="altered"):
    """Change one word of the first string piece."""
    for k, piece in enumerate(pieces):
        if isinstance(piece, str) and len(piece.split(" ")) > n:
            words = piece.split(" ")
            words[n] = new
            pieces[k] = " ".join(words)
            return
    raise ValueError("no text to change")


def reword(ident, n=3, new="altered"):
    return lambda m: _reword(m[_at(m, ident)]["pieces"], n, new)


def add_para(before, numbers, ident="new"):
    def op(m):
        new = P(ident, text(*numbers)) if numbers else P(ident, ("bm", "mark1", "Figure 3"), " " + text(6, 7))
        m.insert(len(m) if before is None else _at(m, before), new)
    return op


def del_para(ident):
    return lambda m: m.pop(_at(m, ident))


def split_para(ident):
    def op(m):
        i = _at(m, ident)
        first, rest = m[i]["pieces"][0].split(". ", 1)
        m[i]["pieces"][0] = first + "."
        m.insert(i + 1, P(ident + "2", rest))
    return op


def merge_paras(ident):
    def op(m):
        i = _at(m, ident)
        j = next(k for k in range(i + 1, len(m)) if m[k]["k"] == "p")
        m[i]["pieces"] = [_text(m[i]["pieces"]) + " " + _text(m[j]["pieces"])]
        m.pop(j)
    return op


def add_empty(before):
    return lambda m: m.insert(len(m) if before is None else _at(m, before), P("empty2"))


def replace_para(ident, numbers):
    def op(m):
        m[_at(m, ident)]["pieces"] = [text(*numbers)]
    return op


def move_para(ident, before):
    def op(m):
        block = m.pop(_at(m, ident))
        m.insert(_at(m, before), block)
    return op


def set_pieces(ident, *pieces):
    def op(m):
        m[_at(m, ident)]["pieces"] = list(pieces)
    return op


def table(ident, change):
    return lambda m: change(m[_at(m, ident)]["rows"])


def add_row(rows):
    rows.append([[[text(40 + j)]] for j in range(len(rows[0]))])


def del_row(rows):
    rows.pop()


def first_row_gone(rows):
    rows.pop(0)


def add_col(rows):
    for i, row in enumerate(rows):
        row.append([[text(44 + i)]])


def del_col(rows):
    for row in rows:
        row.pop()


def cell_reword(rows):
    _reword(rows[0][0][0])


def cell_add_para(rows):
    rows[0][1].append([text(46)])


def cell_del_para(rows):
    rows[1][1].pop()


def cell_del_first_para(rows):
    rows[1][1].pop(0)


def cell_replace(rows):
    rows[1][0][0] = [text(47)]


def add_table(before, ident="t2"):
    def op(m):
        m.insert(_at(m, before), T(ident, [[[[text(50)]], [[text(51)]]]]))
    return op


def move_sentence(src, dst):
    """Take the last sentence of one paragraph to the front of another."""
    def op(m):
        head, tail = m[_at(m, src)]["pieces"][0].rsplit(". ", 1)
        m[_at(m, src)]["pieces"][0] = head + "."
        m[_at(m, dst)]["pieces"][0] = tail + " " + m[_at(m, dst)]["pieces"][0]
    return op


def rewrite_keeping(ident, keep, numbers):
    """Rewrite a paragraph: one of its sentences stays, between new ones."""
    def op(m):
        old = m[_at(m, ident)]["pieces"][0].split(". ")
        kept = old[keep] if old[keep].endswith(".") else old[keep] + "."
        m[_at(m, ident)]["pieces"] = [text(numbers[0]) + " " + kept + " " + text(*numbers[1:])]
    return op


def only(*idents):
    """Drop every block but these."""
    def op(m):
        m[:] = [b for b in m if b["id"] in idents]
    return op


def to_front(ident):
    def op(m):
        m.insert(0, m.pop(_at(m, ident)))
    return op


def to_end(ident):
    def op(m):
        m.append(m.pop(_at(m, ident)))
    return op


def span(first, last, name="whole"):
    """A bookmark over whole paragraphs: its start before the block `first`, its end after the block `last`."""
    def op(m):
        m.insert(_at(m, last) + 1, {"k": "be", "id": name + "-end", "n": 77})
        m.insert(_at(m, first), {"k": "bs", "id": name + "-start", "name": name, "n": 77})
    return op


def both(*ops):
    def op(m):
        for o in ops:
            o(m)
    return op


def nothing(m):
    return None


def variant(*ops):
    """A baseline model: the base with these changes."""
    def make():
        m = base_model()
        for o in ops:
            o(m)
        return m
    return make


BR = lambda kind: set_pieces("c", text(5, 6), ("br", kind), text(7))  # noqa: E731
WITH_TAB = set_pieces("c", text(5), ("tab",), text(6, 7))
WITH_LINK = set_pieces("c", text(5) + " ", ("link", "a linked phrase"), " " + text(6, 7))
WITH_MARK = set_pieces("c", text(5) + " ", ("bm", "mark1", "Figure 3"), " " + text(6, 7))
WITH_BOLD = set_pieces("c", text(5) + " ", ("b", "a bold phrase"), " " + text(6, 7))
WITH_NOTE = set_pieces("d", text(8), ("fn",), " " + text(9))
TWO_TABLES = add_table("h2")
SPLIT_AT_MARK = set_pieces("c", text(5))
PENDING_INS = set_pieces("b", text(3), ("ins", " " + text(4)))
PENDING_DEL = set_pieces("b", text(3, 4), ("del", " " + text(20)))

# name -> (baseline model maker, change that gives the clean model, refusal pattern or None)
CASES = {
    "no change": (variant(), nothing, None),
    "a word changed": (variant(), reword("a"), None),
    "a word changed in two paragraphs": (variant(), both(reword("a"), reword("f")), None),
    "a paragraph added at the start": (variant(), add_para("h1", (20, 21)), None),
    "a paragraph added in the middle": (variant(), add_para("b", (20, 21)), None),
    "a paragraph added before a table": (variant(), add_para("t1", (20, 21)), None),
    "a paragraph added after a table": (variant(), add_para("d", (20, 21)), None),
    "a paragraph added at the end": (variant(), add_para(None, (20, 21)), None),
    "two paragraphs added at the end": (variant(), both(add_para(None, (20, 21)), add_para(None, (22, 23), "new2")), None),
    "a paragraph deleted at the start": (variant(), del_para("h1"), None),
    "a paragraph deleted in the middle": (variant(), del_para("b"), None),
    "a paragraph deleted before a table": (variant(), del_para("c"), None),
    "a paragraph deleted after a table": (variant(), del_para("d"), None),
    "a paragraph deleted at the end": (variant(), del_para("g"), None),
    "two paragraphs deleted at the end": (variant(), both(del_para("g"), del_para("f")), None),
    "a paragraph split in two": (variant(), split_para("a"), None),
    "the last paragraph split in two": (variant(), split_para("g"), None),
    "two paragraphs merged": (variant(), merge_paras("a"), None),
    "the last two paragraphs merged": (variant(), merge_paras("f"), None),
    "a heading merged into the next paragraph": (variant(), merge_paras("h2"), None),
    "an empty paragraph added": (variant(), add_empty("f"), None),
    "an empty paragraph added at the end": (variant(), add_empty(None), None),
    "an empty paragraph removed": (variant(), del_para("e"), None),
    "a paragraph replaced": (variant(), replace_para("b", (22, 23)), None),
    "the last paragraph replaced": (variant(), replace_para("g", (22, 23)), None),
    "a paragraph moved": (variant(), move_para("f", "a"), None),
    "a table row added": (variant(), table("t1", add_row), None),
    "a table row removed": (variant(), table("t1", del_row), None),
    "the first table row removed": (variant(), table("t1", first_row_gone), None),
    "a table column added": (variant(), table("t1", add_col), None),
    "a table column removed": (variant(), table("t1", del_col), None),
    "a word changed in a cell": (variant(), table("t1", cell_reword), None),
    "a paragraph added in a cell": (variant(), table("t1", cell_add_para), None),
    "the last paragraph of a cell deleted": (variant(), table("t1", cell_del_para), None),
    "the first paragraph of a cell deleted": (variant(), table("t1", cell_del_first_para), None),
    "the text of a cell replaced": (variant(), table("t1", cell_replace), None),
    "a table added": (variant(), add_table("h2"), None),
    "the second of two tables removed": (variant(TWO_TABLES), del_para("t2"), None),
    "the first of two tables removed": (variant(TWO_TABLES), del_para("t1"), None),
    "a line break in a changed paragraph": (variant(BR("line")), reword("c"), None),
    "a page break in a changed paragraph": (variant(BR("page")), reword("c"), None),
    "a column break in a changed paragraph": (variant(BR("column")), reword("c"), None),
    "a break that clears floats in a changed paragraph": (variant(BR("all")), reword("c"), None),
    "a break written with clear none": (variant(BR("none")), reword("c"), None),
    "a break written with clear none, nothing changed": (variant(BR("none")), nothing, None),
    "a page break added": (variant(), BR("page"), None),
    "a page break removed": (variant(BR("page")), set_pieces("c", text(5, 6, 7)), None),
    "a column break removed": (variant(BR("column")), set_pieces("c", text(5, 6, 7)), None),
    "a line break removed": (variant(BR("line")), set_pieces("c", text(5, 6) + " " + text(7)), None),
    "a page break turned into a column break": (variant(BR("page")), BR("column"), None),
    "a paragraph with a page break deleted": (variant(BR("page")), del_para("c"), None),
    "a tab in a changed paragraph": (variant(WITH_TAB), reword("c"), None),
    "a link in a changed paragraph": (variant(WITH_LINK), reword("c"), None),
    "bold text in a changed paragraph": (variant(WITH_BOLD), reword("c"), None),
    "a bookmark in a changed paragraph": (variant(WITH_MARK), both(WITH_MARK, reword("c")), None),
    "a bookmark in a paragraph that is split": (variant(), both(SPLIT_AT_MARK, add_para("t1", (), "c2")), None),
    "a pending insertion in the clean file": (variant(), both(PENDING_INS, reword("a")), None),
    "a pending deletion in the clean file": (variant(), both(PENDING_DEL, reword("a")), None),
    "a pending insertion in the baseline": (variant(PENDING_INS), reword("b"), None),
    "a footnote and a header, change elsewhere": (variant(WITH_NOTE), both(WITH_NOTE, reword("a")), None),
    "a footnote in the changed paragraph": (variant(WITH_NOTE), both(WITH_NOTE, reword("d")), r"footnoteReference"),
    "a footnote in a deleted paragraph": (variant(WITH_NOTE), del_para("d"), None),
    "a sentence moved to the next paragraph": (variant(), move_sentence("a", "b"), None),
    "a sentence moved across an empty paragraph": (variant(), move_sentence("b", "c"), None),
    "a paragraph rewritten around one sentence": (variant(), rewrite_keeping("a", 1, (20, 21, 22)), None),
    "two paragraphs merged and reworded": (variant(), both(merge_paras("a"), reword("a", 6, "changed")), None),
    "three paragraphs merged": (variant(), both(merge_paras("f"), del_para("h2"), merge_paras("d")), None),
    "a heading deleted and the next paragraph reworded": (variant(), both(del_para("h2"), reword("f")), None),
    "two paragraphs deleted around an empty one": (variant(), both(del_para("b"), del_para("c")), None),
    "an empty paragraph and its neighbour deleted": (variant(), both(del_para("e"), del_para("b")), None),
    "two empty paragraphs added": (variant(), both(add_empty("b"), add_empty("b")), None),
    "an empty paragraph added at the start": (variant(), add_empty("h1"), None),
    "all but one paragraph deleted": (variant(), only("a"), None),
    "everything new": (variant(only("a")), both(replace_para("a", (20, 21)), add_para(None, (22, 23))), None),
    "the only table removed": (variant(), del_para("t1"), None),
    "a table at the start of the document": (variant(to_front("t1")), reword("a"), None),
    "a table added at the start": (variant(), both(add_table("h1"), reword("a")), None),
    "a table at the start removed": (variant(to_front("t1")), del_para("t1"), None),
    "a table at the end of the document": (variant(to_end("t1")), reword("a"), None),
    "a table at the end removed": (variant(to_end("t1")), del_para("t1"), None),
    "a table moved to the end": (variant(), to_end("t1"), None),
    "a table moved between paragraphs": (variant(), move_para("t1", "g"), None),
    "the paragraph between two tables deleted": (variant(add_table("d")), del_para("d"), None),
    "a table and the paragraph before it deleted": (variant(), both(del_para("t1"), del_para("c")), None),
    "a table added and the paragraph before it split": (variant(), both(add_table("d"), split_para("c")), None),
    "a row added and a cell reworded": (variant(), both(table("t1", cell_reword), table("t1", add_row)), None),
    "a cell emptied": (variant(), table("t1", lambda rows: rows[0][0].__setitem__(0, [])), None),
    "an empty cell filled": (variant(table("t1", lambda rows: rows[0][0].__setitem__(0, []))), table("t1", lambda rows: rows[0][0].__setitem__(0, [text(48)])), None),
    "a paragraph of a cell split": (variant(table("t1", lambda rows: rows[0][0].__setitem__(0, [text(30, 35)]))), table("t1", lambda rows: rows[0].__setitem__(0, [[text(30)], [text(35)]])), None),
    "two paragraphs of a cell merged": (variant(), table("t1", lambda rows: rows[1].__setitem__(1, [[text(33, 34)]])), None),
    "a line break added": (variant(), BR("line"), None),
    "a tab removed": (variant(WITH_TAB), set_pieces("c", text(5, 6, 7)), None),
    "a link reworded": (variant(WITH_LINK), set_pieces("c", text(5) + " ", ("link", "a changed phrase"), " " + text(6, 7)), None),
    "a paragraph with a link deleted": (variant(WITH_LINK), del_para("c"), None),
    "a bookmark around text that is reworded": (variant(WITH_MARK), set_pieces("c", text(5) + " ", ("bm", "mark1", "Figure 4"), " " + text(6, 7)), None),
    "a paragraph with a bookmark deleted": (variant(WITH_MARK), del_para("c"), None),
    "a whitespace paragraph deleted": (variant(set_pieces("e", " ")), del_para("e"), None),
    "a paragraph that is only a page break deleted": (variant(set_pieces("e", ("br", "page"))), del_para("e"), None),
    "a paragraph added inside a bookmark over paragraphs": (variant(span("f", "g")), add_para("g", (20, 21)), None),
    "the last paragraph of a bookmark over paragraphs deleted": (variant(span("f", "g")), del_para("g"), None),
    "a paragraph added after a bookmark over paragraphs": (variant(span("f", "g")), add_para(None, (20, 21)), None),
    "paragraphs joined across the end of a bookmark": (variant(span("a", "a")), merge_paras("a"), None),
    "all but one paragraph and the table deleted": (variant(), only("a", "t1"), r"paragraph mark cannot show"),
    "a paragraph added between two tables": (variant(add_table("d")), add_para("t2", (20, 21)), r"paragraph mark cannot show"),
}

# Changes that can be combined freely: each touches another part of the base model.
MIXABLE = [
    reword("a"), reword("f", 5, "revised"), add_para("b", (20, 21)), add_para(None, (24, 25), "tail"),
    del_para("h2"), del_para("g"), split_para("c"), merge_paras("a"), add_empty("d"), del_para("e"),
    table("t1", add_row), table("t1", cell_reword), table("t1", cell_add_para), replace_para("d", (26, 27)),
    add_para("h1", (28,), "lead"),
]
CONFLICTS = [{2, 7}]  # pairs of MIXABLE changes that need the same block


def mixed_cases(count, seed=20311):
    """`count` combinations of two to four changes, the same ones for the same seed."""
    rng = random.Random(seed)
    out = {}
    while len(out) < count:
        picks = sorted(rng.sample(range(len(MIXABLE)), rng.randint(2, 4)))
        if any(pair <= set(picks) for pair in CONFLICTS):
            continue
        # A merge takes the next block, so it runs before the changes that add or remove blocks.
        order = sorted(picks, key=lambda k: (k != 7, k))
        out["mix " + "+".join(map(str, picks))] = (variant(), both(*[MIXABLE[k] for k in order]), None)
    return out


def models(case):
    make, change, refusal = case
    base = make()
    clean = copy.deepcopy(base)
    change(clean)
    return base, clean, refusal
