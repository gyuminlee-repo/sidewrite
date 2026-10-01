#!/usr/bin/env python3
"""Build a "marked" docx from a baseline docx and a clean docx: the clean document with every change shown.

Colour mode (default): deleted text is struck through in --delete-rgb, inserted text is in --insert-rgb, everything
else is unchanged. Track-changes mode (--mode track_changes): the same diff is written as Word revision marks
(inserted runs inside w:ins, deleted runs inside w:del with w:delText), with --author and --date on every mark. The
diff is the same in both modes; only the writer differs. The output is the clean docx with only word/document.xml
replaced.

Structure is marked like text. A paragraph mark that is new (a new paragraph, a paragraph split in two) or gone (a
deleted paragraph, two paragraphs joined, an empty paragraph that was removed) is an inserted or deleted paragraph
mark: w:ins / w:del in the mark's run properties in track-changes mode, the insert colour / strike and delete
colour on the mark in colour mode. Where a baseline paragraph ended inside what is now one clean paragraph, the
paragraph is split there and the first part ends in a deleted paragraph mark. A table whose rows or columns
changed, that is new, gone or stands at another place is shown as a deleted table and an inserted table: its rows
carry w:trPr/w:del or w:ins in track-changes mode, every paragraph mark in it the mark formatting in colour mode.
The same holds for a table whose cells are merged differently (w:gridSpan, w:vMerge, the grid).

How the diff is made: paragraphs are aligned first (exact matches anchor the alignment), changed paragraphs are
diffed word by word, and runs of paragraphs without a counterpart are re-paired sentence by sentence, so a rewritten
paragraph shows the sentences that stayed. A paragraph that moved (a re-sorted reference list, for example) is paired
with its new place when it is similar enough.

What is read how: pending tracked changes are carried out (accepted) in both files before they are compared, so the
output holds this tool's marks only; pending cell insertions, deletions and merges are refused. Text that is struck through in the
baseline is read as already deleted. In colour mode a clean file that already has struck-through text, or text in the
insert colour, or a paragraph mark formatted like one of this tool's, cannot be marked (the marks could not be told
apart) and is refused.

The result is checked before it is written: accepting every mark must give the clean document and rejecting every
mark the baseline, in text and in structure: every paragraph (empty ones too) with its text, tabs and breaks by kind
(line, page, column, clear) and with what it holds besides text, by kind and number (pictures, objects, fields,
footnote and endnote references, equations, symbols); every table with its grid, its rows, its
cells with their spans and merges, and the cell paragraphs; and on the accept side the text inside every bookmark,
the number of comment marks of every paragraph and the order of paragraphs and tables. Comments are no tracked
change: the output keeps the clean document's, and a comment of the baseline is not carried over. Not compared on the reject side: where a table stands
among the paragraphs, and in colour mode page and column breaks (a deleted one is left out there, because a struck
break would still break the page). The check itself is tested by damaging a copy. When a check fails nothing is
written and the exit code is 1.

What is refused (exit 2) instead of written wrong: nested tables; a changed paragraph with a field, a footnote
reference, a picture or another element that is not text, in the clean document or in the baseline, and a deleted
baseline paragraph with one (the writer cannot rebuild such an element; a paragraph that differs only in such an
element counts as changed, and one with the same text but other kinds of them is refused as well); a new paragraph
with an equation or a simple field (w:fldSimple), which stand outside a run and cannot be marked as inserted; a paragraph of one document before a table or at
the end of a cell or of the document where the other document has no paragraph at all (a paragraph mark cannot
show that); pending cell insertions, deletions and merges. The output is never one of the inputs, and an output or report file
that already exists is refused unless --overwrite is given (the write creates the name exclusively).

Needs lxml.

Usage: build_marked_diff.py --baseline OLD.docx --clean NEW.docx --out MARKED.docx
       [--insert-rgb FF0000] [--delete-rgb 808080] [--mode color|track_changes] [--author NAME] [--date ISO]
Exit code: 0 written, 1 a check failed (nothing written), 2 refused or unreadable input (the reason on stderr).
"""
import argparse
import collections
import copy
import difflib
import itertools
import json
import os
import random
import re
import shutil
import sys
import tempfile
import zipfile

from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# ---- constants (all tunable from the CLI where noted) ----
INSERT_RGB = 'FF0000'          # inserted text colour (--insert-rgb)
DELETE_RGB = '808080'          # deleted text colour, with strike (--delete-rgb)
MODE = 'color'                 # 'color' | 'track_changes' (--mode)
AUTHOR = 'Author'              # w:author of track-changes marks (--author)
DATE = '1970-01-01T00:00:00Z'  # w:date of track-changes marks (--date), a fixed xsd:dateTime so runs are reproducible
_REV_ID = itertools.count(1)    # w:id must be unique across the document
MERGE_GAP_TOKENS = 2           # common-token gap (non-space tokens) merged into one change (--merge-gap)
PARA_MATCH_MIN_RATIO = 0.40    # min word-token ratio to pair two non-identical paragraphs (--min-ratio)
SENT_MATCH_MIN_RATIO = 0.30    # min word-token ratio to pair two sentences inside a replace block (--sent-ratio)
SENT_ABBREV = {'al', 'fig', 'figs', 'vs', 'eg', 'ie', 'ca', 'approx', 'no', 'nos', 'eq', 'eqs', 'ref', 'refs', 'dr', 'cf',
               'resp', 'sp', 'spp', 'suppl', 'st', 'vol', 'pp', 'inc', 'ltd', 'co', 'etc', 'v', 'var', 'subsp', 'ed', 'eds'}
RESCUE_MIN_RATIO = 0.60        # min word-token ratio to re-pair a moved (cross-block) baseline/clean paragraph pair (--rescue-ratio)
EQ_ANCHOR_WEIGHT = 1.0         # weight of an exact-match anchor when it competes with rescued pairs (crossing anchors)
SHORT_PARA_CHARS = 80          # a wholly struck baseline paragraph shorter than this (heading) keeps its own line
# A break is one character of the paragraph text, so that the diff sees it. Each kind has its own character; all of
# them are white space to the tokenizer and none of them can stand in the text of an XML file.
BREAK_CHARS = {'\n': {}, '\f': {'type': 'page'}, '\x1d': {'type': 'column'},
               '\x1c': {'clear': 'left'}, '\x1e': {'clear': 'right'}, '\x1f': {'clear': 'all'}}
PAGE_BREAKS = '\f\x1d'            # page and column breaks: never written as deleted text (a struck break still breaks the page)
SEP_CHAR = '\v'                # in segment text only: a baseline paragraph ends here. Never written to a file: the paragraph is split there
ABSORB_MIN_DICE = 0.25         # min content-word Dice for an unmatched neighbouring baseline paragraph to be struck inside a rewritten clean paragraph (--absorb)
STOPWORDS = frozenset('the a an of and or to in on for with by from as at is are was were be been this that these those it its which who whom into than then also both each per not no'.split())
STRUCK_PARA_LIMIT = 0.90       # metric: paragraph counted when struck share exceeds this
SAMPLE_N, SAMPLE_SEED = 0, 1
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
             'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
             'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
             'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']
TOKEN_RE = re.compile(r'\w*[^\W\d_]\w*|\d+(?:\.\d+)*|\w+|[^\w\s]')
SAFE_CLEAN_TAGS = {'pPr', 'rPr', 'r', 't', 'tab', 'br', 'hyperlink', 'bookmarkStart', 'bookmarkEnd',
                   'proofErr', 'lastRenderedPageBreak'}
# Whatever a paragraph holds besides text, tabs and breaks (a picture, an object, a field mark or its code, a footnote
# or endnote reference, an equation, a symbol, ...) is one INLINE_CHAR in the paragraph text, so
# the diff sees that it is there, and it is counted by kind (Para.inline) for the structure check. The writer
# cannot make such an element, so it never writes this character: a paragraph that would need it is refused.
INLINE_CHAR = '\x02'
# Where a comment is anchored. A comment is no tracked change (Word keeps it when every change is rejected), so
# these are not content: the output has the clean document's comments, checked on the accept side, and a baseline
# comment is not carried over (the output has no comments part of the baseline); in a deleted paragraph it is
# counted as dropped.
COMMENT_TAGS = {'commentRangeStart', 'commentRangeEnd', 'commentReference'}
NOT_INLINE = {'proofErr', 'lastRenderedPageBreak', 'bookmarkStart', 'bookmarkEnd', 'hyperlink', 'smartTag', 'sdt',
              'sdtPr', 'sdtEndPr'} | COMMENT_TAGS   # wrappers and markers that are not content of their own


def q(tag):
    """Qualified WordprocessingML tag name."""
    return '{%s}%s' % (W_NS, tag)


def ln(el):
    """Local tag name of an element."""
    return etree.QName(el).localname


# ---------------------------------------------------------------- rPr helpers
def rpr_set(rpr, tag, **attrs):
    """Set a child of rPr, inserting it at its schema-ordered position."""
    for old in rpr.findall(q(tag)):
        rpr.remove(old)
    el = etree.Element(q(tag))
    for k, v in attrs.items():
        el.set(q(k), v)
    idx = RPR_ORDER.index(tag)
    pos = len(rpr)
    for i, ch in enumerate(rpr):
        name = ln(ch) if isinstance(ch.tag, str) else ''
        if (RPR_ORDER.index(name) if name in RPR_ORDER else len(RPR_ORDER)) > idx:
            pos = i
            break
    rpr.insert(pos, el)


def rpr_key(rpr, _cache={}):
    """Exclusive-c14n string of an rPr (None -> '')."""
    if rpr is None:
        return ''
    k = id(rpr)
    if k not in _cache or _cache[k][0] is not rpr:
        _cache[k] = (rpr, etree.tostring(rpr, method='c14n', exclusive=True))
    return _cache[k][1]


OFF_VALUES = ('0', 'false', 'off')   # w:val of an on/off property that switches it off; absent or anything else is on


def is_on(el):
    """True when an on/off property element (w:strike, w:b, ...) is there and not switched off by its w:val."""
    return el is not None and el.get(q('val')) not in OFF_VALUES


def is_struck(rpr):
    """True when an rPr switches strike or dstrike on (<w:strike w:val="0"/> is not struck)."""
    return rpr is not None and (is_on(rpr.find(q('strike'))) or is_on(rpr.find(q('dstrike'))))


def rpr_color(rpr):
    """Upper-case colour value of an rPr, or None."""
    if rpr is None:
        return None
    c = rpr.find(q('color'))
    return c.get(q('val')).upper() if c is not None and c.get(q('val')) else None


# ---------------------------------------------------------------- pending tracked changes in an input
# Both inputs are compared in their accepted state. A change that is still pending is therefore carried out on the
# XML before anything is compared, so no mark of an input is left among the marks this tool writes.
REV_GONE = ('del', 'moveFrom')            # content that accepting removes
REV_KEPT = ('ins', 'moveTo')              # content that accepting keeps: only the wrapper goes
REV_PROPERTY_CHANGES = ('rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tblPrExChange', 'tblGridChange',
                        'trPrChange', 'tcPrChange', 'numberingChange')   # hold the old formatting: dropped
REV_RANGE_MARKERS = ('moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
                     'customXmlInsRangeStart', 'customXmlInsRangeEnd', 'customXmlMoveToRangeStart',
                     'customXmlMoveToRangeEnd')                          # dropped
REV_REFUSED = ('cellIns', 'cellDel', 'cellMerge', 'customXmlDelRangeStart', 'customXmlDelRangeEnd',
               'customXmlMoveFromRangeStart', 'customXmlMoveFromRangeEnd')   # accepting these is not implemented
REV_PROPERTY_PARENTS = ('rPr', 'numPr', 'trPr', 'ctrlPr')   # here ins/del is a flag on the parent, not a wrapper


def resolve_revisions(xml, side='accept'):
    """(document.xml with every tracked change accepted or, with side='reject', rejected; number of revision
    elements carried out).

    The bytes are returned as they came when the document has no tracked change. Accepting: deleted and moved-away
    content is removed, inserted and moved-in content is unwrapped, a paragraph whose mark is deleted is joined with
    the paragraph after it (which keeps its own properties, as in Word), a deleted table row is removed. Rejecting
    is the same with the two kinds exchanged (and w:delText turned back into w:t). On both sides recorded old
    formatting (w:rPrChange and the like) and move range markers are dropped: the tool compares text and
    structure, not formatting. Pending cell insertions, deletions and merges are refused."""
    root = etree.fromstring(xml)
    gone, kept = (REV_GONE, REV_KEPT) if side == 'accept' else (REV_KEPT, REV_GONE)
    names = set(REV_GONE + REV_KEPT + REV_PROPERTY_CHANGES + REV_RANGE_MARKERS + REV_REFUSED)
    found = [e for e in root.iter() if isinstance(e.tag, str) and e.tag.startswith('{%s}' % W_NS) and ln(e) in names]
    if not found:
        return xml, 0
    refused = sorted({ln(e) for e in found if ln(e) in REV_REFUSED})
    assert not refused, 'pending tracked changes of a kind that is not handled: %s; accept or reject them in Word first' % ', '.join(refused)

    def drop(e):
        if e.getparent() is not None:
            e.getparent().remove(e)

    flags = [e for e in found if ln(e) in gone + kept and ln(e.getparent()) in REV_PROPERTY_PARENTS]
    joined = []
    for e in flags:
        owner, kind = e.getparent(), ln(e)
        drop(e)
        if kind in kept:
            continue
        assert ln(owner) != 'ctrlPr', 'a pending change inside an equation is not handled; accept or reject it in Word first'
        if ln(owner) == 'trPr':                       # a row that goes
            row = owner.getparent()
            tbl = row.getparent()
            drop(row)
            if tbl is not None and tbl.find(q('tr')) is None:
                drop(tbl)
        elif ln(owner) == 'rPr' and ln(owner.getparent()) == 'pPr':   # a paragraph mark that goes
            joined.append(owner.getparent().getparent())
    flagged = {id(e) for e in flags}
    wrappers = [e for e in found if ln(e) in gone + kept and id(e) not in flagged]
    for e in wrappers:
        if ln(e) in gone:
            drop(e)
    for e in wrappers:
        if ln(e) in kept and e.getparent() is not None:
            for t in e.iter(q('delText')):
                t.tag = q('t')
            for ch in list(e):
                e.addprevious(ch)
            drop(e)
    for e in found:
        if ln(e) in REV_PROPERTY_CHANGES + REV_RANGE_MARKERS:
            drop(e)
    for para in joined:                               # document order: a chain ends in the last one
        if para.getparent() is not None:
            join_with_next(para)                      # no next paragraph (end of a cell, a table): the mark cannot go
    for e in [x for x in root.iter(q('rPr')) if len(x) == 0 and ln(x.getparent()) == 'pPr']:
        drop(e)                                       # a paragraph-mark rPr that only held the flag
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True), len(found)


def accept_revisions(xml):
    """Both inputs are read in this state: every pending tracked change carried out."""
    return resolve_revisions(xml, 'accept')


# ---------------------------------------------------------------- paragraph reading
def break_char(br):
    """The character that stands for a w:br: by type (page, column) and, for a line break, by its clear attribute.
    clear="none" is the default and reads like no attribute."""
    kind, clear = br.get(q('type')), br.get(q('clear'))
    want = {'type': kind} if kind in ('page', 'column') else {'clear': clear} if clear in ('left', 'right', 'all') else {}
    return next(c for c, attrs in BREAK_CHARS.items() if attrs == want)


class Para:
    """Text of a paragraph with per-character format handles (key, rPr element, hyperlink element)."""
    def __init__(self):
        self.text = ''
        self.fmt = []
        self.special = collections.Counter()
        self.inline = collections.Counter()   # the elements that stand as INLINE_CHAR in the text, by kind
        self.loose = collections.Counter()    # those of them that stand outside a run (a simple field, an equation)
        self.marks = []     # bookmark markers: (element, text offset it sits at, enclosing hyperlink or None)


def read_para(p, skip_rpr=is_struck, side='accept'):
    """Read a paragraph; runs whose rPr satisfies skip_rpr are ignored.

    side selects how revision marks are read: 'accept' (w:ins kept, w:del dropped: the default, also the only reading
    a colour-mode file needs), 'reject' (w:del kept as text, w:ins dropped), 'ins' / 'del' (only runs inside that mark)."""
    para = Para()
    parts = []
    interned = {}

    def add_run(r, link):
        rpr = r.find(q('rPr'))
        if skip_rpr(rpr):
            return
        k = (rpr_key(rpr), id(link))
        if k not in interned:
            interned[k] = (rpr_key(rpr), rpr, link)
        f = interned[k]
        for ch in r:
            t = ln(ch)
            if t == 'rPr':
                continue
            if t == 't' or t == 'delText':
                s = ch.text or ''
            elif t == 'tab':
                s = '\t'
            elif t == 'br':
                s = break_char(ch)
            else:
                para.special[t] += 1
                if t in NOT_INLINE:
                    continue
                para.inline[t] += 1
                s = INLINE_CHAR
            parts.append(s)
            para.fmt.extend([f] * len(s))

    def inline(t, link):
        """A paragraph-level element that is not text: one INLINE_CHAR."""
        para.inline[t] += 1
        para.loose[t] += 1
        parts.append(INLINE_CHAR)
        para.fmt.append(('', None, link))

    def walk(node, link, mark=None):
        for ch in node:
            if not isinstance(ch.tag, str):
                continue
            t = ln(ch)
            if t == 'r':
                if side in ('accept', 'reject') or mark == side:
                    add_run(ch, link)
            elif t == 'hyperlink':
                para.special['hyperlink'] += 1
                walk(ch, ch, mark)
            elif t == 'ins':
                if side in ('accept', 'ins'):
                    walk(ch, link, 'ins')
            elif t == 'del':
                if side in ('reject', 'del'):
                    walk(ch, link, 'del')
            elif t in ('smartTag', 'fldSimple', 'sdt', 'sdtContent'):
                if t != 'sdtContent':
                    para.special[t] += 1
                if t == 'fldSimple':
                    inline(t, link)
                walk(ch, link, mark)
            elif t in ('moveFrom', 'pPr'):
                continue
            else:
                para.special[t] += 1
                if t in ('bookmarkStart', 'bookmarkEnd'):
                    para.marks.append((ch, len(para.fmt), link))
                elif t not in NOT_INLINE:
                    inline(t, link)

    walk(p, None)
    para.text = ''.join(parts)
    return para


# ---------------------------------------------------------------- tokenising and diff
def toks(text):
    """Tokens (string, start, end): words, numbers with decimals, single punctuation; whitespace is skipped."""
    return [(m.group(), m.start(), m.end()) for m in TOKEN_RE.finditer(text)]


def word_ratio(a, b):
    """Similarity of two paragraphs on their non-space token lists."""
    ta, tb = [t[0] for t in toks(a)], [t[0] for t in toks(b)]
    if not ta or not tb:
        return 0.0
    sm = difflib.SequenceMatcher(None, ta, tb, autojunk=False)
    if sm.real_quick_ratio() < PARA_MATCH_MIN_RATIO or sm.quick_ratio() < PARA_MATCH_MIN_RATIO:
        return 0.0
    return sm.ratio()


def diff_ops(bt, ct, merge_gap=None):
    """Word-level ops [(kind, s, e)]: eq/ins index clean text, del indexes baseline text."""
    merge_gap = MERGE_GAP_TOKENS if merge_gap is None else merge_gap
    tb, tc = toks(bt), toks(ct)
    opc = difflib.SequenceMatcher(None, [t[0] for t in tb], [t[0] for t in tc], autojunk=False).get_opcodes()
    groups = []
    for k, (tag, i1, i2, j1, j2) in enumerate(opc):
        if tag == 'equal' and 0 < k < len(opc) - 1 and (i2 - i1) <= merge_gap:
            tag = 'replace'
        if tag != 'equal' and groups and groups[-1][0] != 'equal':
            groups[-1][2], groups[-1][4] = i2, j2
        else:
            groups.append([tag, i1, i2, j1, j2])
    ops = []

    def add(kind, s, e):
        if e > s:
            if ops and ops[-1][0] == kind and ops[-1][2] == s:
                ops[-1] = (kind, ops[-1][1], e)
            else:
                ops.append((kind, s, e))

    pb = pc = 0  # end of previous token in baseline / clean
    for tag, i1, i2, j1, j2 in groups:
        if tag == 'equal':
            for a, b in zip(range(i1, i2), range(j1, j2)):
                wb, wc = bt[pb:tb[a][1]], ct[pc:tc[b][1]]
                if wb == wc:
                    add('eq', pc, tc[b][2])
                else:
                    add('del', pb, tb[a][1])
                    add('ins', pc, tc[b][1])
                    add('eq', tc[b][1], tc[b][2])
                pb, pc = tb[a][2], tc[b][2]
        else:
            if i2 > i1:
                add('del', pb, tb[i2 - 1][2])
                pb = tb[i2 - 1][2]
            if j2 > j1:
                add('ins', pc, tc[j2 - 1][2])
                pc = tc[j2 - 1][2]
    if bt[pb:] == ct[pc:]:
        add('eq', pc, len(ct))
    else:
        add('del', pb, len(bt))
        add('ins', pc, len(ct))
    return ops


def segments(ops, b, c):
    """Turn ops into merged segments [(kind, text, key, rpr, link)]; deleted text never keeps a hyperlink, and in colour mode
    no page or column break either (a struck break would still break the page; inside w:del it does not)."""
    segs = []
    for kind, s, e in ops:
        src = b if kind == 'del' else c
        for i in range(s, e):
            ch = src.text[i]
            if kind == 'del' and ch in PAGE_BREAKS and MODE == 'color':
                continue
            key, rpr, link = src.fmt[i]
            link = None if kind == 'del' else link
            if segs and segs[-1][0] == kind and segs[-1][2] == key and segs[-1][4] is link:
                segs[-1][1] += ch
            else:
                segs.append([kind, ch, key, rpr, link])
    return segs


def marked_string(segs):
    """Readable form: deletions [-..-], insertions {+..+}."""
    return ''.join({'eq': '%s', 'del': '[-%s-]', 'ins': '{+%s+}'}[k] % t.replace('\t', '<TAB>').replace('\n', '<BR>').replace(SEP_CHAR, '<PARA>')
                   for k, t, *_ in segs)


# ---------------------------------------------------------------- run building
def rev_mark(kind):
    """Empty <w:ins>/<w:del> element with a fresh id, the author and the date (track_changes mode)."""
    e = etree.Element(q(kind))
    e.set(q('id'), str(next(_REV_ID)))
    e.set(q('author'), AUTHOR)
    e.set(q('date'), DATE)
    return e


def make_run(kind, text, rpr, styles):
    """Build a w:r with the clean/baseline rPr deep-copied and the diff overlay applied.

    color mode: 'del' = strike + DELETE_RGB, 'ins' = INSERT_RGB. track_changes mode: the run keeps its formatting and is
    wrapped in <w:del> (text as w:delText) or <w:ins>; the wrapper is what gets returned."""
    tc = MODE == 'track_changes'
    r = etree.Element(q('r'))
    new = copy.deepcopy(rpr) if rpr is not None else etree.Element(q('rPr'))
    if kind == 'del':
        st = new.find(q('rStyle'))
        if st is not None and styles is not None and st.get(q('val')) not in styles:
            new.remove(st)
        if not tc:
            rpr_set(new, 'strike')
            rpr_set(new, 'color', val=DELETE_RGB)
    elif kind == 'ins' and not tc:
        rpr_set(new, 'color', val=INSERT_RGB)
    if len(new):
        r.append(new)
    ttag = 'delText' if (tc and kind == 'del') else 't'
    for part in re.split('([\t%s%s%s])' % (''.join(BREAK_CHARS), SEP_CHAR, INLINE_CHAR), text):
        if part == '':
            continue
        if part == '\t':
            r.append(etree.Element(q('tab')))
        elif part in BREAK_CHARS:
            br = etree.SubElement(r, q('br'))
            for k, v in BREAK_CHARS[part].items():
                br.set(q(k), v)
        elif part == SEP_CHAR:
            raise AssertionError('a paragraph separator reached the run writer')   # rebuild_para splits there
        elif INLINE_CHAR in part:
            raise AssertionError('an element that is not text reached the run writer')   # refused before, see refuse_inline
        else:
            t = etree.SubElement(r, q(ttag))
            t.text = part
            t.set(XML_SPACE, 'preserve')
    if tc and kind in ('ins', 'del'):
        w = rev_mark(kind)
        w.append(r)
        return w
    return r


def build_children(segs, styles):
    """Run/hyperlink elements for segments; consecutive segments sharing a clean hyperlink are re-wrapped."""
    out, cur, cur_link = [], None, None
    for kind, text, key, rpr, link in segs:
        # a bookmark marker from place_marks is the element itself, and stays in its hyperlink like a run
        r = text if kind == 'mark' else make_run(kind, text, rpr, styles)
        if link is not None:
            if cur is None or cur_link is not link:
                cur = copy.deepcopy(link)
                for ch in list(cur):
                    cur.remove(ch)
                out.append(cur)
                cur_link = link
            cur.append(r)
        else:
            cur, cur_link = None, None
            out.append(r)
    return out


def colorize_in_place(p, rgb=None):
    """Mark the whole content of a clean-only paragraph as inserted (structure kept): every run, also one that
    holds a picture, a field mark or a footnote reference instead of text. Colour mode applies the insert colour
    to the run, track_changes mode wraps the run in <w:ins>. What stands outside a run (Para.loose) cannot be
    marked this way; the caller refuses such a paragraph."""
    n = 0
    for r in list(p.iter(q('r'))):
        if any(a.tag == q('r') for a in r.iterancestors()):
            continue                               # inside a text box of a picture: its outer run is marked
        if MODE == 'track_changes':
            w = rev_mark('ins')
            r.addprevious(w)
            w.append(r)
        else:
            rpr = r.find(q('rPr'))
            if rpr is None:
                rpr = etree.Element(q('rPr'))
                r.insert(0, rpr)
            rpr_set(rpr, 'color', val=rgb or INSERT_RGB)
        n += 1
    return n


# ---------------------------------------------------------------- paragraph marks
# A paragraph ends in a paragraph mark, and that mark can be new or gone like any text: a new paragraph, a paragraph
# split in two, two paragraphs joined, a deleted paragraph. The mark's own run properties (w:pPr/w:rPr) say so:
# track_changes mode puts <w:ins> / <w:del> there, as Word does; colour mode formats the mark like inserted or
# deleted text. Dropping a paragraph mark joins the paragraph with the one after it, which keeps its properties.
def para_mark_rpr(p, create=False):
    """The w:rPr of a paragraph's mark, or None."""
    ppr = p.find(q('pPr'))
    if ppr is None:
        if not create:
            return None
        ppr = etree.Element(q('pPr'))
        p.insert(0, ppr)
    rpr = ppr.find(q('rPr'))
    if rpr is None and create:
        rpr = etree.Element(q('rPr'))
        pos = next((i for i, ch in enumerate(ppr) if ln(ch) in ('sectPr', 'pPrChange')), len(ppr))
        ppr.insert(pos, rpr)
    return rpr


def para_mark(p):
    """'ins' | 'del' | None: the state of a paragraph's mark in a marked file."""
    rpr = para_mark_rpr(p)
    if rpr is None:
        return None
    if MODE == 'track_changes':
        return 'ins' if rpr.find(q('ins')) is not None else 'del' if rpr.find(q('del')) is not None else None
    return 'del' if is_deleted_rpr(rpr) else 'ins' if is_inserted_rpr(rpr) else None


def set_para_mark(p, kind):
    """Mark a paragraph's mark as inserted ('ins') or deleted ('del')."""
    rpr = para_mark_rpr(p, create=True)
    if MODE == 'track_changes':
        rpr.insert(0, rev_mark(kind))
    elif kind == 'del':
        rpr_set(rpr, 'strike')
        rpr_set(rpr, 'color', val=DELETE_RGB)
    else:
        rpr_set(rpr, 'color', val=INSERT_RGB)


def set_row_mark(tr, kind):
    """track_changes mode: mark a table row as inserted or deleted (w:trPr/w:ins|w:del)."""
    trpr = tr.find(q('trPr'))
    if trpr is None:
        trpr = etree.Element(q('trPr'))
        tr.insert(1 if len(tr) and ln(tr[0]) == 'tblPrEx' else 0, trpr)
    pos = next((i for i, ch in enumerate(trpr) if ln(ch) == 'trPrChange'), len(trpr))
    trpr.insert(pos, rev_mark(kind))


BETWEEN_PARAGRAPHS = ('bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd', 'commentRangeStart',
                      'commentRangeEnd')   # may stand between two paragraphs without keeping them apart


def after(p):
    """What follows a paragraph: the next paragraph, a table, or whatever else ends the run of paragraphs (None at
    the end). Range markers between two paragraphs are passed over."""
    nxt = p.getnext()
    while nxt is not None and isinstance(nxt.tag, str) and ln(nxt) in BETWEEN_PARAGRAPHS:
        nxt = nxt.getnext()
    return nxt


def join_with_next(p):
    """Drop a paragraph mark: the content goes to the front of the next paragraph, and so do the range markers
    that stood between the two (they keep their place in the text). False when there is no next paragraph."""
    nxt = after(p)
    if nxt is None or nxt.tag != q('p'):
        return False
    between, e = [], p.getnext()
    while e is not nxt:
        between.append(e)
        e = e.getnext()
    at = 1 if len(nxt) and nxt[0].tag == q('pPr') else 0
    for ch in [c for c in p if c.tag != q('pPr')] + between:
        nxt.insert(at, ch)
        at += 1
    p.getparent().remove(p)
    return True


def swap_ppr(a, b):
    """Exchange the paragraph properties of two paragraphs."""
    pa, pb = a.find(q('pPr')), b.find(q('pPr'))
    for el, ppr in ((a, pa), (b, pb)):
        if ppr is not None:
            el.remove(ppr)
    if pb is not None:
        a.insert(0, pb)
    if pa is not None:
        b.insert(0, pa)


def ends(texts):
    """End offset of each text in their concatenation."""
    out, tot = [], 0
    for t in texts:
        tot += len(t)
        out.append(tot)
    return out


def plan_marks(seq, cps, B, C, stats, new_tables=()):
    """Give every paragraph of one container its mark, so that both documents' paragraphs come back.

    seq: the output paragraphs in document order (the clean ones, the deleted ones, the pieces split off where a
    baseline paragraph ended inside a clean one); cps: the clean paragraphs among them; B, C: all baseline and all
    clean paragraphs of the container, empty ones included.

    Every clean paragraph must end where a paragraph ends once the deletions are gone (accept), every baseline
    paragraph where one ends once the insertions are gone (reject). A paragraph end of the output is a point with
    two offsets, one in the accepted text and one in the rejected text. A point that ends a paragraph in both
    documents keeps a plain mark, one that ends only a clean paragraph gets an inserted mark, one that ends only a
    baseline paragraph a deleted mark, and one that ends neither is joined with the paragraph after it. Baseline
    paragraph ends that have no point (empty baseline paragraphs) get a new empty paragraph with a deleted mark.

    A mark that is dropped joins its paragraph with the next one, so the last paragraph before something that is
    not a paragraph (the end of a cell, a table) must keep a plain mark. When it would not, the role moves there
    from an earlier point at the same offset (what Word does when the last paragraph is deleted: the mark before
    it is the deleted one), and the two paragraphs exchange their properties so the accepted one keeps the clean
    properties.

    What cannot be written is refused: a paragraph of one document before a table, or at the end of a cell or of
    the document, where the other document has no paragraph at all (a file whose body ends in a table, for
    example). new_tables: ids of the clean tables that have no baseline table; an inserted paragraph mark right
    before one of them is fine, the table goes when it is rejected.

    Returns, for every baseline paragraph, the output paragraph that ends it."""
    if not seq:
        return []
    side = no_page_breaks if MODE == 'color' else (lambda t: t)
    acc = [read_accept(p).text for p in seq]
    rej = [side(read_reject(p).text) for p in seq]
    want_c, want_b = [x.text for x in C], [side(x.text) for x in B]
    assert ''.join(acc) == ''.join(want_c), 'the marked paragraphs do not add up to the clean text of this part'
    assert ''.join(rej) == ''.join(want_b), 'the marked paragraphs do not add up to the baseline text of this part'
    clean = {id(p) for p in cps}
    pts = [{'el': p, 'a': a, 'r': r, 'C': id(p) in clean, 'B': False} for p, a, r in zip(seq, ends(acc), ends(rej))]
    need = collections.Counter(ends(want_b))
    for y in sorted(need):
        here = [x for x in pts if x['r'] == y]
        assert here or y == 0, 'a baseline paragraph ends inside a paragraph of the marked text'
        made = [x for x in here if not x['C']][:need[y]]
        for x in made + [x for x in here if x['C']][:need[y] - len(made)]:
            x['B'] = True
        for _ in range(need[y] - sum(x['B'] for x in here)):   # empty baseline paragraphs without a counterpart
            new = etree.Element(q('p'))
            if here:
                at = max(i for i, x in enumerate(pts) if x['r'] == y)
                pts[at]['el'].addnext(new)
                pts.insert(at + 1, {'el': new, 'a': pts[at]['a'], 'r': y, 'C': False, 'B': True})
            else:
                pts[0]['el'].addprevious(new)
                pts.insert(0, {'el': new, 'a': 0, 'r': 0, 'C': False, 'B': True})
            stats['n_del_empty'] += 1
    for i, x in enumerate(pts):
        nxt = after(x['el'])
        if (x['C'] and x['B']) or (nxt is not None and nxt.tag == q('p')):
            continue
        give, same = ('B', 'r') if x['C'] else ('C', 'a')     # the role this point lacks, and the offset it shares
        j = i - 1
        while j >= 0 and after(pts[j]['el']) is pts[j + 1]['el'] and pts[j][same] == x[same]:
            if pts[j][give]:
                pts[j][give], x[give] = False, True
                if give == 'C':
                    swap_ppr(pts[j]['el'], x['el'])
                break
            j -= 1
        if not (x['C'] and x['B']) and not (x['C'] and nxt is not None and id(nxt) in new_tables):
            where = 'a table' if nxt is not None and nxt.tag == q('tbl') else 'the end of this part'
            one, other = ('clean document', 'baseline') if x['C'] else ('baseline', 'clean document')
            raise AssertionError('the %s has a paragraph before %s where the %s has none; a paragraph mark cannot '
                                 'show that' % (one, where, other))
    owners = [x['el'] for x in pts if x['B']]
    for x in pts:
        if x['C'] and x['B']:
            continue
        if not x['C'] and not x['B']:
            assert join_with_next(x['el']), 'a piece of a paragraph is left without a place'
            continue
        set_para_mark(x['el'], 'ins' if x['C'] else 'del')
        stats['para_marks']['ins' if x['C'] else 'del'] += 1
    return owners


# ---------------------------------------------------------------- paragraph-level operations
def place_marks(segs, marks):
    """Segments with the clean paragraph's bookmark markers put back where they were in the clean text.

    A marker sits at an offset of the clean text; unchanged and inserted segments are cut there. Deleted text has no
    place in the clean text: it is shown at an offset too, and at that offset a bookmark end goes before it and a
    bookmark start after it, so deleted text next to a bookmark stays outside it. A bookmark that is empty in the
    clean text (start and end at one offset, in this order) stays empty and goes before the deleted text. What this
    cannot do is keep out text that was deleted between the two ends: bookmarks_holding_deletions reports that."""
    if not marks:
        return segs
    pre, post = collections.defaultdict(list), collections.defaultdict(list)
    for off in sorted({o for _, o, _ in marks}):
        here = [(e, link) for e, o, link in marks if o == off]
        seen, empty = set(), set()
        for e, _ in here:
            if ln(e) == 'bookmarkStart':
                seen.add(e.get(q('id')))
            elif e.get(q('id')) in seen:
                empty.add(e.get(q('id')))
        for e, link in here:
            first = ln(e) == 'bookmarkEnd' or e.get(q('id')) in empty
            (pre if first else post)[off].append(['mark', e, None, None, link])
    out, pos = [], 0

    def emit(off, starts=True):
        out.extend(pre.pop(off, []))
        if starts:
            out.extend(post.pop(off, []))

    for seg in segs:
        kind, text = seg[0], seg[1]
        if kind == 'del':
            emit(pos, starts=False)
            out.append(seg)
            continue
        emit(pos)
        done = 0
        for off in sorted(o for o in set(pre) | set(post) if pos < o < pos + len(text)):
            out.append([kind, text[done:off - pos]] + list(seg[2:]))
            emit(off)
            done = off - pos
        out.append([kind, text[done:]] + list(seg[2:]))
        pos += len(text)
    for off in sorted(set(pre) | set(post)):
        emit(off)
    return out


def bookmarks_holding_deletions(root):
    """{bookmark name: characters of deleted text shown between its start and its end}, in document order.

    Only colour mode needs this: there deleted text is ordinary struck text, so a bookmark around it holds more
    than it did (a cross-reference to the bookmark would show the struck words). In track_changes mode a w:del
    inside a bookmark is what Word writes itself, and accepting the changes gives the clean range back."""
    opened, held = {}, collections.OrderedDict()
    for e in root.iter(q('bookmarkStart'), q('bookmarkEnd'), q('r')):
        if ln(e) == 'bookmarkStart':
            opened[e.get(q('id'))] = e.get(q('name')) or ''
        elif ln(e) == 'bookmarkEnd':
            opened.pop(e.get(q('id')), None)
        elif opened and run_kind(e) == 'del':
            n = sum(len(t.text or '') for t in e.iter(q('t'), q('delText')))
            for name in opened.values():
                held[name] = held.get(name, 0) + n
    return {k: v for k, v in held.items() if v}


def rebuild_para(p, b, c, styles, stats, segs=None, ppr_of=None):
    """Rewrite a clean paragraph as a word-level diff against its baseline counterpart (or prebuilt segments).

    Where the segments say that a baseline paragraph ends inside this paragraph (SEP_CHAR), the text before that
    point becomes a paragraph of its own in front of p, with the properties ppr_of gives for that baseline
    paragraph. Returns those new paragraphs; plan_marks gives them their deleted paragraph marks."""
    bad = {k: v for k, v in c.special.items() if k not in SAFE_CLEAN_TAGS and k != 'lastRenderedPageBreak'}
    assert not bad, 'unsupported element in changed clean paragraph %r: %s' % (c.text[:60], bad)
    if segs is None:
        segs = segments(diff_ops(b.text, c.text), b, c)
    pieces, cur = [], []
    for seg in place_marks(segs, c.marks):
        if seg[0] == 'del' and SEP_CHAR in seg[1]:
            assert seg[1].strip(SEP_CHAR) == '', 'a paragraph separator inside deleted text'
            pieces += [(cur, seg[2][1])] + [([], seg[2][1])] * (len(seg[1]) - 1)
            cur = []
        else:
            cur.append(seg)
    for ch in list(p):
        if ln(ch) != 'pPr':
            p.remove(ch)
    new = []
    for part, bi in pieces:
        x = etree.Element(q('p'))
        ppr = ppr_of(bi) if ppr_of else None
        if ppr is not None and len(ppr):
            x.append(ppr)
        for e in build_children(part, styles):
            x.append(e)
        p.addprevious(x)
        new.append(x)
    for e in build_children(cur, styles):
        p.append(e)
    stats['marked'].append((b if isinstance(b, str) else b.text, c.text, marked_string(segs)))
    stats['exp_ins'].append(''.join(t for k, t, *_ in segs if k == 'ins'))
    stats['exp_del'].append(''.join(t for k, t, *_ in segs if k == 'del'))
    return new


def deleted_ppr(bp, styles, numids, stats):
    """Paragraph properties for baseline text shown in the clean document: the baseline's, without structure
    elements and without a style or a numbering the clean document does not have. None when there are none."""
    ppr = bp.find(q('pPr'))
    if ppr is not None:
        ppr = copy.deepcopy(ppr)
        for tag in ('sectPr', 'pageBreakBefore', 'rPr', 'pPrChange'):
            for e in ppr.findall(q(tag)):
                ppr.remove(e)
        ps = ppr.find(q('pStyle'))
        if ps is not None and styles is not None and ps.get(q('val')) not in styles:
            ppr.remove(ps)
            stats['dropped_pstyle'] += 1
        npr = ppr.find(q('numPr'))
        if npr is not None:
            ni = npr.find(q('numId'))
            if numids is not None and (ni is None or ni.get(q('val')) not in numids):
                ppr.remove(npr)
                stats['dropped_numpr'] += 1
    return ppr


def kinds_of(para):
    return ', '.join('%s x%d' % kv for kv in sorted(para.inline.items())) or 'none'


def refuse_inline(b, what):
    """Deleted text is written from the baseline's characters; a picture, a field or a footnote reference among
    them cannot be written (and would need parts of the baseline file). Such a paragraph is refused."""
    assert not b.inline, ('a %s baseline paragraph has elements that deleted text cannot carry (%s): %r'
                          % (what, kinds_of(b), b.text.replace(INLINE_CHAR, '')[:60]))


def make_deleted_para(bp, b, styles, numids, stats):
    """New paragraph holding baseline text as deletions; its mark is set by the caller."""
    refuse_inline(b, 'deleted')
    for k in sorted(COMMENT_TAGS & set(b.special)):
        stats['dropped_baseline_special'][k] += b.special[k]
    p = etree.Element(q('p'))
    ppr = deleted_ppr(bp, styles, numids, stats)
    if ppr is not None and len(ppr):
        p.append(ppr)
    segs = segments([('del', 0, len(b.text))], b, b)
    for e in build_children(segs, styles):
        p.append(e)
    stats['exp_del'].append(''.join(t for k, t, *_ in segs))
    stats['marked'].append((b.text, '', marked_string(segs)))
    return p


def align_paras(bt_all, ct_all):
    """Monotone alignment of non-empty paragraphs: [('eq'|'mod', bi, ci) | ('del', bi, None) | ('ins', None, ci)].

    Empty paragraphs are left out (they carry no text and would create false anchors)."""
    bmap = [i for i, t in enumerate(bt_all) if t]
    cmap = [j for j, t in enumerate(ct_all) if t]
    bt, ct = [bt_all[i] for i in bmap], [ct_all[j] for j in cmap]
    out = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, bt, ct, autojunk=False).get_opcodes():
        if tag == 'equal':
            out += [('eq', i1 + k, j1 + k) for k in range(i2 - i1)]
        elif tag == 'delete':
            out += [('del', i, None) for i in range(i1, i2)]
        elif tag == 'insert':
            out += [('ins', None, j) for j in range(j1, j2)]
        else:
            out += dp_align(bt, ct, i1, i2, j1, j2)
    out = rescue_moved(out, bt, ct)
    return [(t, None if i is None else bmap[i], None if j is None else cmap[j]) for t, i, j in out]


def rescue_moved(out, bt, ct):
    """Re-pair similar paragraphs that the exact-match alignment left in different blocks.

    Cause of the missed pairs: align_paras anchors on exact-equal paragraphs, and dp_align only pairs paragraphs
    inside one block between two anchors. When an anchor was moved (a reference list re-sorted, one entry left at the
    end), a baseline paragraph and its edited clean counterpart end up in different blocks (one as delete-only, the
    other as insert-only) and are never compared, so both are struck / inserted whole.

    Unmatched baseline paragraphs and unmatched clean paragraphs of different blocks are paired when
    word_ratio >= RESCUE_MIN_RATIO. A max-weight order-preserving chain over the exact anchors (EQ_ANCHOR_WEIGHT each)
    and the candidate pairs (their ratio) decides which crossing anchors give way, so a single moved paragraph never
    displaces an exact anchor. The pairs kept are tagged 'modhi'; the gaps between anchors are re-aligned with dp_align;
    a demoted anchor is emitted as 'mdel' + 'mins' (a moved paragraph: struck where it was, inserted where it went).
    Returns out unchanged when no candidate pair survives."""
    gap, gap_of = 0, {}
    for k, (t, i, j) in enumerate(out):
        if t == 'eq':
            gap += 1
        gap_of[k] = gap
    dels = [(k, i) for k, (t, i, j) in enumerate(out) if t == 'del']
    inss = [(k, j) for k, (t, i, j) in enumerate(out) if t == 'ins']
    cands = []
    for kd, i in dels:
        for ki, j in inss:
            if gap_of[kd] != gap_of[ki]:
                r = word_ratio(bt[i], ct[j])
                if r >= RESCUE_MIN_RATIO:
                    cands.append((i, j, r, 'modhi'))
    if not cands:
        return out
    nodes = sorted([(i, j, EQ_ANCHOR_WEIGHT, 'eq') for t, i, j in out if t == 'eq'] + cands)
    best, prev = [], []                                  # weighted longest chain, increasing in both coordinates
    for x, (i, j, w, _) in enumerate(nodes):
        bx, px = w, -1
        for y in range(x):
            if nodes[y][0] < i and nodes[y][1] < j and best[y] + w > bx:
                bx, px = best[y] + w, y
        best.append(bx)
        prev.append(px)
    x = max(range(len(nodes)), key=lambda n: best[n])
    chain = []
    while x >= 0:
        chain.append(nodes[x])
        x = prev[x]
    chain.reverse()
    if not any(n[3] == 'modhi' for n in chain):
        return out
    kept = {(i, j) for i, j, _, _ in chain}
    dem_b = sorted(i for t, i, j in out if t == 'eq' and (i, j) not in kept)   # exact anchors that gave way
    dem_c = sorted(j for t, i, j in out if t == 'eq' and (i, j) not in kept)
    skip_b, skip_c = set(dem_b), set(dem_c)
    res, pi, pj = [], 0, 0
    for i, j, _, tag in chain + [(len(bt), len(ct), 0, None)]:
        bidx = [x for x in range(pi, i) if x not in skip_b]
        cidx = [y for y in range(pj, j) if y not in skip_c]
        if bidx and cidx:
            res += [(t, None if a is None else bidx[a], None if b is None else cidx[b])
                    for t, a, b in dp_align([bt[x] for x in bidx], [ct[y] for y in cidx], 0, len(bidx), 0, len(cidx))]
        else:
            res += [('del', x, None) for x in bidx] + [('ins', None, y) for y in cidx]
        if tag:
            res.append((tag, i, j))
        pi, pj = i + 1, j + 1
    # a demoted anchor becomes a plain delete (baseline place) and a plain insert (clean place), never paired again
    for x in dem_b:
        pos = next((n for n, e in enumerate(res) if e[1] is not None and e[1] > x), len(res))
        res.insert(pos, ('mdel', x, None))
    for y in dem_c:
        pos = next((n for n, e in enumerate(res) if e[2] is not None and e[2] > y), len(res))
        res.insert(pos, ('mins', None, y))
    return res


def dp_align(bt, ct, i1, i2, j1, j2):
    """Order-preserving max-total-ratio pairing inside a replace block (pairs below the threshold are dropped)."""
    n, m = i2 - i1, j2 - j1
    ratio = {}
    for a in range(n):
        for b in range(m):
            r = word_ratio(bt[i1 + a], ct[j1 + b]) if bt[i1 + a] and ct[j1 + b] else 0.0
            if r >= PARA_MATCH_MIN_RATIO:
                ratio[a, b] = r
    best = [[0.0] * (m + 1) for _ in range(n + 1)]
    for a in range(1, n + 1):
        for b in range(1, m + 1):
            v = max(best[a - 1][b], best[a][b - 1])
            if (a - 1, b - 1) in ratio:
                v = max(v, best[a - 1][b - 1] + ratio[a - 1, b - 1])
            best[a][b] = v
    out, a, b = [], n, m
    while a > 0 or b > 0:
        if a > 0 and b > 0 and (a - 1, b - 1) in ratio and abs(best[a][b] - best[a - 1][b - 1] - ratio[a - 1, b - 1]) < 1e-12:
            out.append(('mod', i1 + a - 1, j1 + b - 1))
            a, b = a - 1, b - 1
        elif a > 0 and abs(best[a][b] - best[a - 1][b]) < 1e-12:
            out.append(('del', i1 + a - 1, None))
            a -= 1
        elif b > 0:
            out.append(('ins', None, j1 + b - 1))
            b -= 1
        else:
            out.append(('del', i1 + a - 1, None))
            a -= 1
    return out[::-1]


def split_sentences(text):
    """Sentence spans [(start, end)] covering text; end includes trailing whitespace.

    Split after . ! ? (optionally followed by closing brackets or quotes) and whitespace when the next character
    starts a sentence (upper case, digit, opening bracket). Abbreviations (et al., Fig., vs., e.g.), single-letter
    initials (E. coli) and decimals (no following space) never split."""
    spans, start = [], 0
    for m in re.finditer(r'[.!?][)\]"\u2019\']*(\s+)', text):
        e = m.end()
        if e >= len(text):
            continue
        nxt = text[e]
        if not (nxt.isupper() or nxt.isdigit() or nxt in '([\u201c"'):
            continue
        w = re.search(r'([A-Za-z][A-Za-z.]*)\.[)\]"\u2019\']*\s+$', text[max(0, m.start() - 30):e])
        word = w.group(1).lower().replace('.', '') if w else ''
        if m.group()[0] == '.' and (word in SENT_ABBREV or len(word) == 1):
            continue
        spans.append((start, e))
        start = e
    if start < len(text) or not spans:
        spans.append((start, len(text)))
    return spans


def sent_ratio(ta, tb, min_ratio):
    """Word-token ratio of two token lists; 0.0 below min_ratio (exact ratio otherwise)."""
    if not ta or not tb:
        return 0.0
    sm = difflib.SequenceMatcher(None, ta, tb, autojunk=False)
    if sm.real_quick_ratio() < min_ratio or sm.quick_ratio() < min_ratio:
        return 0.0
    r = sm.ratio()
    return r if r >= min_ratio else 0.0


def content_dice(a, b):
    """Dice coefficient of the content-word multisets (lower-case, stop words dropped) of two texts."""
    ca = collections.Counter(w for w in re.findall(r'[a-z0-9]+', a.lower()) if w not in STOPWORDS and len(w) > 1)
    cb = collections.Counter(w for w in re.findall(r'[a-z0-9]+', b.lower()) if w not in STOPWORDS and len(w) > 1)
    tot = sum(ca.values()) + sum(cb.values())
    return 2 * sum((ca & cb).values()) / tot if tot else 0.0


def sentence_dp(bs, cs, min_ratio):
    """Order-preserving max-total-ratio sentence pairing: [(bi, ci)] over sentence records with token lists at [3]."""
    n, m = len(bs), len(cs)
    ratio = {}
    for a in range(n):
        for b in range(m):
            r = sent_ratio(bs[a][3], cs[b][3], min_ratio)
            if r:
                ratio[a, b] = r
    best = [[0.0] * (m + 1) for _ in range(n + 1)]
    for a in range(1, n + 1):
        for b in range(1, m + 1):
            v = max(best[a - 1][b], best[a][b - 1])
            if (a - 1, b - 1) in ratio:
                v = max(v, best[a - 1][b - 1] + ratio[a - 1, b - 1])
            best[a][b] = v
    out, a, b = [], n, m
    while a > 0 and b > 0:
        if (a - 1, b - 1) in ratio and abs(best[a][b] - best[a - 1][b - 1] - ratio[a - 1, b - 1]) < 1e-12:
            out.append((a - 1, b - 1))
            a, b = a - 1, b - 1
        elif abs(best[a][b] - best[a - 1][b]) < 1e-12:
            a -= 1
        else:
            b -= 1
    return out[::-1]


def pair_segs(bp, bspan, cp, cspan):
    """Word-level segments for one matched sentence pair (trailing whitespace handled separately)."""
    bs, be = bspan
    cs, ce = cspan
    bcore, ccore = bp.text[bs:be].rstrip(), cp.text[cs:ce].rstrip()
    ops = [(k, s0 + (bs if k == 'del' else cs), e0 + (bs if k == 'del' else cs)) for k, s0, e0 in diff_ops(bcore, ccore)]
    btail, ctail = (bs + len(bcore), be), (cs + len(ccore), ce)
    bt_, ct_ = bp.text[btail[0]:btail[1]], cp.text[ctail[0]:ctail[1]]
    if bt_ == ct_:
        if ct_:
            ops.append(('eq', ctail[0], ctail[1]))
    else:
        if bt_:
            ops.append(('del', btail[0], btail[1]))
        if ct_:
            ops.append(('ins', ctail[0], ctail[1]))
    return segments(ops, bp, cp)


def join_segs(parts):
    """Concatenate segment lists, merging neighbours with the same kind, format and hyperlink."""
    out = []
    for seg in parts:
        if out and out[-1][0] == seg[0] and out[-1][2] == seg[2] and out[-1][4] is seg[4]:
            out[-1][1] += seg[1]
        else:
            out.append(list(seg))
    return out


def sentence_realign(ali, B, C, min_ratio):
    """Re-pair every replace block (run of non-eq entries) at sentence level.

    Returns a new alignment whose entries are (tag, bi, ci, extra): 'eq', 'del' (baseline paragraph with no matched
    sentence, kept as a deleted paragraph), 'ins' (clean paragraph without baseline text) and 'seg' (clean paragraph
    with prebuilt segments and the baseline paragraphs it draws on in extra)."""
    out, k = [], 0
    while k < len(ali):
        tag = ali[k][0]
        if tag == 'eq' or (tag == 'mod' and B[ali[k][1]].text == C[ali[k][2]].text):
            out.append((tag, ali[k][1], ali[k][2], None))
            k += 1
            continue
        if tag in ('mdel', 'mins'):               # moved paragraph: stays a plain delete / insert, never paired
            out.append(('del' if tag == 'mdel' else 'ins', ali[k][1], ali[k][2], None))
            k += 1
            continue
        blk = []
        if tag == 'modhi':                        # rescued high-confidence pair: aligned on its own, never mixed with neighbours
            blk.append(('mod', ali[k][1], ali[k][2]))
            k += 1
        else:
            while k < len(ali) and not (ali[k][0] in ('eq', 'modhi', 'mdel', 'mins') or (ali[k][0] == 'mod' and B[ali[k][1]].text == C[ali[k][2]].text)):
                blk.append(ali[k])
                k += 1
        bparas = [bi for t, bi, ci in blk if bi is not None]
        cparas = [ci for t, bi, ci in blk if ci is not None]
        bsent = [(bi, sp, i, [t[0] for t in toks(B[bi].text[sp[0]:sp[1]])]) for bi in bparas for i, sp in enumerate(split_sentences(B[bi].text))]
        csent = [(ci, sp, i, [t[0] for t in toks(C[ci].text[sp[0]:sp[1]])]) for ci in cparas for i, sp in enumerate(split_sentences(C[ci].text))]
        pairs = sentence_dp(bsent, csent, min_ratio) if bsent and csent else []
        c_of_b = dict(pairs)                      # baseline sentence index -> clean sentence index
        matched_b = {}                            # baseline paragraph -> sorted matched sentence indices
        for a, _ in pairs:
            matched_b.setdefault(bsent[a][0], []).append(a)
        before, after = collections.defaultdict(list), collections.defaultdict(list)
        absorb_at = {}
        for a, (bi, sp, i, _) in enumerate(bsent):
            if a in c_of_b or bi not in matched_b:
                continue
            prev = [x for x in matched_b[bi] if x < a]
            if prev:
                after[c_of_b[prev[-1]]].append(a)
            else:
                before[c_of_b[[x for x in matched_b[bi] if x > a][0]]].append(a)
        # absorb wholly unmatched baseline paragraphs that sit next to a matched one and share content words with
        # the clean paragraph that received that neighbour: their text is struck inside that clean paragraph
        absorbed = set()
        first_a = {bi: min(xs) for bi, xs in matched_b.items()}
        last_a = {bi: max(xs) for bi, xs in matched_b.items()}
        changed = True
        while changed:
            changed = False
            for pos, bi in enumerate(bparas):
                if bi in matched_b or bi in absorbed:
                    continue
                cands = []
                if pos + 1 < len(bparas) and (bparas[pos + 1] in matched_b or bparas[pos + 1] in absorbed):
                    cands.append(('before', bparas[pos + 1]))
                if pos > 0 and (bparas[pos - 1] in matched_b or bparas[pos - 1] in absorbed):
                    cands.append(('after', bparas[pos - 1]))
                best = None
                for side, nb in cands:
                    cidx = c_of_b.get(first_a[nb] if side == 'before' else last_a[nb]) if nb in matched_b else absorb_at[nb][1]
                    if cidx is None:
                        continue
                    d = content_dice(B[bi].text, C[csent[cidx][0]].text)
                    if d >= ABSORB_MIN_DICE and (best is None or d > best[0]):
                        best = (d, side, cidx)
                if best:
                    absorbed.add(bi)
                    absorb_at[bi] = (best[1], best[2])
                    idxs = [a for a, x in enumerate(bsent) if x[0] == bi]
                    (before if best[1] == 'before' else after)[best[2]].extend(idxs)
                    changed = True
        # sandwich rule: a run of unassigned baseline paragraphs between two paragraphs that landed in the same
        # clean paragraph must be struck inside it too, otherwise baseline order could not be restored
        def home(bi, last):
            if bi in matched_b:
                return csent[c_of_b[(last_a if last else first_a)[bi]]][0], c_of_b[(last_a if last else first_a)[bi]]
            if bi in absorbed:
                return csent[absorb_at[bi][1]][0], absorb_at[bi][1]
            return None
        pos = 0
        while pos < len(bparas):
            if bparas[pos] in matched_b or bparas[pos] in absorbed:
                pos += 1
                continue
            end = pos
            while end < len(bparas) and bparas[end] not in matched_b and bparas[end] not in absorbed:
                end += 1
            if pos > 0 and end < len(bparas):
                left, right = home(bparas[pos - 1], True), home(bparas[end], False)
                if left and right and left[0] == right[0]:
                    for bi in bparas[pos:end]:
                        absorbed.add(bi)
                        absorb_at[bi] = ('after', left[1])
                        after[left[1]].extend(a for a, x in enumerate(bsent) if x[0] == bi)
            pos = end
        for slot in list(before.values()) + list(after.values()):
            slot.sort()
        used = collections.defaultdict(set)
        segs_of = {}
        for ci in cparas:
            parts = []
            state = {'last_b': None, 'short_del': False}

            def put(bi, segs, whole_del=False):
                """Append segments; bi is the baseline paragraph they come from (None = inserted clean text).
                Where the baseline paragraph changes, a separator says so (SEP_CHAR): rebuild_para splits the
                paragraph there, so a baseline paragraph boundary (and a struck heading's own line) stays one."""
                if bi is None:
                    if state['short_del']:
                        parts.append(['del', SEP_CHAR, ('sep', state['last_b']), None, None])
                    state['short_del'] = False
                else:
                    if state['last_b'] is not None and bi != state['last_b']:
                        parts.append(['del', SEP_CHAR, ('sep', state['last_b']), None, None])
                    state['last_b'] = bi
                    state['short_del'] = whole_del and len(B[bi].text.strip()) < SHORT_PARA_CHARS
                parts.extend(segs)
            for cidx, (cpi, csp, _, _) in enumerate(csent):
                if cpi != ci:
                    continue
                def dele(a):
                    bi, sp, _, _ = bsent[a]
                    used[ci].add(bi)
                    put(bi, segments([('del', sp[0], sp[1])], B[bi], C[ci]), whole_del=bi not in matched_b)
                for a in before[cidx]:
                    dele(a)
                back = [a for a, c2 in pairs if c2 == cidx]
                if back:
                    a = back[0]
                    used[ci].add(bsent[a][0])
                    put(bsent[a][0], pair_segs(B[bsent[a][0]], bsent[a][1], C[ci], csp))
                else:
                    put(None, segments([('ins', csp[0], csp[1])], B[bparas[0]] if bparas else C[ci], C[ci]))
                for a in after[cidx]:
                    dele(a)
            segs_of[ci] = join_segs(parts)
        start = len(out)
        for t, bi, ci in blk:
            if bi is not None and bi not in matched_b and bi not in absorbed:
                out.append(('del', bi, None, None))
            if ci is not None:
                if used[ci]:
                    out.append(('seg', None, ci, (segs_of[ci], sorted(used[ci]))))
                else:
                    out.append(('ins', None, ci, None))
        # keep deleted paragraphs in baseline order relative to clean paragraphs that carry baseline text
        moved = True
        while moved:                              # until nothing moves: a deleted paragraph can wait behind another one
            moved = False
            for p in range(start, len(out)):
                if out[p][0] != 'del':
                    continue
                bi, q = out[p][1], p
                while q > start and out[q - 1][0] == 'seg' and max(out[q - 1][3][1]) > bi and min(out[q - 1][3][1]) > bi:
                    out[q - 1], out[q] = out[q], out[q - 1]
                    q -= 1
                while q + 1 < len(out) and out[q + 1][0] == 'seg' and max(out[q + 1][3][1]) < bi:
                    out[q + 1], out[q] = out[q], out[q + 1]
                    q += 1
                moved = moved or q != p
    return out


def apply_container(bps, cps, styles, numids, stats, tables=None):
    """Diff one list of baseline paragraphs against one list of clean paragraphs, editing the clean tree in place.

    tables (the document body only): {'before': tables that stand before each baseline paragraph, 'match': {id of a
    clean table: its baseline table or None}, 'new': ids of the clean tables that will be marked as inserted}; a
    deleted paragraph is put on the side of a table where it stood.
    Returns, for every baseline paragraph, the output paragraph that ends it (see plan_marks)."""
    B = [read_para(p) for p in bps]
    C = [read_para(p, skip_rpr=lambda r: False) for p in cps]
    ali = align_paras([x.text for x in B], [x.text for x in C])
    ali = sentence_realign(ali, B, C, SENT_MATCH_MIN_RATIO)
    seq = list(cps)                                   # the output paragraphs of this container, in document order
    head = {}                                         # id of a clean paragraph -> the first piece split off in front of it

    def at(p):
        return next(i for i, e in enumerate(seq) if e is p)

    for tag, bi, ci, extra in ali:
        if tag == 'eq':
            # the clean paragraph is kept as it is, so reject-all shows its elements: they must be the baseline's
            assert B[bi].inline == C[ci].inline, (
                'a paragraph with the same text holds other elements that are not text in the baseline (%s) than in '
                'the clean document (%s), which cannot be shown: %r'
                % (kinds_of(B[bi]), kinds_of(C[ci]), C[ci].text.replace(INLINE_CHAR, '')[:60]))
            stats['n_eq'] += 1
        elif tag == 'seg':
            assert not C[ci].special.get('drawing'), 'text differs in clean paragraph with drawing: %r' % C[ci].text[:60]
            segs, used = extra
            for u in used:
                refuse_inline(B[u], 'changed')
            new = rebuild_para(cps[ci], ' \u00b6 '.join(B[u].text for u in used), C[ci], styles, stats, segs,
                               lambda b: deleted_ppr(bps[b], styles, numids, stats))
            seq[at(cps[ci]):at(cps[ci])] = new
            if new:
                head[id(cps[ci])] = new[0]
            stats['n_mod'] += 1
        elif tag == 'ins':
            stats['n_ins'] += 1
            assert not C[ci].loose, (
                'a new paragraph has elements outside a run that cannot be marked as inserted (%s): %r'
                % (', '.join('%s x%d' % kv for kv in sorted(C[ci].loose.items())), C[ci].text.replace(INLINE_CHAR, '')[:60]))
            if C[ci].text:
                colorize_in_place(cps[ci])
                stats['exp_ins'].append(C[ci].text)
                stats['marked'].append(('', C[ci].text, '{+%s+}' % C[ci].text))

    def empties_before(i):
        """How many empty baseline paragraphs stand right before baseline paragraph i."""
        n = 0
        while i - n - 1 >= 0 and not B[i - n - 1].text:
            n += 1
        return n

    blank = {id(x) for x, c in zip(cps, C) if not c.text}   # the empty clean paragraphs

    def is_table(e, later):
        """Is e a clean table that this deleted paragraph must stay behind (later=False) or ahead of (later=True)?"""
        if e is None or e.tag != q('tbl') or tables is None or tables['match'].get(id(e)) is None:
            return False
        return (tables['match'][id(e)] >= tables['before'][bi]) == later

    last = None
    for idx, (tag, bi, ci, extra) in enumerate(ali):
        if tag != 'del':
            last = cps[ci] if ci is not None else last
            continue
        stats['n_del'] += 1
        newp = make_deleted_para(bps[bi], B[bi], styles, numids, stats)
        nxt = next((cps[a[2]] for a in ali[idx + 1:] if a[0] != 'del'), None)
        if nxt is not None:
            first = head.get(id(nxt), nxt)
            k = at(first)
            run = 0                                   # empty clean paragraphs right before it ...
            while k - run - 1 >= 0 and id(seq[k - run - 1]) in blank:
                run += 1
            first = seq[k - (run - min(run, empties_before(bi)))]   # ... of which the baseline had this many before it
            anchor = first
            while is_table(anchor.getprevious(), later=True):   # the table came after this paragraph: stay ahead of it
                anchor = anchor.getprevious()
            anchor.addprevious(newp)
            seq.insert(at(first), newp)
        elif last is None:                            # the clean paragraphs of this part are all empty
            cps[0].addprevious(newp)
            seq.insert(at(cps[0]), newp)
            last = newp
        else:
            k = at(last)
            for _ in range(empties_before(bi)):       # behind as many empty clean paragraphs as the baseline had before it
                if k + 1 < len(seq) and id(seq[k + 1]) in blank:
                    k += 1
            anchor = seq[k]
            while is_table(anchor.getnext(), later=False):      # the table came before this paragraph: go behind it
                anchor = anchor.getnext()
            anchor.addnext(newp)
            seq.insert(k + 1, newp)
            last = newp
        stats['deleted_paras'].append(B[bi].text[:100])
    return plan_marks(seq, cps, B, C, stats, tables['new'] if tables else ())


# ---------------------------------------------------------------- document level
def in_box(p):
    """True for paragraphs inside drawings or text boxes."""
    return any(a.tag in (q('drawing'), q('pict'), q('txbxContent')) for a in p.iterancestors())


def top_tables(root):
    """Top-level tables in document order (nested tables are not supported)."""
    tbls = [t for t in root.iter(q('tbl')) if not any(a.tag == q('tbl') for a in t.iterancestors())]
    assert all(t.find('.//' + q('tbl')) is None for t in tbls), 'nested tables are not supported'
    return tbls


def body_paras(root):
    """Paragraphs outside tables and drawings, document order (sdt content included)."""
    return [p for p in root.iter(q('p')) if not in_box(p) and not any(a.tag == q('tbl') for a in p.iterancestors())]


def cell_layout(tc):
    """[columns a cell spans, its vertical merge (None | 'restart' | 'continue'), its old-style horizontal merge]."""
    pr = tc.find(q('tcPr'))
    span = pr.find(q('gridSpan')) if pr is not None else None
    vm = pr.find(q('vMerge')) if pr is not None else None
    hm = pr.find(q('hMerge')) if pr is not None else None
    return [int(span.get(q('val')) or 1) if span is not None else 1,
            None if vm is None else vm.get(q('val')) or 'continue',
            None if hm is None else hm.get(q('val')) or 'continue']


def row_skips(tr):
    """[grid columns a row leaves out before its cells, and after them]."""
    def n(tag):
        e = tr.find(q('trPr') + '/' + q(tag))
        return int(e.get(q('val')) or 0) if e is not None else 0
    return [n('gridBefore'), n('gridAfter')]


def table_layout(t):
    """[grid columns, rows of [row_skips, cell_layout of every cell]]: two tables with the same layout can be
    compared cell by cell."""
    return [len(t.findall(q('tblGrid') + '/' + q('gridCol'))),
            [[row_skips(tr)] + [cell_layout(tc) for tc in tr.findall(q('tc'))] for tr in t.findall(q('tr'))]]


def table_dims(t):
    """Cell counts per row."""
    return [len(tr.findall(q('tc'))) for tr in t.findall(q('tr'))]


def cell_paras(tc):
    """Direct paragraph children of a cell."""
    ps = tc.findall(q('p'))
    assert len(list(tc.iter(q('p')))) == len(ps), 'paragraphs nested inside a cell child are not supported'
    return ps


def mark_table(t, kind):
    """Mark a whole table as inserted or deleted: its rows in track_changes mode (w:trPr/w:ins|w:del); in colour
    mode every paragraph mark in it (a row whose paragraph marks are all inserted, or all deleted, is an inserted
    or deleted row: that is how the check reads it)."""
    if MODE == 'track_changes':
        for tr in t.iter(q('tr')):
            set_row_mark(tr, kind)
    else:
        for p in t.iter(q('p')):
            if not in_box(p):
                set_para_mark(p, kind)


def deleted_table(bt, styles, numids, stats):
    """Baseline table copy, marked as deleted: every paragraph as deleted text (each cell keeps a paragraph)."""
    t = copy.deepcopy(bt)
    for tc in t.iter(q('tc')):
        for p in tc.findall(q('p')):
            b = read_para(p)
            newp = make_deleted_para(p, b, styles, numids, stats)
            p.addprevious(newp)
            tc.remove(p)
        if tc.find(q('p')) is None:
            tc.append(etree.Element(q('p')))
    mark_table(t, 'del')
    return t


def match_tables(bts, cts):
    """Pair baseline and clean tables in order: [(baseline index | None, clean index | None)].
    Tables with the same text anchor the pairing; between two anchors tables are paired in order."""
    text = lambda t: ''.join(t.itertext())
    out = []
    sm = difflib.SequenceMatcher(None, [text(t) for t in bts], [text(t) for t in cts], autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        n = min(i2 - i1, j2 - j1)
        out += [(i1 + k, j1 + k) for k in range(n)]
        out += [(i, None) for i in range(i1 + n, i2)] + [(None, j) for j in range(j1 + n, j2)]
    return out


def table_places(root):
    """(body paragraphs before each top-level table, top-level tables before each body paragraph)."""
    held = top_tables(root)                           # held: an id is only stable while its element object lives
    tops = {id(t) for t in held}
    paras_before, tables_before = [], []
    for e in root.iter(q('p'), q('tbl')):
        if e.tag == q('tbl'):
            if id(e) in tops:
                paras_before.append(len(tables_before))
        elif not in_box(e) and not any(a.tag == q('tbl') for a in e.iterancestors()):
            tables_before.append(len(paras_before))
    return paras_before, tables_before


def build_document(base_xml, clean_xml, styles, numids):
    """Return (new document.xml bytes, stats)."""
    broot, croot = etree.fromstring(base_xml), etree.fromstring(clean_xml)
    stats = {'n_eq': 0, 'n_mod': 0, 'n_ins': 0, 'n_del': 0, 'n_del_empty': 0, 'marked': [], 'exp_ins': [],
             'exp_del': [], 'deleted_paras': [], 'dropped_pstyle': 0, 'dropped_numpr': 0,
             'dropped_baseline_special': collections.Counter(), 'tables': [],
             'para_marks': collections.Counter()}
    bts, cts = top_tables(broot), top_tables(croot)
    pairs = match_tables(bts, cts)
    paras_before, tables_before = table_places(broot)
    at_base = {id(cts[ck]): bk for bk, ck in pairs if ck is not None}   # output table -> its baseline table (or None)
    bps, cps = body_paras(broot), body_paras(croot)
    assert bps and cps, 'a document without a paragraph outside its tables'
    same_dims = lambda bk, ck: table_layout(bts[bk]) == table_layout(cts[ck])   # cells, spans and merges alike
    new = {id(cts[ck]) for bk, ck in pairs if ck is not None and bk is None}   # no deleted table will stand next to these
    owners = apply_container(bps, cps, styles, numids, stats, {'before': tables_before, 'match': at_base, 'new': new})
    assert len(owners) == len(bps), 'not every baseline paragraph has its place in the marked text'
    out_paras = body_paras(croot)
    first = out_paras[0]
    place = {id(x): i for i, x in enumerate(out_paras)}
    clean_before, _ = table_places(croot)             # output body paragraphs before each clean table

    def in_place(bk, ck):
        """Does the clean table stand between the same two baseline paragraphs as its baseline table?"""
        n, here = paras_before[bk], clean_before[ck]
        return (n == 0 or place[id(owners[n - 1])] < here) and (n == len(bps) or place[id(owners[n])] >= here)
    held = []                                         # the deleted table copies: at_base is keyed by id

    def known(e, bk, later):
        """Is e a table of the output whose baseline table stands after (later) or before the baseline table bk?"""
        if e is None or e.tag != q('tbl') or at_base.get(id(e)) is None:
            return False
        return (at_base[id(e)] > bk) == later

    for bk, ck in pairs:
        if ck is not None and bk is not None and same_dims(bk, ck) and in_place(bk, ck):
            for btc, ctc in zip(bts[bk].iter(q('tc')), cts[ck].iter(q('tc'))):
                apply_container(cell_paras(btc), cell_paras(ctc), styles, numids, stats)
            stats['tables'].append('table %d: cell-by-cell (%s)' % (ck + 1, table_dims(cts[ck])))
            continue
        old = deleted_table(bts[bk], styles, numids, stats) if bk is not None else None
        if ck is not None:
            ct = cts[ck]
            for p in ct.iter(q('p')):
                text = read_para(p, lambda r: False).text
                if text:
                    colorize_in_place(p)
                    stats['exp_ins'].append(text)
                    stats['marked'].append(('', text, '{+%s+}' % text))
            mark_table(ct, 'ins')
            stats['tables'].append('table %d: whole-table insert (baseline dims %s, clean dims %s)'
                                   % (ck + 1, table_dims(bts[bk]) if bk is not None else None, table_dims(ct)))
        if old is None:
            continue
        if ck is not None and in_place(bk, ck):
            cts[ck].addprevious(old)
            stats['tables'].append('table %d: baseline table inserted as deleted' % (ck + 1))
        elif paras_before[bk]:                        # a table that is gone or moved: behind the paragraph it followed
            anchor = owners[paras_before[bk] - 1]
            while known(anchor.getnext(), bk, later=False):
                anchor = anchor.getnext()
            anchor.addnext(old)
            stats['tables'].append('baseline table %d: deleted' % (bk + 1))
        else:                                         # ... or, when it opened the document, ahead of the first paragraph
            anchor = first
            while known(anchor.getprevious(), bk, later=True):
                anchor = anchor.getprevious()
            anchor.addprevious(old)
            stats['tables'].append('baseline table %d: deleted' % (bk + 1))
        held.append(old)
        at_base[id(old)] = bk
    stats['warnings'] = []
    if MODE == 'color':
        stats['warnings'] = ['bookmark "%s" now also holds %d character(s) of deleted text (struck); a cross-reference '
                             'to it shows that text too' % (name, n) for name, n in bookmarks_holding_deletions(croot).items()]
    return etree.tostring(croot, xml_declaration=True, encoding='UTF-8', standalone=True), stats


def style_ids(z):
    """(styleIds, numIds) declared in a docx zip; None when the part is absent."""
    def ids(part, tag, attr):
        if part not in z.namelist():
            return None
        return {e.get(q(attr)) for e in etree.fromstring(z.read(part)).iter(q(tag))}
    return ids('word/styles.xml', 'style', 'styleId'), ids('word/numbering.xml', 'num', 'numId')


class Refused(Exception):
    """The request cannot be carried out; the message says why."""


def same_file(a, b):
    """True when two paths name one file: equal real paths, or one existing file reached two ways."""
    ra, rb = os.path.realpath(a), os.path.realpath(b)
    if ra == rb or os.path.normcase(ra) == os.path.normcase(rb):
        return True
    return os.path.exists(a) and os.path.exists(b) and os.path.samefile(a, b)


def place(tmp, out, overwrite=False):
    """Move the finished temporary file to `out`. Without `overwrite` an existing `out` is never replaced:
    the name is created exclusively (a hard link, or an exclusive copy where links are not supported), so a
    file that appeared after the first check is refused here as well."""
    if overwrite:
        os.replace(tmp, out)
        return
    try:
        os.link(tmp, out)
    except FileExistsError:
        raise Refused('the output already exists and is not overwritten: %s' % out)
    except OSError:  # no hard links on this file system
        try:
            dst = open(out, 'xb')
        except FileExistsError:
            raise Refused('the output already exists and is not overwritten: %s' % out)
        try:
            with dst, open(tmp, 'rb') as src:
                shutil.copyfileobj(src, dst)
        except OSError:
            os.unlink(out)  # this call created the file: a half-written copy is not left behind
            raise
    os.unlink(tmp)


def write_docx(clean_path, out_path, doc_xml):
    """Copy the clean zip with only word/document.xml replaced (ZipInfo, order and comment preserved)."""
    with zipfile.ZipFile(clean_path) as zin, zipfile.ZipFile(out_path, 'w') as zout:
        zout.comment = zin.comment
        for info in zin.infolist():
            data = doc_xml if info.filename == 'word/document.xml' else zin.read(info.filename)
            ni = copy.copy(info)
            zout.writestr(ni, data, compress_type=info.compress_type)


def colour_conflicts(clean_xml):
    """(struck runs, runs in the insert colour, paragraph marks formatted like a mark of this tool) in the clean
    document: colour mode cannot mark such a file, the marks could not be told apart."""
    struck = coloured = 0
    root = etree.fromstring(clean_xml)
    marks = sum(1 for p in root.iter(q('p')) if para_mark(p))
    for r in root.iter(q('r')):
        if not any(ln(ch) in ('t', 'tab', 'br') for ch in r if isinstance(ch.tag, str)):
            continue
        rpr = r.find(q('rPr'))
        if is_struck(rpr):
            struck += 1
        elif rpr_color(rpr) == INSERT_RGB:
            coloured += 1
    return struck, coloured, marks


# ---------------------------------------------------------------- verification and metrics
def para_list(root):
    """All paragraphs (body, sdt, table cells) in document order, excluding drawing text boxes."""
    return [p for p in root.iter(q('p')) if not in_box(p)]


def is_deleted_rpr(rpr):
    """Deletion style of the marked file: strike plus DELETE_RGB."""
    return is_struck(rpr) and rpr_color(rpr) == DELETE_RGB


def is_inserted_rpr(rpr):
    """Insertion style of the marked file."""
    return rpr_color(rpr) == INSERT_RGB and not is_struck(rpr)


def run_kind(r, insert_rgbs=None):
    """'ins' | 'del' | 'eq' of a w:r in either mode (track_changes: by the enclosing mark; colour: by rPr)."""
    if MODE == 'track_changes':
        par = r.getparent()
        t = ln(par) if par is not None and isinstance(par.tag, str) else ''
        return t if t in ('ins', 'del') else 'eq'
    rpr = r.find(q('rPr'))
    if is_struck(rpr):
        return 'del'
    return 'ins' if rpr_color(rpr) in (insert_rgbs or (INSERT_RGB,)) else 'eq'


def read_accept(p):
    """Accept-all text of a marked paragraph (deletions dropped)."""
    return read_para(p, is_deleted_rpr) if MODE != 'track_changes' else read_para(p, lambda r: False)


def read_reject(p):
    """Reject-all text of a marked paragraph (insertions dropped, separators kept)."""
    return read_para(p, is_inserted_rpr) if MODE != 'track_changes' else read_para(p, lambda r: False, side='reject')


def read_only(p, kind):
    """Text of only the inserted ('ins') or deleted ('del') runs of a marked paragraph."""
    if MODE != 'track_changes':
        keep = is_inserted_rpr if kind == 'ins' else is_deleted_rpr
        return read_para(p, lambda r: not keep(r))
    return read_para(p, lambda r: False, side=kind)


def no_page_breaks(text):
    """Text without its page and column breaks (deleted ones are not written, so the reject side is compared without them)."""
    return ''.join(ch for ch in text if ch not in PAGE_BREAKS)


def view_root(marked_xml, side):
    """The marked document with every mark accepted ('accept') or rejected ('reject'), as a tree without marks.

    track_changes mode: resolve_revisions. Colour mode: the runs marked as deleted (accept) or inserted (reject)
    go, a row whose paragraph marks are all of that kind goes, and a paragraph whose mark is of that kind is
    joined with the next one."""
    if MODE == 'track_changes':
        return etree.fromstring(resolve_revisions(marked_xml, side)[0])
    root = etree.fromstring(marked_xml)
    gone = 'del' if side == 'accept' else 'ins'
    for tr in list(root.iter(q('tr'))):
        if {para_mark(x) for x in tr.iter(q('p'))} == {gone}:
            tr.getparent().remove(tr)
    for t in list(root.iter(q('tbl'))):
        if t.find(q('tr')) is None:
            t.getparent().remove(t)
    for r in list(root.iter(q('r'))):
        if run_kind(r) == gone:
            r.getparent().remove(r)
    for x in [x for x in root.iter(q('p')) if para_mark(x) == gone]:
        join_with_next(x)
    return root


def structure(root, baseline=False):
    """What the round-trip check compares of a document without marks:
    'flow'      the order of body paragraphs ('p') and top-level tables ('t'),
    'body'      every body paragraph, empty ones included, as [text with tabs, typed breaks (BREAK_CHARS) and one
                INLINE_CHAR per element that is not text, those elements counted by kind],
    'tables'    for every top-level table [grid columns, rows of cells as [columns spanned, vertical merge,
                old-style horizontal merge, the paragraphs of the cell], each row closed by its row_skips],
    'bookmarks' the text between the start and the end of every bookmark,
    'comments'  for every paragraph, also in tables, how many comment marks it holds (COMMENT_TAGS).
    baseline=True reads struck text as already deleted, as the diff does."""
    def read(x):
        para = read_para(x) if baseline else read_para(x, lambda r: False)
        return [para.text, sorted(para.inline.items())]
    held = top_tables(root)                           # held: an id is only stable while its element object lives
    tops = {id(t) for t in held}
    flow, body = [], []
    for e in root.iter(q('p'), q('tbl')):
        if e.tag == q('tbl'):
            if id(e) in tops:
                flow.append('t')
        elif not in_box(e) and not any(a.tag == q('tbl') for a in e.iterancestors()):
            flow.append('p')
            body.append(read(e))
    tables = [[len(t.findall(q('tblGrid') + '/' + q('gridCol'))),
               [[cell_layout(tc) + [[read(x) for x in tc.iter(q('p')) if not in_box(x)]] for tc in tr.findall(q('tc'))]
                + [row_skips(tr)] for tr in t.findall(q('tr'))]] for t in held]
    opened, marks = {}, {}
    for e in root.iter(q('bookmarkStart'), q('bookmarkEnd'), q('r')):
        if ln(e) == 'bookmarkStart':
            opened[e.get(q('id'))] = e.get(q('name')) or ''
            marks.setdefault(e.get(q('name')) or '', '')
        elif ln(e) == 'bookmarkEnd':
            opened.pop(e.get(q('id')), None)
        else:
            text = ''.join(t.text or '' for t in e.iter(q('t'), q('delText')))
            for name in opened.values():
                marks[name] += text
    comments = [sum(1 for e in x.iter() if isinstance(e.tag, str) and ln(e) in COMMENT_TAGS)
                for x in root.iter(q('p')) if not in_box(x)]
    return {'flow': flow, 'body': body, 'tables': tables, 'bookmarks': marks, 'comments': comments}


def first_difference(name, got, want):
    """[] when equal, else one line that says where two lists (or dicts) differ."""
    if got == want:
        return []
    if isinstance(got, dict):
        key = next(k for k in sorted(set(got) | set(want)) if got.get(k) != want.get(k))
        return ['%s: %r is %r where %r is expected' % (name, key, got.get(key), want.get(key))]
    if len(got) != len(want):
        return ['%s: %d where %d are expected' % (name, len(got), len(want))]
    i = next(i for i, (g, w) in enumerate(zip(got, want)) if g != w)
    return ['%s: #%d is %s where %s is expected' % (name, i + 1, repr(got[i])[:70], repr(want[i])[:70])]


def structure_failures(marked_xml, base_xml, clean_xml):
    """The round trip on structure: accepting every mark must give the structure of the clean document, rejecting
    every mark the structure of the baseline (see structure). Returns (accept failures, reject failures).

    Not compared on the reject side: where a table stands among the paragraphs, bookmarks and comment marks (the
    output carries the clean document's; a comment is no tracked change), and, in colour mode only, page and column breaks (a deleted one is not written)."""
    want_c = structure(etree.fromstring(clean_xml))
    want_b = structure(etree.fromstring(base_xml), baseline=True)
    acc, rej = structure(view_root(marked_xml, 'accept')), structure(view_root(marked_xml, 'reject'))
    a = []
    for key, label in (('flow', 'order of paragraphs and tables'), ('body', 'body paragraphs'),
                       ('tables', 'tables'), ('bookmarks', 'bookmarks'), ('comments', 'comment marks')):
        a += first_difference('accept-all, %s' % label, acc[key], want_c[key])
    if MODE == 'color':
        strip = lambda v: no_page_breaks(v) if isinstance(v, str) else [strip(x) for x in v] if isinstance(v, (list, tuple)) else v
        rej, want_b = {k: strip(rej[k]) for k in ('body', 'tables')}, {k: strip(want_b[k]) for k in ('body', 'tables')}
    r = first_difference('reject-all, body paragraphs', rej['body'], want_b['body'])
    r += first_difference('reject-all, tables', rej['tables'], want_b['tables'])
    return a, r


def words_of(texts):
    """The words of some texts as a multiset (which paragraph holds them does not matter)."""
    return collections.Counter(w for t in texts for w in t.replace(SEP_CHAR, ' ').split())


def nonempty(texts):
    """Drop empty entries."""
    return [t for t in texts if t]


def group_compare(name, got, want, max_group=12):
    """Order-preserving comparison that allows a baseline paragraph to be split or merged across clean paragraphs:
    consecutive groups of got and want (each at most max_group paragraphs) must have equal concatenated text."""
    fails, i, j = [], 0, 0
    while i < len(got) and j < len(want):
        if got[i] == want[j]:
            i, j = i + 1, j + 1
            continue
        for tot in range(2, 2 * max_group + 1):
            hit = next(((a, tot - a) for a in range(1, min(max_group, tot - 1) + 1) if 0 < tot - a <= max_group
                        and ''.join(got[i:i + a]) == ''.join(want[j:j + tot - a])), None)
            if hit:
                i, j = i + hit[0], j + hit[1]
                break
        else:
            fails.append('%s: no matching group at got #%d %r vs want #%d %r' % (name, i, got[i][:80], j, want[j][:80]))
            break
    if not fails and (i != len(got) or j != len(want)):
        fails.append('%s: trailing paragraphs unmatched (%d of %d, %d of %d)' % (name, i, len(got), j, len(want)))
    if ''.join(got) != ''.join(want):
        fails.append('%s: whole-document concatenation differs' % name)
    return fails


def boundary_offsets(units):
    """Cumulative character offsets at which one unit ends (the last one excluded)."""
    offs, tot = set(), 0
    for u in units[:-1]:
        tot += len(u)
        offs.add(tot)
    return offs


def boundary_check(orig_raw, old, context=30):
    """Every baseline paragraph boundary must still be a boundary in the marked file.

    orig_raw: reject-all text of each non-empty marked paragraph, struck separators (SEP_CHAR) kept in place; old: baseline
    paragraph texts. A boundary is visible when it is a paragraph edge or a separator in the marked file. The restoration
    check compares concatenations and cannot see a lost boundary, so this one runs on offsets of the concatenated text."""
    units = []
    for t in orig_raw:
        units += [u for u in t.split(SEP_CHAR) if u]
    want, got = boundary_offsets(old), boundary_offsets(units)
    joined = ''.join(old)
    if ''.join(units) != joined:
        return []                                   # restoration already fails; offsets would be meaningless
    lost = sorted(want - got)
    return ['paragraph boundaries: %d of %d baseline boundaries lost, first: %r | %r'
            % (len(lost), len(want), joined[max(0, lost[0] - context):lost[0]], joined[lost[0]:lost[0] + context])] if lost else []


def verify(marked_xml, base_xml, clean_xml):
    """Two-way restoration of the text paragraph by paragraph, and of the structure; returns counts and failures."""
    m, b, c = (etree.fromstring(x) for x in (marked_xml, base_xml, clean_xml))
    plain = lambda x: read_para(x, lambda r: False).text
    acc = nonempty([plain(p) for p in para_list(view_root(marked_xml, 'accept'))])
    orig_raw = nonempty([no_page_breaks(plain(p)) for p in para_list(view_root(marked_xml, 'reject'))])
    orig = orig_raw
    new = nonempty([read_para(p, lambda r: False).text for p in para_list(c)])
    old = nonempty([no_page_breaks(read_para(p).text) for p in para_list(b)])
    fails = []
    if len(acc) != len(new):
        fails.append('accept-all == clean: paragraph count %d vs %d' % (len(acc), len(new)))
    for i, (x, y) in enumerate(zip(acc, new)):
        if x != y:
            fails.append('accept-all == clean: paragraph #%d differs: %r vs %r' % (i, x[:80], y[:80]))
    fails += group_compare('reject-all == baseline', orig, old)
    fails += boundary_check(orig_raw, old)
    tc_no_p = sum(1 for tc in m.iter(q('tc')) if tc.find(q('p')) is None)
    ins_by_para = words_of(read_only(p, 'ins').text for p in para_list(m))
    del_by_para = words_of(read_only(p, 'del').text for p in para_list(m))
    s_acc, s_rej = structure_failures(marked_xml, base_xml, clean_xml)
    return {'paragraphs_compared': len(new) + len(old), 'failures': fails, 'structure_accept': s_acc,
            'structure_reject': s_rej, 'wrong': fails + s_acc + s_rej,
            'tc_no_p': tc_no_p, 'insert_runs': sum(1 for r in m.iter(q('r')) if run_kind(r) == 'ins'),
            'delete_runs': sum(1 for r in m.iter(q('r')) if run_kind(r) == 'del'),
            'ins_by_para': ins_by_para, 'del_by_para': del_by_para}


def metrics(xml):
    """Minimality metrics of a marked document.xml."""
    root = etree.fromstring(xml)
    struck = ins = 0
    over = []
    frags = []
    frags_paired = []  # fragments of paragraphs that also keep unchanged text
    for p in para_list(root):
        cur, cur_len, tot, st = None, 0, 0, 0
        n_before, n_eq = len(frags), 0
        for r in p.iter(q('r')):
            n = sum(len(t.text or '') for t in r.iter(q('t'), q('delText')))
            if not n:
                continue
            kind = run_kind(r)
            tot += n
            if kind == 'del':
                struck += n
                st += n
            elif kind == 'ins':
                ins += n
            else:
                n_eq += n
            if kind == cur:
                cur_len += n
            else:
                if cur in ('del', 'ins'):
                    frags.append(cur_len)
                cur, cur_len = kind, n
        if cur in ('del', 'ins'):
            frags.append(cur_len)
        if n_eq:
            frags_paired += frags[n_before:]
        if tot and st / tot > STRUCK_PARA_LIMIT:
            over.append((''.join(t.text or '' for t in p.iter(q('t'), q('delText')))[:60], st, tot))
    return {'struck_chars': struck, 'insert_chars': ins, 'paras_over_90pct_struck': len(over),
            'fragments': len(frags), 'fragment_mean_len': round(sum(frags) / len(frags), 1) if frags else 0,
            'fragment_max_len': max(frags) if frags else 0,
            'struck_chars_in_paras_with_kept_text': struck - sum(o[1] for o in over),
            'fragments_in_paras_with_kept_text': len(frags_paired),
            'fragment_mean_len_in_paras_with_kept_text': round(sum(frags_paired) / len(frags_paired), 1) if frags_paired else 0,
            'fragment_max_len_in_paras_with_kept_text': max(frags_paired) if frags_paired else 0,
            'over90_detail': over}


def tamper(xml, kind):
    """Positive control: delete one word from the first run of the given kind ('eq'|'del'|'ins') with >= 3 words."""
    root = etree.fromstring(xml)
    for r in root.iter(q('r')):
        if run_kind(r) != kind:
            continue
        for t in list(r.findall(q('t'))) + list(r.findall(q('delText'))):
            words = (t.text or '').split(' ')
            if len(words) >= 3:
                del words[1]
                t.text = ' '.join(words)
                return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return None


def tamper_merge_paragraphs(xml):
    """Positive control D: merge two adjacent all-baseline paragraphs (no inserted text) into one, losing a boundary."""
    root = etree.fromstring(xml)
    ps = para_list(root)
    for a, b in zip(ps, ps[1:]):
        if a.getparent() is not b.getparent() or ln(a.getparent()) == 'tc':
            continue
        ta, tb = (read_reject(x).text.replace(SEP_CHAR, '') for x in (a, b))
        if ta.strip() and tb.strip() and all(run_kind(r) != 'ins' for r in list(a.iter(q('r'))) + list(b.iter(q('r')))):
            for ch in list(b):
                if ln(ch) != 'pPr':
                    a.append(ch)
            b.getparent().remove(b)
            return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return None


def tamper_clear_mark(xml):
    """Positive control E: take the mark off the first paragraph mark this tool marked as inserted or deleted."""
    root = etree.fromstring(xml)
    for x in para_list(root):
        kind = para_mark(x)
        if kind:
            rpr = para_mark_rpr(x)
            for tag in ((kind,) if MODE == 'track_changes' else ('strike', 'color')):
                for e in rpr.findall(q(tag)):
                    rpr.remove(e)
            return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return None


def zip_checks(clean_path, out_path):
    """testzip, part-by-part byte identity against the clean zip, lxml parse."""
    res = []
    with zipfile.ZipFile(clean_path) as a, zipfile.ZipFile(out_path) as b:
        res.append(('testzip', b.testzip() is None))
        res.append(('namelist identical incl. order', a.namelist() == b.namelist()))
        res.append(('comment identical', a.comment == b.comment))
        same = all(a.read(n) == b.read(n) for n in a.namelist() if n != 'word/document.xml')
        res.append(('all other parts byte-identical', same))
        meta = all((x.date_time, x.compress_type, x.external_attr) == (y.date_time, y.compress_type, y.external_attr)
                   for x, y in zip(a.infolist(), b.infolist()))
        res.append(('ZipInfo date/compress/attr identical', meta))
        etree.fromstring(b.read('word/document.xml'))
        res.append(('document.xml parses (lxml)', True))
    return res


def rgb(text):
    """argparse type: six hex digits, returned in upper case."""
    if not re.fullmatch(r'[0-9A-Fa-f]{6}', text):
        raise argparse.ArgumentTypeError('a colour is six hex digits, for example FF0000')
    return text.upper()


def main(argv=None):
    """CLI entry: refuse what cannot be done, then build, verify, report."""
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--baseline', required=True, help='the earlier version')
    ap.add_argument('--clean', required=True, help='the current version, without marks')
    ap.add_argument('--out', required=True, help='the marked copy to write')
    ap.add_argument('--insert-rgb', type=rgb, default=INSERT_RGB, help='colour of inserted text (colour mode)')
    ap.add_argument('--delete-rgb', type=rgb, default=DELETE_RGB, help='colour of struck deleted text (colour mode)')
    ap.add_argument('--mode', choices=('color', 'track_changes'), default=MODE, help='how changes are shown')
    ap.add_argument('--author', default=AUTHOR, help='w:author for track_changes marks')
    ap.add_argument('--date', default=DATE, help='w:date for track_changes marks (xsd:dateTime)')
    ap.add_argument('--report', help='JSON report path')
    ap.add_argument('--json', action='store_true', help='print one line of JSON with the counts as the last line')
    ap.add_argument('--overwrite', action='store_true', help='replace the output and the report when they exist (default: refuse)')
    ap.add_argument('--merge-gap', type=int, default=MERGE_GAP_TOKENS)
    ap.add_argument('--min-ratio', type=float, default=PARA_MATCH_MIN_RATIO)
    ap.add_argument('--sample', type=int, default=SAMPLE_N, help='print N changed paragraphs side by side (prints document text)')
    ap.add_argument('--sent-ratio', type=float, default=SENT_MATCH_MIN_RATIO)
    ap.add_argument('--absorb', type=float, default=ABSORB_MIN_DICE)
    ap.add_argument('--rescue-ratio', type=float, default=RESCUE_MIN_RATIO)
    a = ap.parse_args(argv)
    try:
        return run(a)
    except Refused as e:
        print(str(e), file=sys.stderr)
    except (OSError, zipfile.BadZipFile, KeyError, etree.XMLSyntaxError) as e:
        print('cannot read the input: %s' % e, file=sys.stderr)
    except AssertionError as e:
        print('this document has a structure the tool does not handle: %s' % e, file=sys.stderr)
    return 2


def run(a):
    """Build and verify in a temporary file; move it to --out only when every check passes."""
    global MERGE_GAP_TOKENS, PARA_MATCH_MIN_RATIO, SENT_MATCH_MIN_RATIO, ABSORB_MIN_DICE, RESCUE_MIN_RATIO
    global MODE, AUTHOR, DATE, INSERT_RGB, DELETE_RGB
    MERGE_GAP_TOKENS, PARA_MATCH_MIN_RATIO = a.merge_gap, a.min_ratio
    SENT_MATCH_MIN_RATIO, ABSORB_MIN_DICE, RESCUE_MIN_RATIO = a.sent_ratio, a.absorb, a.rescue_ratio
    MODE, AUTHOR, DATE, INSERT_RGB, DELETE_RGB = a.mode, a.author, a.date, a.insert_rgb, a.delete_rgb

    for path in (a.baseline, a.clean):
        if not os.path.isfile(path):
            raise Refused('no such file: %s' % path)
        if same_file(path, a.out):
            raise Refused('the output is the same file as an input: %s' % a.out)
        if a.report and same_file(path, a.report):
            raise Refused('the report is the same file as an input: %s' % a.report)
    if a.report and same_file(a.out, a.report):
        raise Refused('the report is the same file as the output: %s' % a.report)
    out_dir = os.path.dirname(os.path.abspath(a.out))
    if not os.path.isdir(out_dir):
        raise Refused('the output folder does not exist: %s' % out_dir)
    for path in (a.out, a.report):
        if path and not a.overwrite and os.path.lexists(path):
            raise Refused('the output already exists and is not overwritten: %s' % path)
    if INSERT_RGB == DELETE_RGB:
        raise Refused('--insert-rgb and --delete-rgb must differ')

    with zipfile.ZipFile(a.baseline) as zb, zipfile.ZipFile(a.clean) as zc:
        base_xml, clean_xml = zb.read('word/document.xml'), zc.read('word/document.xml')
        styles, numids = style_ids(zc)
    (base_xml, n_base), (clean_xml, n_clean) = accept_revisions(base_xml), accept_revisions(clean_xml)
    if MODE == 'color':
        struck, coloured, marks = colour_conflicts(clean_xml)
        if marks:
            raise Refused('the clean document has %d paragraph mark(s) formatted in the insert colour %s or struck in '
                          'the delete colour %s, which is how colour mode shows a new or deleted paragraph; choose '
                          'other colours with --insert-rgb / --delete-rgb' % (marks, INSERT_RGB, DELETE_RGB))
        if coloured:
            raise Refused('the clean document already has %d run(s) in the insert colour %s; '
                          'choose another colour with --insert-rgb' % (coloured, INSERT_RGB))
        if struck:
            raise Refused('the clean document has %d struck-through run(s), which colour mode could not tell from '
                          'deletions; use --mode track_changes' % struck)
    doc_xml, st = build_document(base_xml, clean_xml, styles, numids)
    st['accepted_revisions'] = {'baseline': n_base, 'clean': n_clean}
    fd, tmp = tempfile.mkstemp(prefix='.marked-', suffix='.tmp', dir=out_dir)
    os.close(fd)
    try:
        write_docx(a.clean, tmp, doc_xml)
        failed, summary = report(a, tmp, base_xml, clean_xml, st)
        if not failed:
            if same_file(a.clean, a.out) or same_file(a.baseline, a.out):
                raise Refused('the output is the same file as an input: %s' % a.out)
            place(tmp, a.out, a.overwrite)
    finally:
        if os.path.exists(tmp):  # anything but a completed move: leave no half-written file
            os.unlink(tmp)
    if a.json:
        print(json.dumps(dict(summary, written=not failed)))
    return 1 if failed else 0


def report(a, out_path, base_xml, clean_xml, st):
    """Run every check on the file at out_path, print the result; return (failed checks, summary of counts)."""

    total, failed = 0, 0
    lines = []

    def rec(name, ok, detail=''):
        nonlocal total, failed
        total += 1
        failed += 0 if ok else 1
        lines.append('%s %s %s' % ('PASS' if ok else 'FAIL', name, detail))

    for name, ok in zip_checks(a.clean, out_path):
        rec(name, ok)
    with zipfile.ZipFile(out_path) as zo:
        out_xml = zo.read('word/document.xml')
    v = verify(out_xml, base_xml, clean_xml)
    rec('two-way restoration, paragraph by paragraph (%d paragraphs)' % v['paragraphs_compared'],
        not v['failures'], '; '.join(v['failures'][:5]))
    rec('accept-all has the structure of the clean document (paragraphs with empty ones, breaks, elements that are not text, tables with merged cells, bookmarks)',
        not v['structure_accept'], '; '.join(v['structure_accept'][:3]))
    rec('reject-all has the structure of the baseline (paragraphs with empty ones, breaks, elements that are not text, tables with merged cells)',
        not v['structure_reject'], '; '.join(v['structure_reject'][:3]))
    rec('every w:tc contains a w:p', v['tc_no_p'] == 0, 'w:tc without w:p: %d' % v['tc_no_p'])
    exp_i, exp_d = words_of(st['exp_ins']), words_of(st['exp_del'])
    got_i, got_d = v['ins_by_para'], v['del_by_para']
    ins_lab, del_lab = ('w:ins', 'w:del') if MODE == 'track_changes' else (INSERT_RGB, 'strike')
    rec('words in %s runs == words the diff inserted (%d words)' % (ins_lab, sum(exp_i.values())),
        exp_i == got_i, 'missing %s extra %s' % (list((exp_i - got_i))[:2], list((got_i - exp_i))[:2]))
    rec('words in %s runs == words the diff deleted (%d words)' % (del_lab, sum(exp_d.values())),
        exp_d == got_d, 'missing %s extra %s' % (list((exp_d - got_d))[:2], list((got_d - exp_d))[:2]))
    for label, pick in (('positive control A (word removed from an unchanged run)', 'eq'),
                        ('positive control B (word removed from a deleted run)', 'del'),
                        ('positive control C (word removed from an inserted run)', 'ins')):
        tam = tamper(out_xml, pick)
        if tam is None:
            lines.append('SKIP %s: no suitable run' % label)
            continue
        caught = bool(verify(tam, base_xml, clean_xml)['wrong'])
        rec(label + ' is caught as FAIL', caught)
    for label, fn in (('positive control D (two baseline paragraphs merged into one)', tamper_merge_paragraphs),
                      ('positive control E (the mark taken off an inserted or deleted paragraph mark)', tamper_clear_mark)):
        tam = fn(out_xml)
        if tam is None:
            lines.append('SKIP %s: no suitable site' % label)
            continue
        rec(label + ' is caught as FAIL', bool(verify(tam, base_xml, clean_xml)['wrong']))
    m_after = metrics(out_xml)

    print('\n'.join(lines))
    print('mode: %s (insert runs %d, delete runs %d)' % (MODE, v['insert_runs'], v['delete_runs']))
    print('paragraphs: eq %d, mod %d, ins %d, del %d (empty baseline del %d); paragraph marks: inserted %d, deleted %d; tables: %s' %
          (st['n_eq'], st['n_mod'], st['n_ins'], st['n_del'], st['n_del_empty'], st['para_marks']['ins'],
           st['para_marks']['del'], st['tables']))
    print('dropped: pStyle %d, numPr %d, baseline non-text elements in deleted paras %s' %
          (st['dropped_pstyle'], st['dropped_numpr'], dict(st['dropped_baseline_special'])))
    for w in st['warnings']:
        print('WARNING ' + w)
    if any(st['accepted_revisions'].values()):
        print('pending tracked changes read as accepted: baseline %(baseline)d, clean %(clean)d' % st['accepted_revisions'])
    print('METRICS:', {k: v for k, v in m_after.items() if k != 'over90_detail'})
    mods = [x for x in st['marked'] if x[0] and x[1]]
    random.Random(SAMPLE_SEED).shuffle(mods)
    for i, (bt, ct, mk) in enumerate(mods[:a.sample], 1):
        print('--- sample %d\n BASE : %s\n CLEAN: %s\n MARK : %s' % (i, bt, ct, mk))
    print('VERIFY SUMMARY: checks %d, failures %d' % (total, failed))
    if a.report:
        rep = {'constants': {'MODE': MODE, 'INSERT_RGB': INSERT_RGB, 'DELETE_RGB': DELETE_RGB, 'MERGE_GAP_TOKENS': MERGE_GAP_TOKENS,
                             'PARA_MATCH_MIN_RATIO': PARA_MATCH_MIN_RATIO, 'SENT_MATCH_MIN_RATIO': SENT_MATCH_MIN_RATIO,
                             'RESCUE_MIN_RATIO': RESCUE_MIN_RATIO},
               'paragraphs': {k: st[k] for k in ('n_eq', 'n_mod', 'n_ins', 'n_del', 'n_del_empty')},
               'tables': st['tables'], 'accepted_revisions': st['accepted_revisions'],
               'warnings': st['warnings'], 'paragraph_marks': dict(st['para_marks']),
               'dropped': {'pStyle': st['dropped_pstyle'], 'numPr': st['dropped_numpr'],
                           'baseline_special': dict(st['dropped_baseline_special'])},
               'metrics': {k: v for k, v in m_after.items() if k != 'over90_detail'}, 'verify_lines': lines,
               'checks': total, 'failures': failed}
        with open(a.report, 'w' if a.overwrite else 'x', encoding='utf-8') as f:
            json.dump(rep, f, ensure_ascii=False, indent=1)
    return failed, {'checks': total, 'failures': failed, 'mode': MODE, 'insert_runs': v['insert_runs'],
                    'delete_runs': v['delete_runs'], 'accepted_revisions': st['accepted_revisions'],
                    'warnings': st['warnings'], 'paragraph_marks': dict(st['para_marks']),
                    'paragraphs': {k: st[k] for k in ('n_eq', 'n_mod', 'n_ins', 'n_del', 'n_del_empty')}}


if __name__ == '__main__':
    sys.exit(main())
