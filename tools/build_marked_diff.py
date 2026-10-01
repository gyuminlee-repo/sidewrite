#!/usr/bin/env python3
"""Build a "marked" docx from a baseline docx and a clean docx: the clean document with every change shown.

Colour mode (default): deleted text is struck through in --delete-rgb, inserted text is in --insert-rgb, everything
else is unchanged. Track-changes mode (--mode track_changes): the same diff is written as Word revision marks
(inserted runs inside w:ins, deleted runs inside w:del with w:delText, a wholly deleted paragraph gets a deleted
paragraph mark), with --author and --date on every mark. The diff is the same in both modes; only the writer differs.
The output is the clean docx with only word/document.xml replaced.

How the diff is made: paragraphs are aligned first (exact matches anchor the alignment), changed paragraphs are
diffed word by word, and runs of paragraphs without a counterpart are re-paired sentence by sentence, so a rewritten
paragraph shows the sentences that stayed. A paragraph that moved (a re-sorted reference list, for example) is paired
with its new place when it is similar enough. When sentences of several baseline paragraphs end up struck inside one
clean paragraph, a struck line break keeps the baseline paragraph boundaries visible.

What is read how: pending tracked changes are read as accepted in both files. Text that is struck through in the
baseline is read as already deleted. In colour mode a clean file that already has struck-through text, or text in the
insert colour, cannot be marked (the marks could not be told apart) and is refused.

The result is checked before it is written: accepting every mark must give the clean text and rejecting every mark
the baseline text, paragraph by paragraph, and the check itself is tested by damaging a copy. When a check fails
nothing is written and the exit code is 1. The output is never one of the inputs, and an output or report file
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
SEP_CHAR = '\v'                # text form of the struck line break that separates baseline paragraphs inside one paragraph
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
DROP_IN_BASELINE = {'fldChar', 'instrText', 'drawing', 'pict', 'footnoteReference', 'endnoteReference',
                    'lastRenderedPageBreak', 'commentReference', 'object', 'sym'}


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


# ---------------------------------------------------------------- paragraph reading
class Para:
    """Text of a paragraph with per-character format handles (key, rPr element, hyperlink element)."""
    def __init__(self):
        self.text = ''
        self.fmt = []
        self.special = collections.Counter()


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
                if ch.get(q('type')) in ('page', 'column'):
                    s = '\f'
                elif ch.get(q('clear')) == 'none':  # separator written by this tool (see make_run)
                    s = SEP_CHAR
                else:
                    s = '\n'
            else:
                para.special[t] += 1
                continue
            parts.append(s)
            para.fmt.extend([f] * len(s))

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
                walk(ch, link, mark)
            elif t in ('moveFrom', 'pPr'):
                continue
            else:
                para.special[t] += 1

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
    """Turn ops into merged segments [(kind, text, key, rpr, link)]; deleted text never keeps a hyperlink or \\f."""
    segs = []
    for kind, s, e in ops:
        src = b if kind == 'del' else c
        for i in range(s, e):
            ch = src.text[i]
            if kind == 'del' and ch == '\f':
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
    for part in re.split('([\t\n\f%s])' % SEP_CHAR, text):
        if part == '':
            continue
        if part == '\t':
            r.append(etree.Element(q('tab')))
        elif part == '\n':
            r.append(etree.Element(q('br')))
        elif part == '\f':
            br = etree.SubElement(r, q('br'))
            br.set(q('type'), 'page')
        elif part == SEP_CHAR:
            br = etree.SubElement(r, q('br'))
            br.set(q('clear'), 'none')
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
        r = make_run(kind, text, rpr, styles)
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
    """Mark every text-bearing run of a clean-only paragraph as inserted (structure kept):
    colour mode applies the insert colour, track_changes mode wraps the run in <w:ins>."""
    n = 0
    for r in list(p.iter(q('r'))):
        if any(ln(ch) in ('t', 'tab', 'br') for ch in r):
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


# ---------------------------------------------------------------- paragraph-level operations
def rebuild_para(p, b, c, styles, stats, segs=None):
    """Rewrite a clean paragraph as a word-level diff against its baseline counterpart (or prebuilt segments)."""
    bad = {k: v for k, v in c.special.items() if k not in SAFE_CLEAN_TAGS and k != 'lastRenderedPageBreak'}
    assert not bad, 'unsupported element in changed clean paragraph %r: %s' % (c.text[:60], bad)
    if segs is None:
        segs = segments(diff_ops(b.text, c.text), b, c)
    bms = [e for e in p.iter(q('bookmarkStart'), q('bookmarkEnd'))]
    for e in bms:
        e.getparent().remove(e)
    for ch in list(p):
        if ln(ch) != 'pPr':
            p.remove(ch)
    for e in bms:
        if ln(e) == 'bookmarkStart':
            p.append(e)
    for e in build_children(segs, styles):
        p.append(e)
    for e in bms:
        if ln(e) == 'bookmarkEnd':
            p.append(e)
    stats['marked'].append((b if isinstance(b, str) else b.text, c.text, marked_string(segs)))
    stats['exp_ins'].append(''.join(t for k, t, *_ in segs if k == 'ins'))
    stats['exp_del'].append(''.join(t for k, t, *_ in segs if k == 'del'))


def make_deleted_para(bp, b, styles, numids, stats):
    """New paragraph holding baseline text as deletions; pPr comes from the baseline without structure elements."""
    p = etree.Element(q('p'))
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
    if MODE == 'track_changes':   # deleted paragraph mark, so accept-all removes the paragraph itself
        if ppr is None:
            ppr = etree.Element(q('pPr'))
        prpr = etree.SubElement(ppr, q('rPr'))
        prpr.append(rev_mark('del'))
    if ppr is not None and len(ppr):
        p.append(ppr)
    for k, v in b.special.items():
        if k in DROP_IN_BASELINE:
            stats['dropped_baseline_special'][k] += v
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
                A struck line break keeps baseline paragraph boundaries (and a struck heading's own line) visible."""
                if bi is None:
                    if state['short_del']:
                        parts.append(['del', SEP_CHAR, '', None, None])
                    state['short_del'] = False
                else:
                    if state['last_b'] is not None and bi != state['last_b']:
                        parts.append(['del', SEP_CHAR, '', None, None])
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
    return out


def apply_container(bps, cps, styles, numids, stats):
    """Diff one list of baseline paragraphs against one list of clean paragraphs, editing the clean tree in place."""
    B = [read_para(p) for p in bps]
    assert not any(SEP_CHAR in x.text for x in B), 'baseline already contains a separator break'
    C = [read_para(p, skip_rpr=lambda r: False) for p in cps]
    ali = align_paras([x.text for x in B], [x.text for x in C])
    ali = sentence_realign(ali, B, C, SENT_MATCH_MIN_RATIO)
    for tag, bi, ci, extra in ali:
        if tag == 'eq':
            stats['n_eq'] += 1
        elif tag == 'seg':
            assert not C[ci].special.get('drawing'), 'text differs in clean paragraph with drawing: %r' % C[ci].text[:60]
            segs, used = extra
            rebuild_para(cps[ci], ' \u00b6 '.join(B[u].text for u in used), C[ci], styles, stats, segs)
            stats['n_mod'] += 1
        elif tag == 'ins':
            stats['n_ins'] += 1
            if C[ci].text:
                colorize_in_place(cps[ci])
                stats['exp_ins'].append(C[ci].text)
                stats['marked'].append(('', C[ci].text, '{+%s+}' % C[ci].text))
    last = None
    for idx, (tag, bi, ci, extra) in enumerate(ali):
        if tag != 'del':
            last = cps[ci] if ci is not None else last
            continue
        stats['n_del'] += 1
        if not B[bi].text.strip():
            stats['n_del_empty'] += 1
            continue
        newp = make_deleted_para(bps[bi], B[bi], styles, numids, stats)
        nxt = next((cps[a[2]] for a in ali[idx + 1:] if a[0] != 'del'), None)
        if nxt is not None:
            nxt.addprevious(newp)
        else:
            assert last is not None, 'no clean paragraph to anchor a deleted paragraph'
            last.addnext(newp)
            last = newp
        stats['deleted_paras'].append(B[bi].text[:100])


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


def table_dims(t):
    """Cell counts per row."""
    return [len(tr.findall(q('tc'))) for tr in t.findall(q('tr'))]


def cell_paras(tc):
    """Direct paragraph children of a cell."""
    ps = tc.findall(q('p'))
    assert len(list(tc.iter(q('p')))) == len(ps), 'paragraphs nested inside a cell child are not supported'
    return ps


def deleted_table(bt, styles, numids, stats):
    """Baseline table copy with every paragraph converted to deleted text (each cell keeps a paragraph)."""
    t = copy.deepcopy(bt)
    for tc in t.iter(q('tc')):
        for p in tc.findall(q('p')):
            b = read_para(p)
            newp = make_deleted_para(p, b, styles, numids, stats)
            p.addprevious(newp)
            tc.remove(p)
        if tc.find(q('p')) is None:
            tc.append(etree.Element(q('p')))
    return t


def build_document(base_xml, clean_xml, styles, numids):
    """Return (new document.xml bytes, stats)."""
    broot, croot = etree.fromstring(base_xml), etree.fromstring(clean_xml)
    stats = {'n_eq': 0, 'n_mod': 0, 'n_ins': 0, 'n_del': 0, 'n_del_empty': 0, 'marked': [], 'exp_ins': [],
             'exp_del': [], 'deleted_paras': [], 'dropped_pstyle': 0, 'dropped_numpr': 0,
             'dropped_baseline_special': collections.Counter(), 'tables': []}
    apply_container(body_paras(broot), body_paras(croot), styles, numids, stats)
    bts, cts = top_tables(broot), top_tables(croot)
    for k, ct in enumerate(cts):
        if k < len(bts) and table_dims(bts[k]) == table_dims(ct):
            for btc, ctc in zip(bts[k].iter(q('tc')), ct.iter(q('tc'))):
                apply_container(cell_paras(btc), cell_paras(ctc), styles, numids, stats)
            stats['tables'].append('table %d: cell-by-cell (%s)' % (k + 1, table_dims(ct)))
        else:
            for p in ct.iter(q('p')):
                text = read_para(p, lambda r: False).text
                if text:
                    colorize_in_place(p)
                    stats['exp_ins'].append(text)   # the check counts inserted text per paragraph: cells are paragraphs too
                    stats['marked'].append(('', text, '{+%s+}' % text))
            stats['tables'].append('table %d: whole-table insert (baseline dims %s, clean dims %s)'
                                   % (k + 1, table_dims(bts[k]) if k < len(bts) else None, table_dims(ct)))
            if k < len(bts):
                ct.addprevious(deleted_table(bts[k], styles, numids, stats))
                stats['tables'].append('table %d: baseline table inserted as deleted' % (k + 1))
    for k in range(len(cts), len(bts)):
        if not cts and not ''.join(bts[k].itertext()).strip():
            # textless baseline table (empty layout table) removed in clean: nothing to strike, no anchor needed
            stats['tables'].append('baseline table %d: textless, removed in clean, nothing to strike' % (k + 1))
            continue
        assert cts, 'baseline has tables but clean has none: no anchor'
        cts[-1].addnext(deleted_table(bts[k], styles, numids, stats))
        stats['tables'].append('baseline table %d: deleted, placed after last clean table' % (k + 1))
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
    """(struck runs, runs in the insert colour) with text in the clean document: colour mode cannot mark such a file."""
    struck = coloured = 0
    for r in etree.fromstring(clean_xml).iter(q('r')):
        if not any(ln(ch) in ('t', 'tab', 'br') for ch in r if isinstance(ch.tag, str)):
            continue
        rpr = r.find(q('rPr'))
        if is_struck(rpr):
            struck += 1
        elif rpr_color(rpr) == INSERT_RGB:
            coloured += 1
    return struck, coloured


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
    """Paragraph-level two-way restoration and colour checks; returns dict with counts and failure list."""
    m, b, c = (etree.fromstring(x) for x in (marked_xml, base_xml, clean_xml))
    acc = nonempty([read_accept(p).text for p in para_list(m)])
    orig_raw = nonempty([read_reject(p).text.replace('\f', '') for p in para_list(m)])
    orig = nonempty([t.replace(SEP_CHAR, '') for t in orig_raw])
    new = nonempty([read_para(p, lambda r: False).text for p in para_list(c)])
    old = nonempty([read_para(p).text.replace('\f', '') for p in para_list(b)])
    fails = []
    if len(acc) != len(new):
        fails.append('accept-all == clean: paragraph count %d vs %d' % (len(acc), len(new)))
    for i, (x, y) in enumerate(zip(acc, new)):
        if x != y:
            fails.append('accept-all == clean: paragraph #%d differs: %r vs %r' % (i, x[:80], y[:80]))
    fails += group_compare('reject-all == baseline', orig, old)
    fails += boundary_check(orig_raw, old)
    tc_no_p = sum(1 for tc in m.iter(q('tc')) if tc.find(q('p')) is None)
    ins_by_para = collections.Counter(read_only(p, 'ins').text for p in para_list(m))
    del_by_para = collections.Counter(read_only(p, 'del').text for p in para_list(m))
    return {'paragraphs_compared': len(new) + len(old), 'failures': fails,
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


def tamper_drop_separators(xml):
    """Positive control E: delete every struck separator break, gluing the baseline paragraphs it separated."""
    root = etree.fromstring(xml)
    brs = [br for br in root.iter(q('br')) if br.get(q('clear')) == 'none']
    for br in brs:
        br.getparent().remove(br)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True) if brs else None


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
    if MODE == 'color':
        struck, coloured = colour_conflicts(clean_xml)
        if coloured:
            raise Refused('the clean document already has %d run(s) in the insert colour %s; '
                          'choose another colour with --insert-rgb' % (coloured, INSERT_RGB))
        if struck:
            raise Refused('the clean document has %d struck-through run(s), which colour mode could not tell from '
                          'deletions; use --mode track_changes' % struck)
    doc_xml, st = build_document(base_xml, clean_xml, styles, numids)
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
    rec('every w:tc contains a w:p', v['tc_no_p'] == 0, 'w:tc without w:p: %d' % v['tc_no_p'])
    exp_i, exp_d = collections.Counter(x for x in st['exp_ins'] if x), collections.Counter(x for x in st['exp_del'] if x)
    got_i = collections.Counter(k for k in v['ins_by_para'].elements() if k)
    got_d = collections.Counter(k for k in v['del_by_para'].elements() if k)
    ins_lab, del_lab = ('w:ins', 'w:del') if MODE == 'track_changes' else (INSERT_RGB, 'strike')
    rec('%s runs concatenated per paragraph == expected insert text (%d paragraphs)' % (ins_lab, sum(exp_i.values())),
        exp_i == got_i, 'missing %s extra %s' % (list((exp_i - got_i))[:2], list((got_i - exp_i))[:2]))
    rec('%s runs concatenated per paragraph == expected delete text (%d paragraphs)' % (del_lab, sum(exp_d.values())),
        exp_d == got_d, 'missing %s extra %s' % (list((exp_d - got_d))[:2], list((got_d - exp_d))[:2]))
    for label, pick in (('positive control A (word removed from an unchanged run)', 'eq'),
                        ('positive control B (word removed from a deleted run)', 'del'),
                        ('positive control C (word removed from an inserted run)', 'ins')):
        tam = tamper(out_xml, pick)
        if tam is None:
            lines.append('SKIP %s: no suitable run' % label)
            continue
        caught = bool(verify(tam, base_xml, clean_xml)['failures'])
        rec(label + ' is caught as FAIL', caught)
    for label, fn in (('positive control D (two baseline paragraphs merged into one)', tamper_merge_paragraphs),
                      ('positive control E (separator breaks removed)', tamper_drop_separators)):
        tam = fn(out_xml)
        if tam is None:
            lines.append('SKIP %s: no suitable site' % label)
            continue
        rec(label + ' is caught as FAIL', bool(verify(tam, base_xml, clean_xml)['failures']))
    m_after = metrics(out_xml)

    print('\n'.join(lines))
    print('mode: %s (insert runs %d, delete runs %d)' % (MODE, v['insert_runs'], v['delete_runs']))
    print('paragraphs: eq %d, mod %d, ins %d, del %d (empty baseline del %d); tables: %s' %
          (st['n_eq'], st['n_mod'], st['n_ins'], st['n_del'], st['n_del_empty'], st['tables']))
    print('dropped: pStyle %d, numPr %d, baseline non-text elements in deleted paras %s' %
          (st['dropped_pstyle'], st['dropped_numpr'], dict(st['dropped_baseline_special'])))
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
               'tables': st['tables'],
               'dropped': {'pStyle': st['dropped_pstyle'], 'numPr': st['dropped_numpr'],
                           'baseline_special': dict(st['dropped_baseline_special'])},
               'metrics': {k: v for k, v in m_after.items() if k != 'over90_detail'}, 'verify_lines': lines,
               'checks': total, 'failures': failed}
        with open(a.report, 'w' if a.overwrite else 'x', encoding='utf-8') as f:
            json.dump(rep, f, ensure_ascii=False, indent=1)
    return failed, {'checks': total, 'failures': failed, 'mode': MODE, 'insert_runs': v['insert_runs'],
                    'delete_runs': v['delete_runs'],
                    'paragraphs': {k: st[k] for k in ('n_eq', 'n_mod', 'n_ins', 'n_del', 'n_del_empty')}}


if __name__ == '__main__':
    sys.exit(main())
