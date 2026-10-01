#!/usr/bin/env python3
"""Read an existing response letter (.docx) and propose profile overrides.
Standard library only.

  python3 tools/response_read.py --docx letter.docx [--profile profile.json]

What is read, per paragraph of the body in document order (table cells
included):
  text        the accepted view: w:ins and w:moveTo runs count, w:del and
              w:moveFrom runs do not
  color       the text colour most characters carry, "NONE" when they carry none
  colors      {colour: characters}; characters are code points
  themes      {colour: [theme names]} for runs coloured through a theme colour
  highlights  {w:highlight value: characters}
  shades      {w:shd fill: characters}
  in_table, tracked

The colour of a run is, in this order: its own w:color, its character style
(w:rStyle, following w:basedOn up to 5 steps), the paragraph style. A theme
colour counts with the hex Word wrote in w:val. "auto" and a theme colour
without a hex are "NONE". Highlight and shading are never text colours.

propose() reads the colours, the location line frame and the earlier answers
off the letter. It only proposes: nothing is saved here.

Prints one JSON line. Exit 0, or 1 with {"ok": false, "error"} when the file
is not a readable .docx.
"""

import argparse
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
SKIP = {W + "del", W + "moveFrom", MC_FALLBACK}  # not part of the accepted text
TRACK = {W + "ins", W + "del", W + "moveFrom", W + "moveTo"}
HEX = re.compile(r"^[0-9A-Fa-f]{6}$")
NONE = "NONE"
MAX_PART_BYTES = 64 << 20
BASED_ON_STEPS = 5

REVIEWER_HEADS = [r"^(?:Reviewer|Referee)\s*#?\s*(\d+)"]
COMMENT_HEADS = [r"^Comment\s*#?\s*(\d+)", r"^Point\s*#?\s*(\d+)", r"^Q\s*#?\s*(\d+)"]
LOCATION = re.compile(r"^\[.+\]$")
FRAME = re.compile(r"^\[(.+? in )(.+?), (lines? [\d,\s\-–]+)\]$")
PREFIX = re.compile(r"^\[[^\]]{1,8}\]\s")


class NotADocx(Exception):
    pass


# ---- reading -----------------------------------------------------------------

def _parse(data, name):
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        raise NotADocx("%s declares a document type" % name)
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        raise NotADocx("%s is not XML: %s" % (name, e))


def _colour_of(el):
    """(colour or 'NONE', theme or None) of a w:color element."""
    val, theme = el.get(W + "val"), el.get(W + "themeColor")
    return (val.upper(), theme) if val and HEX.match(val) else (NONE, theme)


def style_colours(root):
    """styleId -> (colour or 'NONE', theme or None) for styles that set a
    colour themselves or through w:basedOn (at most 5 steps)."""
    own, parent = {}, {}
    for st in root.iter(W + "style"):
        sid = st.get(W + "styleId")
        if not sid:
            continue
        col = st.find(W + "rPr/" + W + "color")
        if col is not None:
            own[sid] = _colour_of(col)
        based = st.find(W + "basedOn")
        if based is not None and based.get(W + "val"):
            parent[sid] = based.get(W + "val")
    out = {}
    for sid in set(own) | set(parent):
        cur = sid
        for _ in range(BASED_ON_STEPS + 1):
            if cur in own:
                out[sid] = own[cur]
                break
            cur = parent.get(cur)
            if cur is None:
                break
    return out


def run_colour(r, styles, para_colour):
    """(colour or 'NONE', theme or None): direct w:color, then w:rStyle, then the paragraph style."""
    rpr = r.find(W + "rPr")
    if rpr is not None:
        col = rpr.find(W + "color")
        if col is not None:
            return _colour_of(col)
        rs = rpr.find(W + "rStyle")
        if rs is not None and rs.get(W + "val") in styles:
            return styles[rs.get(W + "val")]
    return para_colour if para_colour is not None else (NONE, None)


def _run_text(r):
    """Text of one run; None for a run that holds a picture or a text box."""
    parts = []
    for child in r:
        if child.tag == W + "t":
            parts.append(child.text or "")
        elif child.tag == W + "tab":
            parts.append("\t")
        elif child.tag in (W + "br", W + "cr"):
            parts.append("\n")
        elif child.tag not in (W + "rPr", W + "delText", W + "lastRenderedPageBreak"):
            if child.tag in (W + "drawing", W + "pict", W + "object") or child.tag.endswith("}AlternateContent"):
                return None
    return "".join(parts)


def _runs(node):
    """Runs of a paragraph in document order; never descends into SKIP."""
    for child in node:
        if child.tag in SKIP:
            continue
        if child.tag == W + "r":
            yield child
        elif child.tag != W + "pPr":
            yield from _runs(child)


def _body_paragraphs(node, in_table=False):
    """(w:p, in_table) of the body in document order, table cells included."""
    for child in node:
        if child.tag in SKIP:
            continue
        if child.tag == W + "p":
            yield child, in_table
        elif child.tag == W + "tbl":
            yield from _body_paragraphs(child, True)
        else:
            yield from _body_paragraphs(child, in_table)


def _paragraph(p, in_table, styles):
    para_colour = None
    ps = p.find(W + "pPr/" + W + "pStyle")
    if ps is not None:
        para_colour = styles.get(ps.get(W + "val"))
    text, colors, themes, highlights, shades = [], {}, {}, {}, {}
    for r in _runs(p):
        t = _run_text(r)
        if not t:
            continue
        text.append(t)
        n = len(t)
        colour, theme = run_colour(r, styles, para_colour)
        colors[colour] = colors.get(colour, 0) + n
        if theme and theme not in themes.setdefault(colour, []):
            themes[colour].append(theme)
        rpr = r.find(W + "rPr")
        if rpr is not None:
            hl = rpr.find(W + "highlight")
            if hl is not None and hl.get(W + "val") not in (None, "none"):
                highlights[hl.get(W + "val")] = highlights.get(hl.get(W + "val"), 0) + n
            shd = rpr.find(W + "shd")
            if shd is not None and HEX.match(shd.get(W + "fill") or ""):
                fill = shd.get(W + "fill").upper()
                shades[fill] = shades.get(fill, 0) + n
    return {
        "text": "".join(text),
        "color": max(colors, key=colors.get) if colors else NONE,
        "colors": colors,
        "themes": {c: sorted(names) for c, names in themes.items()},
        "highlights": highlights,
        "shades": shades,
        "in_table": in_table,
        "tracked": any(el.tag in TRACK for el in p.iter()),
    }


def paragraphs_from_xml(document_xml, styles_xml=None):
    data = document_xml.encode("utf-8") if isinstance(document_xml, str) else document_xml
    root = _parse(data, "word/document.xml")
    styles = {}
    if styles_xml:
        sdata = styles_xml.encode("utf-8") if isinstance(styles_xml, str) else styles_xml
        styles = style_colours(_parse(sdata, "word/styles.xml"))
    body = root.find(W + "body")
    if body is None:
        raise NotADocx("word/document.xml has no body")
    return [_paragraph(p, in_table, styles) for p, in_table in _body_paragraphs(body)]


def paragraphs(path):
    """[{text, color, colors, themes, highlights, shades, in_table, tracked}]
    for every paragraph of the body. Raises NotADocx."""
    try:
        with zipfile.ZipFile(str(path)) as z:
            names = set(z.namelist())
            if "word/document.xml" not in names:
                raise NotADocx("no word/document.xml in the file")
            parts = {}
            for name in ("word/document.xml", "word/styles.xml"):
                if name not in names:
                    continue
                if z.getinfo(name).file_size > MAX_PART_BYTES:
                    raise NotADocx("%s is too large" % name)
                parts[name] = z.read(name)
    except (zipfile.BadZipFile, OSError) as e:
        raise NotADocx(str(e))
    return paragraphs_from_xml(parts["word/document.xml"], parts.get("word/styles.xml"))


# ---- proposal ----------------------------------------------------------------

def _head(paras, profile_value, candidates):
    """The profile's head pattern when it fits this letter, else the built-in
    candidate with most matching paragraphs, else None."""
    def hits(rx):
        return sum(1 for p in paras if rx.search(p["text"]))
    if isinstance(profile_value, str) and profile_value:
        try:
            rx = re.compile(profile_value)
            if hits(rx):
                return rx
        except re.error:
            pass
    best = max((re.compile(c) for c in candidates), key=hits)
    return best if hits(best) else None


def _number(match, fallback):
    if match.re.groups and match.group(1) and match.group(1).isdigit():
        return int(match.group(1))
    return fallback


def _blocks(paras, reviewer_rx, comment_rx):
    """[{reviewer, n, paras}] : the paragraphs after each comment heading."""
    blocks, reviewer, count = [], 1, 0
    for p in paras:
        m = reviewer_rx.search(p["text"]) if reviewer_rx else None
        if m:
            reviewer, count = _number(m, reviewer + 1), 0
            continue
        m = comment_rx.search(p["text"]) if comment_rx else None
        if m:
            count = _number(m, count + 1)
            blocks.append({"reviewer": reviewer, "n": count, "paras": []})
        elif blocks:
            blocks[-1]["paras"].append(p)
    return blocks


def _roles(block):
    """[(role, paragraph)]: location | quote | response for the coloured
    paragraphs of a block. A paragraph in square brackets is a location line;
    paragraphs of the same colour right after it are its quote; any other
    coloured paragraph is an answer."""
    out, after = [], None
    for p in block["paras"]:
        text = p["text"].strip()
        if LOCATION.match(text):
            out.append(("location", p))
            after = p["color"]
        elif after is not None and p["color"] == after and p["color"] != NONE and text:
            out.append(("quote", p))
        else:
            after = None
            if p["color"] != NONE and text:
                out.append(("response", p))
    return out


def _top(counter):
    return counter.most_common(1)[0][0] if counter else None


def propose(paras, profile=None):
    """{"overrides", "evidence", "previous"} read off an existing letter."""
    profile = profile or {}
    notes = []
    reviewer_rx = _head(paras, profile.get("reviewer_head"), REVIEWER_HEADS)
    comment_rx = _head(paras, profile.get("comment_head"), COMMENT_HEADS)
    blocks = _blocks(paras, reviewer_rx, comment_rx)
    if not blocks:
        notes.append("no comment headings found: answers cannot be told apart")
    roles = [(b, _roles(b)) for b in blocks]

    locations = [p for _, rs in roles for role, p in rs if role == "location"]
    revised = _top(Counter(p["color"] for p in locations if p["color"] != NONE))
    answers = [(b, p) for b, rs in roles for role, p in rs if role == "response"]
    body = _top(Counter(p["color"] for _, p in answers))
    if blocks and not answers:
        notes.append("no coloured paragraphs: answers cannot be told from comments")

    # A third colour is the translation when it comes back in two blocks or
    # all its paragraphs open with the same short bracket tag ("[T] ").
    colors, letter = {}, {}
    others = [(b, p) for b, p in answers if p["color"] not in (body, revised)]
    gloss = _top(Counter(p["color"] for _, p in others))
    if gloss:
        mine = [(b, p) for b, p in others if p["color"] == gloss]
        tags = {m.group(0) if m else None for m in (PREFIX.match(p["text"]) for _, p in mine)}
        tag = tags.pop() if len(tags) == 1 else None
        if tag or len({id(b) for b, _ in mine}) >= 2:
            colors["response_kr"] = gloss
            if tag:
                letter["gloss_prefix"] = tag
        else:
            gloss = None
    if body:
        colors["response_body"] = body
    if revised:
        colors["response_revised"] = revised

    for p in locations:
        m = FRAME.match(p["text"].strip())
        if m:
            letter["location_line"] = "[%s{section}, {lines}]" % m.group(1)
            break

    highlights, shades = Counter(), Counter()
    for p in paras:
        highlights.update(p["highlights"])
        shades.update(p["shades"])
    if highlights:
        colors["highlight"] = _top(highlights)

    previous = {}
    for b, rs in roles:
        text = "\n".join(p["text"].strip() for role, p in rs if role == "response" and p["color"] == body)
        if not text:
            continue
        key = "R%d.%d" % (b["reviewer"], b["n"])
        if key in previous:
            notes.append("%s occurs twice; the first answer is kept" % key)
        else:
            previous[key] = text

    role_of = {body: "response_body", revised: "response_revised", gloss: "response_kr"}
    ev_colors = {}
    for p in paras:
        for colour, chars in p["colors"].items():
            if colour == NONE:
                continue
            e = ev_colors.setdefault(colour, {"paragraphs": 0, "chars": 0, "role": role_of.get(colour),
                                              "sample": p["text"][:60], "theme": []})
            e["chars"] += chars
            if p["color"] == colour:
                e["paragraphs"] += 1
            for name in p["themes"].get(colour, []):
                if name not in e["theme"]:
                    e["theme"].append(name)
    for colour, e in ev_colors.items():
        e["theme"].sort()
        if e["theme"]:
            notes.append("%s is a theme colour: check the value in Word" % colour)
    if any(NONE in p["themes"] for p in paras):
        notes.append("a theme colour without a value was read as no colour")
    tracked = sum(1 for p in paras if p["tracked"])
    if tracked:
        notes.append("%d paragraph(s) hold tracked changes; they were read as accepted" % tracked)

    overrides = {}
    if colors:
        overrides["colors"] = colors
    if letter:
        overrides["response_letter"] = letter
    return {
        "overrides": overrides,
        "evidence": {
            "paragraphs": len(paras),
            "blocks": len(blocks),
            "heads": {"reviewer_head": reviewer_rx.pattern if reviewer_rx else None,
                      "comment_head": comment_rx.pattern if comment_rx else None},
            "colors": ev_colors,
            "highlights": dict(highlights),
            "shades": dict(shades),
            "tracked_paragraphs": tracked,
            "notes": notes,
        },
        "previous": previous,
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--docx", required=True)
    ap.add_argument("--profile", help="JSON file of the profile in use (head patterns)")
    a = ap.parse_args(argv)
    try:
        profile = {}
        if a.profile:
            with open(a.profile, encoding="utf-8") as fh:
                profile = json.load(fh)
        out = propose(paragraphs(a.docx), profile if isinstance(profile, dict) else {})
    except (NotADocx, OSError, ValueError) as e:
        print(json.dumps({"ok": False, "error": str(e)}))
        return 1
    print(json.dumps({"ok": True, **out}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
