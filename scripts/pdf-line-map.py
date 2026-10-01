#!/usr/bin/env python3
"""Map a PDF with margin line numbers (Word "Line Numbers") to text lines.

For each page, digit-only words whose right edge lies left of the body text's
left edge are taken as line numbers. Every other word is assigned to the
nearest line number by vertical centre (within half the number's height).

With --quote, the quote is located in the concatenated lines after removing
all whitespace and hyphens (ASCII '-', U+00AD, U+2010, U+2011) on both sides,
so line wraps, hyphenated line breaks, and real compounds broken at their
hyphen all match. NFKC normalisation folds ligatures.

Usage:
  python3 scripts/pdf-line-map.py <pdf>                      # JSON list of lines
  python3 scripts/pdf-line-map.py <pdf> --quote "..."        # JSON match, exit 1 if absent
  python3 scripts/pdf-line-map.py <pdf> --quotes-json FILE   # many quotes in one pass

Output range: "a-b" (or "a" when the quote sits on one line). Line numbers are
those printed in the PDF margin, so they follow Word's restart setting.

--quotes-json FILE holds [{"id", "quote", "within"?}]. `within` is the text of
the paragraph the quote comes from: when that paragraph is in the PDF exactly
once, only matches inside it count, so a phrase that also occurs elsewhere
still gets one line range. Prints one JSON line
  {"lines_indexed", "first_line", "last_line", "monotonic", "results": [{"id", "match"}]}
and exits 0; a quote that is absent has "match": null. `monotonic` is false
when the margin numbers do not rise through the document (numbering restarts
on each page or section). Exit 2 with {"error": "no_line_numbers"} when the
PDF has no margin numbers, exit 1 when FILE is not such a list.
"""

import argparse
import json
import sys
import unicodedata

import fitz  # PyMuPDF

STRIP_CHARS = set(" \t\r\n\u00a0" + "-\u00ad\u2010\u2011")


def _is_int(word):
    return word.isdigit()


def page_lines(page, page_no):
    words = page.get_text("words")  # (x0, y0, x1, y1, text, block, line, wordno)
    if not words:
        return []
    body = [w for w in words if not _is_int(w[4])]
    if not body:
        return []
    body_left = min(w[0] for w in body)
    numbers = [w for w in words if _is_int(w[4]) and w[2] < body_left]
    if not numbers:
        return []
    anchors = []
    for w in numbers:
        anchors.append({"line_no": int(w[4]), "yc": (w[1] + w[3]) / 2, "tol": max(2.0, (w[3] - w[1]) / 2), "words": []})
    number_ids = {id(w) for w in numbers}
    for w in words:
        if id(w) in number_ids:
            continue
        yc = (w[1] + w[3]) / 2
        best = min(anchors, key=lambda a: abs(a["yc"] - yc))
        if abs(best["yc"] - yc) <= best["tol"]:
            best["words"].append(w)
    out = []
    for a in sorted(anchors, key=lambda a: a["yc"]):
        ws = sorted(a["words"], key=lambda w: w[0])
        out.append({"page": page_no, "line_no": a["line_no"], "text": " ".join(w[4] for w in ws)})
    return out


def extract_lines(pdf_path):
    lines = []
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc, start=1):
            lines.extend(page_lines(page, i))
    return lines


def _squeeze(text):
    """NFKC, then drop whitespace and hyphens."""
    chars = []
    for ch in unicodedata.normalize("NFKC", text):
        if ch not in STRIP_CHARS:
            chars.append(ch)
    return "".join(chars)


def _stream(lines):
    """All lines squeezed into one string, and the line index of each character."""
    parts, owner = [], []
    for idx, ln in enumerate(lines):
        sq = _squeeze(ln["text"])
        parts.append(sq)
        owner.extend([idx] * len(sq))
    return "".join(parts), owner


def _hits(hay, needle):
    out = []
    if not needle:
        return out
    pos = hay.find(needle)
    while pos != -1:
        out.append(pos)
        pos = hay.find(needle, pos + 1)
    return out


def find_quote(lines, quote, within=None):
    """Lines the quote spans, anchored at the line where it starts.

    match_count is the number of matches kept and match_count_document the
    number in the whole PDF. When `within` (the text of the quote's paragraph)
    occurs exactly once, only matches inside it are kept.
    """
    hay, owner = _stream(lines)
    needle = _squeeze(quote)
    starts = _hits(hay, needle)
    total = len(starts)
    if within and total > 1:
        w = _squeeze(within)
        region = _hits(hay, w)
        if len(region) == 1:
            lo, hi = region[0], region[0] + len(w)
            starts = [s for s in starts if lo <= s and s + len(needle) <= hi] or starts
    if not starts:
        return None
    s = starts[0]
    first, last = lines[owner[s]], lines[owner[s + len(needle) - 1]]
    a, b = first["line_no"], last["line_no"]
    return {
        "range": str(a) if (a == b and first["page"] == last["page"]) else f"{a}-{b}",
        "start": {"page": first["page"], "line_no": a},
        "end": {"page": last["page"], "line_no": b},
        "match_count": len(starts),
        "match_count_document": total,
    }


def batch(lines, requests):
    """One result per request, in order: {"id", "match"} (match None when absent)."""
    numbers = [ln["line_no"] for ln in lines]
    return {
        "lines_indexed": len(lines),
        "first_line": numbers[0],
        "last_line": numbers[-1],
        "monotonic": all(a < b for a, b in zip(numbers, numbers[1:])),
        "results": [{"id": r.get("id"), "match": find_quote(lines, r.get("quote") or "", r.get("within"))}
                    for r in requests],
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("pdf")
    ap.add_argument("--quote")
    ap.add_argument("--quotes-json", metavar="FILE")
    args = ap.parse_args(argv)
    lines = extract_lines(args.pdf)
    if args.quotes_json is not None:
        # The PDF is judged before the quotes file is read.
        if not lines:
            print(json.dumps({"error": "no_line_numbers"}))
            return 2
        try:
            with open(args.quotes_json, encoding="utf-8") as fh:
                requests = json.load(fh)
            if not isinstance(requests, list) or not all(isinstance(r, dict) for r in requests):
                raise ValueError("expected a JSON list of objects")
        except (OSError, ValueError) as e:
            print(json.dumps({"error": "bad_quotes_file", "detail": str(e)}))
            return 1
        print(json.dumps(batch(lines, requests), ensure_ascii=False))
        return 0
    if args.quote is None:
        print(json.dumps(lines, ensure_ascii=False, indent=1))
        return 0
    hit = find_quote(lines, args.quote)
    print(json.dumps({"quote": args.quote, "match": hit, "lines_indexed": len(lines)}, ensure_ascii=False))
    return 0 if hit else 1


if __name__ == "__main__":
    sys.exit(main())
