#!/usr/bin/env python3
"""Write an invented figure in two versions, for the figure lane tests.

  python3 scripts/e2e/make_figure_fixture.py --out <folder>

The figure is a made-up growth curve (no data of any study): optical density
against time, with the axis label "OD600". The "after" version has a second
curve, a legend and a longer time axis label, and is wider, so a comparison
has something to show and the two pictures differ in pixel size.

  fig2_before.svg  fig2_after.svg     always (standard library only)
  fig2_before.pdf  fig2_after.pdf     when PyMuPDF is installed

Prints one JSON line {"ok": true, "files": [...], "pdf": true|false}.
"""

import argparse
import json
import math
from pathlib import Path

FONT = 'font-family="Helvetica, Arial, sans-serif"'


def _curve(x0, y0, w, h, rate, top):
    """A logistic curve inside the plot box, as an SVG path."""
    pts = []
    for i in range(41):
        t = i / 40
        y = top / (1 + math.exp(-rate * (t - 0.45)))
        pts.append("%s%.1f %.1f" % ("M" if i == 0 else "L", x0 + t * w, y0 + h - y * h))
    return " ".join(pts)


def svg(after):
    width, height = (400, 240) if after else (320, 240)
    x0, y0, w, h = 50, 20, width - 70, 170
    parts = [
        '<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">'
        % (width, height, width, height),
        '<rect width="%d" height="%d" fill="white"/>' % (width, height),
        '<path d="M%d %d L%d %d L%d %d" stroke="black" stroke-width="1.2" fill="none"/>'
        % (x0, y0, x0, y0 + h, x0 + w, y0 + h),
        '<path d="%s" stroke="#1f77b4" stroke-width="2" fill="none"/>' % _curve(x0, y0, w, h, 9, 0.85),
        '<text x="14" y="125" %s font-size="12" transform="rotate(-90 14 125)">OD600</text>' % FONT,
        '<text x="%d" y="%d" %s font-size="12">%s</text>'
        % (x0 + w / 2 - 20, y0 + h + 30, FONT, "Time (h)" if after else "Time"),
    ]
    if after:
        parts += [
            '<path d="%s" stroke="#d62728" stroke-width="2" fill="none" stroke-dasharray="5 3"/>'
            % _curve(x0, y0, w, h, 6, 0.6),
            '<text x="%d" y="34" %s font-size="11" fill="#1f77b4">treated</text>' % (x0 + 20, FONT),
            '<text x="%d" y="50" %s font-size="11" fill="#d62728">control</text>' % (x0 + 20, FONT),
        ]
    parts.append("</svg>")
    return "\n".join(parts) + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    files = []
    for name, after in (("fig2_before", False), ("fig2_after", True)):
        (out / (name + ".svg")).write_text(svg(after), encoding="utf-8")
        files.append(name + ".svg")
    try:
        import fitz  # PyMuPDF
    except ImportError:
        fitz = None
    if fitz is not None:
        for name in ("fig2_before", "fig2_after"):
            with fitz.open(str(out / (name + ".svg"))) as doc:
                (out / (name + ".pdf")).write_bytes(doc.convert_to_pdf())
            files.append(name + ".pdf")
    print(json.dumps({"ok": True, "files": files, "pdf": fitz is not None}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
