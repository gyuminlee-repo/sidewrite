#!/usr/bin/env python3
"""Render the first page of a figure file to PNG and report what the file holds.

  python3 tools/figure_render.py --in fig2.pdf --out fig2.png --dpi 150
  python3 tools/figure_render.py --probe

  --in     PDF, .ai (opened as PDF), SVG, PNG, JPEG or TIFF. Only read.
  --out    the PNG to write. A file that is there is not replaced.
  --dpi    resolution for PDF, .ai and SVG (default 150). A PNG, JPEG or TIFF
           is written at its own pixel size.
  --probe  say whether the renderer can run, and render nothing.

Needs PyMuPDF (pip install pymupdf). Prints one JSON line:

  {"ok": true, "type": "PDF", "pages": 1, "width_px": 625, "height_px": 417,
   "width_pt": 300.0, "height_pt": 200.0, "words": 12, "images": 0,
   "paths": 31, "fonts": [{"name": "Helvetica", "embedded": false}]}

  words   words of live text on the first page (0: the text is outlines or pixels)
  images  embedded pictures on the first page
  paths   vector drawings on the first page
  fonts   fonts the first page names, and whether the file carries them
          (PDF and .ai only; an SVG or a picture file names none here)
  width_pt, height_pt  page size in points; null for a PNG, JPEG or TIFF

On failure {"ok": false, "code": ..., "error": ...}. Exit 0 done, 1 the file
could not be read or written, 2 bad arguments, 3 PyMuPDF is not installed.
"""

import argparse
import json
import os
import sys

VECTOR = {".pdf": "PDF", ".ai": "AI", ".svg": "SVG"}
RASTER = {".png": "PNG", ".jpg": "JPEG", ".jpeg": "JPEG", ".tif": "TIFF", ".tiff": "TIFF"}


class Refused(Exception):
    pass


def _fitz():
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return None
    return fitz


def _save(pix, out):
    """Write the PNG to `out`, which must not be there."""
    data = pix.tobytes("png")
    try:
        with open(out, "xb") as fh:
            fh.write(data)
    except FileExistsError:
        raise Refused("%s is there already" % out)


def render(fitz, src, out, dpi=150):
    ext = os.path.splitext(src)[1].lower()
    if ext in RASTER:
        pix = fitz.Pixmap(src)
        if pix.alpha:
            pix = fitz.Pixmap(pix, 0)  # without the alpha channel
        if pix.colorspace is None or pix.colorspace.n != 3:
            pix = fitz.Pixmap(fitz.csRGB, pix)  # grey and CMYK: to RGB
        _save(pix, out)
        return {"type": RASTER[ext], "pages": 1, "width_px": pix.width, "height_px": pix.height,
                "width_pt": None, "height_pt": None, "words": 0, "images": 1, "paths": 0, "fonts": []}
    if ext not in VECTOR:
        raise Refused("not a figure file this tool reads: %s" % (ext or src))
    # An .ai file saved with PDF compatibility is a PDF.
    with fitz.open(src, filetype="svg" if ext == ".svg" else "pdf") as doc:
        if doc.page_count < 1:
            raise Refused("the file has no page")
        page = doc[0]
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        _save(pix, out)
        fonts = []
        if ext != ".svg":
            # (xref, ext, type, basefont, name, encoding, referencer): ext is
            # "n/a" for a font the file names without carrying it.
            seen = set()
            for f in page.get_fonts(full=True):
                name = f[3] or f[4]
                if name in seen:
                    continue
                seen.add(name)
                fonts.append({"name": name, "embedded": f[1] != "n/a"})
        return {"type": VECTOR[ext], "pages": doc.page_count,
                "width_px": pix.width, "height_px": pix.height,
                "width_pt": round(page.rect.width, 2), "height_pt": round(page.rect.height, 2),
                "words": len(page.get_text("words")),
                "images": len(page.get_images(full=True)),
                "paths": len(page.get_drawings()),
                "fonts": fonts}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--in", dest="src")
    ap.add_argument("--out")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--probe", action="store_true")
    a = ap.parse_args(argv)
    fitz = _fitz()
    if fitz is None:
        print(json.dumps({"ok": False, "code": "no_pymupdf",
                          "error": "PyMuPDF is not installed (pip install pymupdf)"}))
        return 3
    if a.probe:
        print(json.dumps({"ok": True, "pymupdf": fitz.VersionBind}))
        return 0
    if not (a.src and a.out) or a.dpi < 1:
        print(json.dumps({"ok": False, "code": "arguments", "error": "--in and --out are required"}))
        return 2
    try:
        facts = render(fitz, a.src, a.out, a.dpi)
    except Refused as e:
        print(json.dumps({"ok": False, "code": "refused", "error": str(e)}))
        return 1
    except Exception as e:  # MuPDF raises its own error types for a broken file
        print(json.dumps({"ok": False, "code": "unreadable", "error": "%s: %s" % (type(e).__name__, e)}))
        return 1
    print(json.dumps({"ok": True, **facts}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
