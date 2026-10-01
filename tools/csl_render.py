#!/usr/bin/env python3
"""Render references with a CSL style, for comparing with a reference list. Report only.

Optional: it works only when the citeproc-py package can be imported. This
repository does not install it and does not depend on it. Without it the
tool exits with code 3 and says how to get it; nothing else of the add-in
needs it.

  csl_render.py --check
      exit 0 when citeproc-py is available, 3 when it is not.
  csl_render.py --style STYLE.csl --items ITEMS.json
      ITEMS.json is a list of CSL-JSON items. Prints
      {"entries": [{"id", "text"}]} with each item rendered as a bibliography
      entry in plain text.

A dependent style has no layout of its own; render with its parent style.

Exit code: 0 done, 2 bad input (the reason on stderr), 3 citeproc-py missing.
"""

import argparse
import json
import re
import sys

MISSING = (
    "citeproc-py is not installed, so references cannot be rendered. "
    "It is optional: install it yourself with `pip install citeproc-py` if you want this comparison."
)


class BadInput(Exception):
    """The request cannot be carried out; the message says why."""


def load_citeproc():
    """The citeproc names the rendering needs, or None when the package cannot be imported."""
    try:
        import citeproc
        from citeproc.source.json import CiteProcJSON
    except ImportError:
        return None
    return citeproc, CiteProcJSON


def parent_of(style_xml):
    """The parent style name of a dependent style, or None."""
    for link in re.findall(r"<link\b[^>]*>", style_xml):
        if re.search(r'\brel\s*=\s*"independent-parent"', link):
            href = re.search(r'\bhref\s*=\s*"([^"]*)"', link)
            return href.group(1).rstrip("/").split("/")[-1] if href else ""
    return None


def render(style_path, items, names):
    """[{"id", "text"}] for every item, in the order given."""
    citeproc, CiteProcJSON = names
    source = CiteProcJSON(items)
    style = citeproc.CitationStylesStyle(style_path, validate=False)
    bibliography = citeproc.CitationStylesBibliography(style, source, citeproc.formatter.plain)
    ids = [str(item["id"]).lower() for item in items]
    for key in ids:
        bibliography.register(citeproc.Citation([citeproc.CitationItem(key)]))
    texts = [str(entry) for entry in bibliography.bibliography()]
    return [{"id": item["id"], "text": text} for item, text in zip(items, texts)]


def run(a):
    if a.check:
        if load_citeproc() is None:
            print(MISSING, file=sys.stderr)
            return 3
        print("citeproc-py is available")
        return 0
    if not a.style or not a.items:
        raise BadInput("give --check, or both --style and --items")
    try:
        with open(a.style, encoding="utf-8") as f:
            style_xml = f.read()
        with open(a.items, encoding="utf-8") as f:
            items = json.load(f)
    except (OSError, ValueError) as e:
        raise BadInput("cannot read the input: %s" % e)
    parent = parent_of(style_xml)
    if parent is not None:
        raise BadInput("this is a dependent style; render with its parent style %s" % parent)
    if not isinstance(items, list) or not all(isinstance(i, dict) and "id" in i for i in items):
        raise BadInput("--items must be a list of CSL-JSON items, each with an id")
    names = load_citeproc()
    if names is None:
        print(MISSING, file=sys.stderr)
        return 3
    print(json.dumps({"entries": render(a.style, items, names)}, ensure_ascii=False, indent=1))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--check", action="store_true", help="only say whether citeproc-py is available")
    ap.add_argument("--style", help="an independent CSL style file")
    ap.add_argument("--items", help="a JSON file with a list of CSL-JSON items")
    a = ap.parse_args(argv)
    try:
        return run(a)
    except BadInput as e:
        print(str(e), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
