#!/usr/bin/env python3
"""Verify a written response letter against its model and the manuscript.
Standard library only.

  python3 tools/response_verify.py --docx letter.docx --model model.json --manuscript manuscript.json
  python3 tools/response_verify.py --selftest

  --docx        the letter as written (tools/response_docx.py)
  --model       the paragraph model it was written from
  --manuscript  JSON list of strings: the manuscript paragraphs in the accepted
                view, in document order

A figure line (kind "figure": an adopted figure change, with the name of the
submitted file) is checked like any other paragraph, on count, text and
colour; it cites a file, so no quote follows it.

The file on disk is read back (tools/response_read.py) and checked on five axes:
  count                the docx has as many paragraphs as the model
  text                 every paragraph reads as its model row
  colour               every paragraph carries exactly the colour of its model row
  location_has_quote   a location line is followed by its quote, unless it is a removal
  quote_in_manuscript  every quote is inside one manuscript paragraph (runs of
                       whitespace count as one space)

Prints one JSON line {"ok", "checked": {"paragraphs", "quotes", "locations"},
"failures": [{"axis", "index", "detail"}]}. Exit 0 pass, 1 fail, 2 an input is
missing or not what it should be.

--selftest plants one defect per axis in a small letter and exits 0 only when
the good letter passes and every defect is found on its own axis.
"""

import argparse
import copy
import json
import re
import sys

import response_docx
import response_read

NONE = response_read.NONE
_WS = re.compile(r"\s+")


def _squash(s):
    return _WS.sub(" ", s).strip()


def verify(read, model, manuscript):
    """read: response_read paragraphs of the docx; model: {"paragraphs": [...]};
    manuscript: list of paragraph texts."""
    rows = model["paragraphs"]
    failures = []

    def fail(axis, index, detail):
        failures.append({"axis": axis, "index": index, "detail": detail})

    aligned = len(read) == len(rows)
    if not aligned:
        fail("count", None, "docx has %d paragraphs, the model %d" % (len(read), len(rows)))
    else:
        for i, (got, want) in enumerate(zip(read, rows)):
            text = response_docx.CONTROL.sub("", str(want.get("text", "")))
            if got["text"] != text:
                fail("text", i, "docx %r, model %r" % (got["text"][:60], text[:60]))
            colour = (want.get("color") or NONE).upper()
            have = set(got["colors"]) or {NONE}
            if text and have != {colour}:
                fail("colour", i, "docx %s, model %s" % ("+".join(sorted(have)), colour))

    locations = quotes = 0
    body = [_squash(p) for p in manuscript]
    for i, row in enumerate(rows):
        kind = row.get("kind")
        if kind == "location":
            locations += 1
            nxt = rows[i + 1].get("kind") if i + 1 < len(rows) else None
            if not row.get("removed") and nxt != "quote":
                fail("location_has_quote", i, "no quote after %r" % str(row.get("text", ""))[:60])
        elif kind == "quote":
            quotes += 1
            # What the file says when it can be told which paragraph is which.
            quote = _squash(read[i]["text"] if aligned else str(row.get("text", "")))
            if not quote or not any(quote in p for p in body):
                fail("quote_in_manuscript", i, "not in any manuscript paragraph: %r" % quote[:60])
    return {"ok": not failures,
            "checked": {"paragraphs": len(rows), "quotes": quotes, "locations": locations},
            "failures": failures}


class BadInput(Exception):
    pass


def _load_json(path, what):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError) as e:
        raise BadInput("%s %s: %s" % (what, path, e))


def verify_files(docx, model_path, manuscript_path):
    manuscript = _load_json(manuscript_path, "manuscript")
    if not isinstance(manuscript, list) or not all(isinstance(p, str) for p in manuscript):
        raise BadInput("manuscript %s: not a JSON list of strings" % manuscript_path)
    model = _load_json(model_path, "model")
    if not isinstance(model, dict) or not isinstance(model.get("paragraphs"), list) \
            or not all(isinstance(p, dict) for p in model["paragraphs"]):
        raise BadInput("model %s: no paragraph list" % model_path)
    try:
        read = response_read.paragraphs(docx)
    except response_read.NotADocx as e:
        raise BadInput("docx %s: %s" % (docx, e))
    return verify(read, model, manuscript)


# ---- self-test -----------------------------------------------------------------

_MODEL = {"paragraphs": [
    {"kind": "reviewer_head", "text": "Reviewer #1", "color": None},
    {"kind": "comment_head", "text": "Comment 1", "color": None},
    {"kind": "comment", "text": "State the unit.", "color": None},
    {"kind": "response", "text": "We now give it.", "color": "C00000"},
    {"kind": "location", "text": "[Revised in 2. Methods, line 9]", "color": "0070C0"},
    {"kind": "quote", "text": "The temperature was held at 30 °C.", "color": "0070C0"},
    {"kind": "location", "text": "[Removed from 1. Introduction]", "color": "0070C0", "removed": True},
    {"kind": "figure", "text": "[Revised Figure 2, file fig2_v2.pdf]", "color": "0070C0"}]}
_MANUSCRIPT = ["2. Methods", "Flasks were shaken. The temperature was held at 30 °C. Samples were drawn."]


def _read(model):
    return response_read.paragraphs_from_xml(response_docx.document_xml(model))


def selftest():
    """[(case, expected axes, found axes)]: one planted defect per axis."""
    def without_quote():
        m = copy.deepcopy(_MODEL)
        del m["paragraphs"][5]
        return m

    def recoloured():
        m = copy.deepcopy(_MODEL)
        m["paragraphs"][3]["color"] = "C00001"
        return m

    def reworded():
        m = copy.deepcopy(_MODEL)
        m["paragraphs"][3]["text"] = "We refuse."
        return m

    def shorter():
        m = copy.deepcopy(_MODEL)
        del m["paragraphs"][2]
        return m

    def refiled():
        m = copy.deepcopy(_MODEL)
        m["paragraphs"][7]["text"] = "[Revised Figure 2, file fig2_v3.pdf]"
        return m

    cases = [
        ("good letter", [], _read(_MODEL), _MODEL, _MANUSCRIPT),
        ("a run in another colour", ["colour"], _read(recoloured()), _MODEL, _MANUSCRIPT),
        ("a paragraph with other words", ["text"], _read(reworded()), _MODEL, _MANUSCRIPT),
        ("a paragraph missing in the docx", ["count"], _read(shorter()), _MODEL, _MANUSCRIPT),
        ("a location line without its quote", ["location_has_quote"], _read(without_quote()), without_quote(), _MANUSCRIPT),
        ("a quote that is not in the manuscript", ["quote_in_manuscript"], _read(_MODEL), _MODEL,
         ["2. Methods", "The temperature was kept at 30 °C."]),
        # A figure line cites a file, not a sentence: the good letter ends with
        # one and no quote, and passes. Its words are checked like any other.
        ("a figure line that names another file", ["text"], _read(refiled()), _MODEL, _MANUSCRIPT),
    ]
    out = []
    for name, want, read, model, manuscript in cases:
        found = sorted({f["axis"] for f in verify(read, model, manuscript)["failures"]})
        out.append({"case": name, "expected": want, "found": found, "ok": found == want})
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--docx")
    ap.add_argument("--model")
    ap.add_argument("--manuscript")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args(argv)
    if a.selftest:
        cases = selftest()
        ok = all(c["ok"] for c in cases)
        print(json.dumps({"ok": ok, "cases": cases}))
        return 0 if ok else 1
    if not (a.docx and a.model and a.manuscript):
        print(json.dumps({"ok": False, "error": "--docx, --model and --manuscript are required"}))
        return 2
    try:
        out = verify_files(a.docx, a.model, a.manuscript)
    except BadInput as e:
        print(json.dumps({"ok": False, "error": str(e)}))
        return 2
    print(json.dumps(out, ensure_ascii=False))
    return 0 if out["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
