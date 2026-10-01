#!/usr/bin/env python3
"""Write baseline/clean docx pairs of the marked-diff corpus (tests/marked_diff_corpus.py) for the Word check.

  marked_diff_pairs.py --list
      one JSON line: [{"name": ..., "refused": true|false}] for every named pair of the corpus
  marked_diff_pairs.py --out DIR --names FILE
      FILE is a JSON list of pair names. Writes DIR/<nn>/base.docx and DIR/<nn>/clean.docx (nn = 01, 02, ... in
      the order of the list; DIR/<nn> must not exist) and prints one JSON line: [{"name", "dir"}].
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tests"))
import marked_diff_corpus as corpus  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--out")
    ap.add_argument("--names")
    a = ap.parse_args()
    if a.list:
        print(json.dumps([{"name": n, "refused": bool(c[2])} for n, c in corpus.CASES.items()]))
        return 0
    if not a.out or not a.names:
        ap.error("--out and --names are required")
    names = json.loads(Path(a.names).read_text(encoding="utf-8"))
    unknown = [n for n in names if n not in corpus.CASES]
    if unknown:
        print("not in the corpus: %s" % ", ".join(unknown), file=sys.stderr)
        return 2
    made = []
    for i, name in enumerate(names, 1):
        folder = Path(a.out) / ("%02d" % i)
        folder.mkdir(parents=True, exist_ok=False)
        base, clean, _ = corpus.models(corpus.CASES[name])
        corpus.write(base, folder / "base.docx", header="An invented running head")
        corpus.write(clean, folder / "clean.docx", header="An invented running head")
        made.append({"name": name, "dir": str(folder)})
    print(json.dumps(made))
    return 0


if __name__ == "__main__":
    sys.exit(main())
