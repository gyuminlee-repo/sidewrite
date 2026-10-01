#!/usr/bin/env python3
"""Put what the tool changed next to what the authors changed.

  replay-compare.py --before <docx> --after <docx> --tool <changes.json> --out <dir>

--before  the manuscript the revision started from (clean)
--after   the manuscript the authors submitted (clean)
--tool    changes.json of scripts/e2e/replay.mjs, run on a copy of --before
--out     a folder for the two files below (created when missing)

Writes
  reference_changes.json  the paragraphs that differ between --before and
                          --after: [{ kind, paragraph_index_before,
                          paragraph_index_after, before, after, similarity }]
                          kind: "modified" (the same paragraph, edited),
                          "rewritten" (a removed and an added paragraph that
                          share enough words to be one paragraph written
                          anew), "deleted" or "inserted"
  compare.md              counts, then four lists: paragraphs both changed
                          (before, the authors' text, the tool's text),
                          paragraphs only the authors changed, paragraphs only
                          the tool changed, tool changes that could not be
                          placed on a paragraph of --before; then, per
                          comment, what the tool said instead of changing:
                          its questions for the author and the requests it
                          declared to lie outside this file

The script judges nothing: it lines texts up for a reader. Without --tool
only reference_changes.json and the counts are written.

A paragraph is a w:p of the body in document order, table cells included; a
text box is read once (not its fallback copy). Its text is its w:t runs with
tabs and line breaks as spaces, white space folded. Deleted tracked text
(w:delText) is not text, so a clean file and a file whose changes are all
accepted read the same. Paragraph numbers count every w:p from 0, empty ones
too, which is close to and not always equal to the index Word reports.

Standard library only. The input files are opened for reading and nothing is
written outside --out.
"""
import argparse
import difflib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
P, T, TAB, BR, CR, TXBX = W + "p", W + "t", W + "tab", W + "br", W + "cr", W + "txbxContent"
FALLBACK = MC + "Fallback"

# Two paragraphs are "the same paragraph, changed" from this similarity on.
PAIR_MIN = 0.5
# A removed and an added paragraph are "one paragraph, rewritten" from this
# share of common words on (in order, among what PAIR_MIN left unpaired).
REWRITE_MIN = 0.3
# A tool change is placed on a paragraph by its text from this similarity on.
PLACE_MIN = 0.6


def fold(text):
    return re.sub(r"\s+", " ", text).strip()


def _text(node, out):
    """The text of a paragraph without what a text box in it holds."""
    for child in node:
        if child.tag in (TXBX, FALLBACK):
            continue
        if child.tag == T:
            out.append(child.text or "")
        elif child.tag in (TAB, BR, CR):
            out.append(" ")
        else:
            _text(child, out)


def _boxes(node, out):
    for child in node:
        if child.tag == FALLBACK:
            continue
        if child.tag == TXBX:
            _walk(child, out)
        else:
            _boxes(child, out)


def _walk(node, out):
    for child in node:
        if child.tag == FALLBACK:
            continue
        if child.tag == P:
            parts = []
            _text(child, parts)
            out.append(fold("".join(parts)))
            _boxes(child, out)
        else:
            _walk(child, out)


def paragraphs(path):
    """The paragraph texts of a docx, in document order, empty ones included."""
    with zipfile.ZipFile(path) as z:
        root = ElementTree.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    out = []
    _walk(body if body is not None else root, out)
    return out


def ratio(a, b):
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def word_ratio(a, b):
    """How much of their words two texts share, in order (0 to 1)."""
    return difflib.SequenceMatcher(None, a.lower().split(), b.lower().split(), autojunk=False).ratio()


def pair_block(before, after, score=ratio, least=PAIR_MIN):
    """Pairs (i, j, similarity) of two lists of paragraphs, in order, so that
    the similarities add up to the most; a pair needs `least`."""
    PAIR_MIN = least
    n, m = len(before), len(after)
    sim = [[score(before[i], after[j]) for j in range(m)] for i in range(n)]
    best = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            best[i][j] = max(best[i + 1][j], best[i][j + 1])
            if sim[i][j] >= PAIR_MIN:
                best[i][j] = max(best[i][j], sim[i][j] + best[i + 1][j + 1])
    pairs, i, j = [], 0, 0
    while i < n and j < m:
        if sim[i][j] >= PAIR_MIN and best[i][j] == sim[i][j] + best[i + 1][j + 1]:
            pairs.append((i, j, sim[i][j]))
            i, j = i + 1, j + 1
        elif best[i][j] == best[i + 1][j]:
            i += 1
        else:
            j += 1
    return pairs


def reference_changes(before, after):
    """The paragraphs that differ between two lists of paragraph texts."""
    a = [(i, t) for i, t in enumerate(before) if t]
    b = [(j, t) for j, t in enumerate(after) if t]
    sm = difflib.SequenceMatcher(None, [t for _, t in a], [t for _, t in b], autojunk=False)
    out = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        left, right = a[i1:i2], b[j1:j2]
        pairs = pair_block([t for _, t in left], [t for _, t in right])
        used_l = {i for i, _, _ in pairs}
        used_r = {j for _, j, _ in pairs}
        rows = []
        for i, j, s in pairs:
            rows.append((left[i][0], right[j][0], "modified", left[i][1], right[j][1], round(s, 3)))
        for i, (index, text) in enumerate(left):
            if i not in used_l:
                rows.append((index, None, "deleted", text, "", None))
        for j, (index, text) in enumerate(right):
            if j not in used_r:
                rows.append((None, index, "inserted", "", text, None))
        # Document order: by the paragraph of --before, inserted ones by --after.
        rows.sort(key=lambda r: (r[0] if r[0] is not None else float("inf"), r[1] or 0))
        for ib, ia, kind, tb, ta, s in rows:
            out.append({
                "kind": kind,
                "paragraph_index_before": ib,
                "paragraph_index_after": ia,
                "before": tb,
                "after": ta,
                "similarity": s,
            })
    return pair_rewritten(out)


def pair_rewritten(rows):
    """A second pass over what is left as deleted and inserted: in document
    order, a removed and an added paragraph that share REWRITE_MIN of their
    words become one row of kind "rewritten"."""
    gone = [r for r in rows if r["kind"] == "deleted"]
    new = [r for r in rows if r["kind"] == "inserted"]
    pairs = pair_block([r["before"] for r in gone], [r["after"] for r in new], word_ratio, REWRITE_MIN)
    drop = set()
    for i, j, s in pairs:
        gone[i].update(kind="rewritten", paragraph_index_after=new[j]["paragraph_index_after"],
                       after=new[j]["after"], similarity=round(s, 3))
        drop.add(id(new[j]))
    return [r for r in rows if id(r) not in drop]


def place_tool_changes(report, before):
    """The tool's changes by paragraph of --before.

    report: changes.json. A change carries the paragraph text before and
    after it. It is placed on the paragraph of --before with that text; a
    later change of the same paragraph starts from the text the one before
    left, and is placed with it; a text that is close enough (PLACE_MIN) is
    taken when there is one best paragraph. Returns (placed, unplaced):
    placed: { index: { before, after, refs, change_ids } } with the text
    before the first and after the last change of that paragraph.
    """
    where = {}
    for i, t in enumerate(before):
        if t:
            where.setdefault(t, i)
    placed, unplaced, now = {}, [], {}
    for comment in report:
        for ch in comment.get("changes", []):
            tb, ta = fold(ch.get("before") or ""), fold(ch.get("after") or "")
            index = now.get(tb) if tb else None
            if index is None and tb:
                index = where.get(tb)
            if index is None and tb:
                scored = sorted(((ratio(tb, t), i) for i, t in enumerate(before) if t), reverse=True)
                if scored and scored[0][0] >= PLACE_MIN and (len(scored) == 1 or scored[0][0] > scored[1][0]):
                    index = scored[0][1]
            if index is None:
                unplaced.append({"ref": comment.get("ref"), "change_id": ch.get("change_id"),
                                 "kind": ch.get("kind"), "state": ch.get("state"), "before": tb, "after": ta})
                continue
            row = placed.setdefault(index, {"before": before[index], "after": ta, "refs": [], "change_ids": []})
            row["after"] = ta
            if comment.get("ref") not in row["refs"]:
                row["refs"].append(comment.get("ref"))
            row["change_ids"].append(ch.get("change_id"))
            now[ta] = index
    return placed, unplaced


def cell(text):
    """A text as one cell of a Markdown table."""
    return fold(text).replace("|", "\\|") or "(empty)"


def table(head, rows):
    lines = ["| " + " | ".join(head) + " |", "|" + "|".join(" --- " for _ in head) + "|"]
    lines += ["| " + " | ".join(r) + " |" for r in rows]
    return "\n".join(lines)


def said_instead(report):
    """Per comment, what the tool said instead of changing the text: the
    questions it put to the author and the requests it declared outside this
    file, with the state of every request. So a reader can tell a request the
    tool passed over in silence from one it asked about."""
    out = ["", "## What the tool said instead of changing (by comment)"]
    any_row = False
    for c in report:
        questions, outside = c.get("questions") or [], c.get("outside") or []
        states = [(q.get("id"), q.get("state")) for q in c.get("requirements") or []]
        if not questions and not outside and not any(s in ("open", "needs_author", "outside_this_file")
                                                       for _, s in states):
            continue
        any_row = True
        out += ["", "### %s" % c.get("ref"), ""]
        if states:
            out.append("Requests: " + ", ".join("%s %s" % (i, s) for i, s in states))
            out.append("")
        if questions:
            out.append(table(
                ["Question", "Request", "Asked", "Why", "Answer"],
                [[str(q.get("id")), str(q.get("requirement_id") or "(none)"), cell(q.get("question") or ""),
                  cell(q.get("why") or ""), cell(q.get("answer") or "")] for q in questions]))
            out.append("")
        if outside:
            out.append(table(
                ["Request declared outside this file", "Why"],
                [[str(n.get("requirement_id")), cell(n.get("why") or "")] for n in outside]))
            out.append("")
        if not questions and not outside:
            out.append("No question and no outside note.")
            out.append("")
    if not any_row:
        out += ["", "No question, no outside note and no request left open."]
    return out


def compare_md(before, after, reference, placed, unplaced, tool_given, report=None):
    in_before = [r for r in reference if r["paragraph_index_before"] is not None]
    ref_index = {r["paragraph_index_before"]: r for r in in_before}
    both = sorted(set(ref_index) & set(placed))
    only_ref = [r for r in reference if r["paragraph_index_before"] not in placed]
    only_tool = sorted(set(placed) - set(ref_index))
    count = lambda kind: sum(1 for r in reference if r["kind"] == kind)
    share = lambda n, d: "n/a" if not d else "%d of %d (%.0f%%)" % (n, d, 100.0 * n / d)
    out = [
        "# Replay comparison",
        "",
        "Texts are lined up paragraph by paragraph. Nothing here is a judgment.",
        "",
        "## Counts",
        "",
        "- Paragraphs with text: %d before, %d after" % (sum(1 for t in before if t), sum(1 for t in after if t)),
        "- Changed by the authors: %d modified, %d rewritten, %d deleted, %d inserted"
        % (count("modified"), count("rewritten"), count("deleted"), count("inserted")),
    ]
    if tool_given:
        out += [
            "- Paragraphs changed by the tool: %d" % len(placed),
            "- Tool changes not placed on a paragraph: %d" % len(unplaced),
            "- Changed by both: %d" % len(both),
            "- Of the paragraphs the authors modified, rewrote or deleted, also changed by the tool: %s"
            % share(len(both), len(in_before)),
            "- The same, by what the authors did: %s"
            % ", ".join("%s %s" % (k, share(sum(1 for i in both if ref_index[i]["kind"] == k),
                                            sum(1 for r in in_before if r["kind"] == k)))
                        for k in ("modified", "rewritten", "deleted")),
            "- Of the paragraphs the tool changed, also changed by the authors: %s"
            % share(len(both), len(placed)),
            "",
            "## Changed by both",
            "",
            table(
                ["Paragraph", "Comments", "Authors did", "Before", "Authors", "Tool"],
                [[str(i), ", ".join(placed[i]["refs"]), ref_index[i]["kind"], cell(before[i]),
                  cell(ref_index[i]["after"]) if ref_index[i]["kind"] != "deleted" else "(deleted)",
                  cell(placed[i]["after"])] for i in both],
            ) if both else "None.",
            "",
            "## Changed by the authors only",
        ]
    else:
        out += ["", "## Changed by the authors"]
    out += [
        "",
        table(
            ["Paragraph", "Kind", "Before", "Authors"],
            [[str(r["paragraph_index_before"]) if r["paragraph_index_before"] is not None
              else "after %d" % r["paragraph_index_after"],
              r["kind"], cell(r["before"]), cell(r["after"])] for r in only_ref],
        ) if only_ref else "None.",
    ]
    if tool_given:
        out += [
            "",
            "## Changed by the tool only",
            "",
            table(
                ["Paragraph", "Comments", "Before", "Tool"],
                [[str(i), ", ".join(placed[i]["refs"]), cell(placed[i]["before"]), cell(placed[i]["after"])]
                 for i in only_tool],
            ) if only_tool else "None.",
            "",
            "## Tool changes not placed on a paragraph",
            "",
            table(
                ["Comment", "Change", "Kind", "Before", "Tool"],
                [[str(u["ref"]), str(u["change_id"]), str(u["kind"]), cell(u["before"]), cell(u["after"])]
                 for u in unplaced],
            ) if unplaced else "None.",
        ]
        out += said_instead(report or [])
    while out and out[-1] == "":
        out.pop()
    return "\n".join(out) + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--before", required=True)
    ap.add_argument("--after", required=True)
    ap.add_argument("--tool")
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    before, after = paragraphs(a.before), paragraphs(a.after)
    reference = reference_changes(before, after)
    placed, unplaced, report = {}, [], []
    if a.tool:
        report = json.loads(Path(a.tool).read_text(encoding="utf-8"))
        placed, unplaced = place_tool_changes(report, before)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    (out / "reference_changes.json").write_text(
        json.dumps(reference, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (out / "compare.md").write_text(
        compare_md(before, after, reference, placed, unplaced, bool(a.tool), report), encoding="utf-8")
    kinds = {k: sum(1 for r in reference if r["kind"] == k)
             for k in ("modified", "rewritten", "deleted", "inserted")}
    # Counts only on the screen: no text of the manuscript.
    print(json.dumps({
        "paragraphs_before": sum(1 for t in before if t),
        "paragraphs_after": sum(1 for t in after if t),
        **kinds,
        "tool_paragraphs": len(placed),
        "tool_unplaced": len(unplaced),
        "tool_questions": sum(len(c.get("questions") or []) for c in report),
        "tool_outside": sum(len(c.get("outside") or []) for c in report),
        "both": len(set(placed) & {r["paragraph_index_before"] for r in reference}),
        "out": str(out),
    }))
    return 0


if __name__ == "__main__":
    sys.exit(main())
