#!/usr/bin/env python3
"""Generate a public, fake scientific manuscript for the stage-2 spike.

The document has three Heading 1 sections and Word line numbering turned on
(w:lnNumType countBy=1 restart=continuous) so a Word-rendered PDF carries
margin line numbers for scripts/pdf-line-map.py.

All text is invented filler. It names no real study, person, or institution.

Usage: python3 scripts/make-test-manuscript.py [output.docx]
Default output: tests/fixtures/spike_manuscript.docx
"""

import datetime
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

DEFAULT_OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "spike_manuscript.docx"
FIXED_TIME = datetime.datetime(2026, 1, 1, 0, 0, 0)

TITLE = "Growth response of a model bacterium to a synthetic carbon source"

# Phrases a Word test can target with find/replace:
#   "the incubation temperature was held at 30 degrees" (occurs exactly once)
#   "cell-free extract" (a real hyphenated compound; first copy of three paragraphs)
SECTIONS = [
    (
        "1. Introduction",
        [
            "Microbial growth on single carbon sources is a common starting point for studies of "
            "central metabolism. A model bacterium was chosen here because its growth is easy to "
            "follow by optical density and because its genetic tools are widely available.",
            "Previous work in this illustrative text suggested that a synthetic carbon source can "
            "support slow growth, although the response varies with medium composition. The aim of "
            "this fictional study was to describe the growth curve and to test whether a "
            "cell-free extract retains the activity of the first enzyme in the pathway.",
            "The remainder of this manuscript is placeholder prose written only to fill several "
            "pages. It contains no real data, and every value below is invented for a software "
            "test of line numbering and review comments.",
        ],
    ),
    (
        "2. Materials and Methods",
        [
            "Cultures were grown in a defined mineral medium in sealed flasks with shaking. "
            "For every run the incubation temperature was held at 30 degrees and the headspace "
            "was refreshed once per day. Optical density was read at 600 nm in a plate reader.",
            "A cell-free extract was prepared by sonication followed by centrifugation, and the "
            "supernatant was kept on ice until use. Protein concentration was estimated with a "
            "dye-binding assay against a standard curve.",
            "Enzyme activity was measured spectrophotometrically by following the change in "
            "absorbance over five minutes. Each assay was repeated three times, and the mean "
            "and standard deviation were reported for every condition.",
            "Statistical comparisons used a two-sided test with a significance threshold of "
            "0.05. No correction for multiple comparisons was applied because only two "
            "conditions were compared in this invented design.",
        ],
    ),
    (
        "3. Results",
        [
            "Growth on the synthetic carbon source began after a lag of roughly one day and "
            "reached a plateau after four days. The final optical density was lower than on the "
            "reference substrate, which is consistent with a slower uptake step.",
            "The cell-free extract showed measurable activity that declined over two hours on "
            "ice. Heat treatment removed the activity completely, which suggests that the signal "
            "depends on a protein component rather than on a small molecule.",
            "Replicate cultures agreed closely during exponential growth and diverged slightly "
            "at the plateau. The spread at the plateau did not change the main conclusion of "
            "this illustrative example.",
        ],
    ),
]

# Repeat the body so the document spans two to three pages.
BODY_REPEATS = 2


def enable_line_numbers(section):
    """Insert w:lnNumType into sectPr at its schema position (after w:pgMar)."""
    sect_pr = section._sectPr
    for old in sect_pr.findall(qn("w:lnNumType")):
        sect_pr.remove(old)
    ln = OxmlElement("w:lnNumType")
    ln.set(qn("w:countBy"), "1")
    ln.set(qn("w:restart"), "continuous")
    # CT_SectPr order: ... pgSz, pgMar, paperSrc, pgBorders, lnNumType, pgNumType, cols, ...
    successors = ("w:pgNumType", "w:cols", "w:formProt", "w:vAlign", "w:noEndnote",
                  "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
                  "w:printerSettings", "w:sectPrChange")
    for tag in successors:
        nxt = sect_pr.find(qn(tag))
        if nxt is not None:
            nxt.addprevious(ln)
            return
    sect_pr.append(ln)


def build(out_path):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 2.0

    cp = doc.core_properties
    cp.author = "Spike Fixture"
    cp.last_modified_by = "Spike Fixture"
    cp.title = TITLE
    cp.created = FIXED_TIME
    cp.modified = FIXED_TIME
    cp.revision = 1

    doc.add_paragraph(TITLE, style="Title")
    for heading, paras in SECTIONS:
        doc.add_paragraph(heading, style="Heading 1")
        for rep in range(BODY_REPEATS):
            for text in paras:
                # Keep the target phrases unique: only the first copy carries them.
                if rep > 0:
                    text = (text.replace("the incubation temperature was held at 30 degrees",
                                         "the incubation temperature was kept constant")
                                .replace("cell-free extract", "crude extract"))
                doc.add_paragraph(text)

    enable_line_numbers(doc.sections[0])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


def main(argv):
    out = Path(argv[1]) if len(argv) > 1 else DEFAULT_OUT
    print(build(out))


if __name__ == "__main__":
    main(sys.argv)
