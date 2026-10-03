#!/usr/bin/env node
// Counts of the places of the terms the reviewers named, in a replay run:
// node scripts/e2e/term-family-count.mjs <run folder>
//
// Reads <run>/manuscript.revision.json (its comments and snapshot.paragraphs,
// the text after the turns) and, when present, <run>/changes.json (which
// terms the replay listed per comment). Prints counts only, never a term or
// a paragraph: a run holds an unpublished manuscript. One line per term:
//   <ref> t<i> words=<n> exact=<places matched as written>
//   family=<places matched only by a word of the family> total=<both>
// and a total line. Run it on two checkouts (before / after a change of
// taskpane/shared/named-terms.js) to compare.
import fs from "node:fs";
import path from "node:path";
import { termPlaces } from "../../taskpane/shared/named-terms.js";

const run = process.argv[2];
if (!run) {
  console.error("usage: node scripts/e2e/term-family-count.mjs <run folder>");
  process.exit(2);
}
const rev = JSON.parse(fs.readFileSync(path.join(run, "manuscript.revision.json"), "utf8"));
const paragraphs = rev.snapshot?.paragraphs ?? [];
const ids = rev.snapshot?.ids ?? null;
let listed = null;
try {
  const changes = JSON.parse(fs.readFileSync(path.join(run, "changes.json"), "utf8"));
  listed = new Map(changes.map((c) => [c.ref, (c.named_terms ?? []).length]));
} catch {
  // no changes.json: every comment of the review file is counted
}
const sum = { terms: 0, present: 0, exact: 0, family: 0 };
for (const c of rev.comments ?? []) {
  const rows = termPlaces(c, paragraphs, ids);
  if (!rows.length) continue;
  rows.forEach((r, i) => {
    const exact = r.places.filter((p) => !p.form).length;
    const family = r.places.length - exact;
    sum.terms += 1;
    sum.present += r.places.length ? 1 : 0;
    sum.exact += exact;
    sum.family += family;
    console.log(
      `${c.ref} t${i + 1} words=${r.term.split(/\s+/).length} exact=${exact} family=${family} total=${r.places.length}`,
    );
  });
  if (listed && listed.get(c.ref) !== rows.length) {
    console.log(
      `${c.ref}: changes.json lists ${listed.get(c.ref) ?? 0} terms, this code ${rows.length}`,
    );
  }
}
console.log(
  `paragraphs=${paragraphs.length} terms=${sum.terms} still_present=${sum.present} exact_places=${sum.exact} family_places=${sum.family} total_places=${sum.exact + sum.family}`,
);
