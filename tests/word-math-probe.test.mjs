// The reading of document.xml that the equation probe prints
// (scripts/e2e/word-math-probe-lib.mjs), on XML written by hand.

import { test } from "node:test";
import assert from "node:assert/strict";
import { markRelations, mathSummary, parseXml } from "../scripts/e2e/word-math-probe-lib.mjs";

const doc = (...paragraphs) =>
  `<?xml version="1.0"?><!-- a comment --><w:document xmlns:w="w" xmlns:m="m"><w:body>${paragraphs.join("")}<w:sectPr/></w:body></w:document>`;
const INS = '<w:ins w:id="1" w:author="A &amp; B" w:date="2031-01-02T03:04:05Z">';

test("the XML is read into a tree: nesting, empty elements, attributes with a slash, text", () => {
  const root = parseXml("<a x=\"1/2\"><b/><c y='>'>one<d/>two</c></a>");
  const a = root.children[0];
  assert.equal(a.attrs, 'x="1/2"');
  assert.deepEqual(
    a.children.map((n) => n.tag),
    ["b", "c"],
  );
  assert.equal(a.children[1].text, "onetwo");
  assert.equal(a.children[1].children[0].parent, a.children[1]);
});

test("a paragraph without an equation and without a change is left out", () => {
  assert.deepEqual(mathSummary(doc("<w:p><w:r><w:t>Plain.</w:t></w:r></w:p>")), []);
  assert.deepEqual(mathSummary("<x/>"), ["no w:body in this XML"]);
});

test("an equation inside a change mark is shown under it, run by run", () => {
  const p =
    "<w:p><w:r><w:t>Before </w:t></w:r>" +
    `${INS}<m:oMath><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>x</m:t></m:r></m:oMath></w:ins>` +
    "</w:p>";
  assert.deepEqual(mathSummary(doc("<w:p><w:r><w:t>Plain.</w:t></w:r></w:p>", p)), [
    'paragraph 2  mark: -  text: "Before x"',
    '  w:r "Before " [w:t]',
    "  w:ins",
    "    m:oMath",
    '      m:r "x" [m:t] m:rPr',
  ]);
});

test("a change mark inside an equation, in run properties and in control properties, is named", () => {
  const p =
    '<w:p><w:pPr><w:rPr><w:del w:id="3"/></w:rPr></w:pPr><m:oMath>' +
    '<m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:rFonts/><w:ins w:id="4"/></w:rPr><w:del w:id="5"/></m:ctrlPr></m:sSupPr>' +
    '<m:e><m:r><w:rPr><w:del w:id="6"/></w:rPr><m:t>y</m:t></m:r></m:e></m:sSup>' +
    `${INS}<m:r><m:t>z</m:t></m:r></w:ins></m:oMath></w:p>`;
  assert.deepEqual(mathSummary(doc(p)), [
    'paragraph 1  mark: w:del  text: "yz"',
    "  m:oMath",
    "    m:sSup",
    "      m:sSupPr",
    "        m:ctrlPr[w:rPr[w:ins], w:del]",
    "      m:e",
    '        m:r "y" [m:t] w:rPr[w:del]',
    "    w:ins",
    '      m:r "z" [m:t]',
  ]);
  assert.deepEqual(markRelations(doc(p)), [
    "  1 x m:ctrlPr > w:rPr > w:ins (a flag)",
    "  1 x m:r > w:rPr > w:del (a flag)",
    "  1 x m:sSupPr > m:ctrlPr > w:del (a flag)",
    "  1 x w:p > m:oMath > w:ins > {m:r}",
    "  1 x w:pPr > w:rPr > w:del (a flag)",
  ]);
});

test("deleted text of a run is its text, and what the run holds is listed once", () => {
  const p =
    '<w:p><w:del w:id="1"><w:r><w:rPr><w:b/></w:rPr><w:delText>gone </w:delText><w:delText>now</w:delText></w:r></w:del></w:p>';
  assert.deepEqual(mathSummary(doc(p)), [
    'paragraph 1  mark: -  text: "gone now"',
    "  w:del",
    '    w:r "gone now" [w:delText] w:rPr[]',
  ]);
  assert.deepEqual(markRelations(doc(p)), ["  1 x w:body > w:p > w:del > {w:r}"]);
});
