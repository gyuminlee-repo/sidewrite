// The fake task panes of the review tests read a document that holds no
// field (no citation, no reference list, no hyperlink), and most of them
// give no reviewed text. The real pane says from the XML of the paragraphs
// whether the document holds a field (collectReviewSnapshot has_fields); a
// snapshot that does not say is taken as "not known", and then a paragraph
// without reviewed text is not edited (field-text.js fieldsUnread). This
// puts that answer into the request of a fake pane that left it out.
export function plainPane(payload) {
  const s = payload?.snapshot;
  if (!s || typeof s !== "object" || "has_fields" in s) return payload;
  return { ...payload, snapshot: { ...s, has_fields: false } };
}
