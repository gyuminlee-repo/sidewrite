// Wording of the Tools tab (tools-pane.js). Pure functions: no DOM, no Word.

// What a selection request was about, as the head of its answer.
const QUOTE_CHARS = 80;
export function selectionHead(kind, text, r) {
  const flat = String(text).replace(/\s+/g, " ").trim();
  const quote = flat.length > QUOTE_CHARS ? `${flat.slice(0, QUOTE_CHARS)}...` : flat;
  const what = kind === "translate" ? `Translation into ${r.lang}` : "Answer";
  return `${what} (${r.model}) for: "${quote}"`;
}

// Which file the copies are made from.
export const docLine = (doc) =>
  doc
    ? `Saved file: ${doc.doc_name}. Profile: ${doc.profile_name}. Copies are made from the saved file, so save before you make one.`
    : "Save the document to a local folder to use the tools.";
