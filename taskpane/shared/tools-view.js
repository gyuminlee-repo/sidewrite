// Wording of the Tools tab (tools-pane.js). Pure functions: no DOM, no Word.

// Which file the copies are made from.
export const docLine = (doc) =>
  doc
    ? `Saved file: ${doc.doc_name}. Profile: ${doc.profile_name}. Copies are made from the saved file, so save before you make one.`
    : "Save the document to a local folder to use the tools.";
