// Going to a place in the pane: a tab and one element of it.
//
// createGoTo(setActiveTab, doc) returns goTo(tabName, elementId): the tab is
// shown, a fold (<details>) the element is in is opened, and the element is
// brought into view with the focus on it. An element that takes no focus by
// itself (a list, a section) is made to take it. An id the page does not have
// still shows the tab.
//
// Used by the submission checklist (review-pane.js): the place where a row is
// fixed. setActiveTab is module-local in taskpane.js, so it is handed in.

export function createGoTo(setActiveTab, doc = globalThis.document) {
  return function goTo(tabName, elementId) {
    setActiveTab(tabName);
    const target = doc.getElementById(elementId);
    if (!target) return false;
    for (let d = target.closest?.("details"); d; d = d.parentElement?.closest?.("details"))
      d.open = true;
    if (target.tabIndex < 0) target.setAttribute("tabindex", "-1");
    target.scrollIntoView?.({ block: "center" });
    target.focus?.();
    return true;
  };
}
