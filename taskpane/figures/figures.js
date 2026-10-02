// The figures page: figure changes under reviewer comments.
// Talks to /figures/api/* (daemon/figure-http.mjs) over plain HTTP with the
// bridge token in a header. No WebSocket: a second connection for the same
// document would take the place of the Word pane.
(() => {
  "use strict";

  const API = "/figures/api/";
  let token = "";
  let manuscripts = [];
  let doc = null; // the chosen manuscript, as its pane names it
  let state = null;
  let files = { folder: "figures", files: [] };

  // What the page remembers between two drawings of the list.
  const modeOf = new Map(); // figure id -> "side" | "overlay" | "difference"
  const zoomOf = new Map(); // figure id -> zoom the reader set
  const pictures = new Map(); // "<id>:<side>:<hash>" -> Promise<{ url, width, height } | { error }>
  const forms = new Map(); // comment ref -> the open "Add figure change" form

  const $ = (id) => document.getElementById(id);

  // el("div", { class: "x", onclick: fn }, child, "text", ...)
  function el(tag, props = {}, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(props)) {
      if (v === false || v == null) continue;
      if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
      else if (k === "class") node.className = v;
      else if (k in node && k !== "list") node[k] = v;
      else node.setAttribute(k, v === true ? "" : v);
    }
    for (const c of children.flat()) if (c != null && c !== false) node.append(c);
    return node;
  }

  async function api(name, { query = null, body = null, blob = false } = {}) {
    const url = API + name + (query ? `?${new URLSearchParams(query)}` : "");
    const res = await fetch(url, {
      method: body ? "POST" : "GET",
      headers: {
        "X-Sidewrite-Token": token,
        ...(body ? { "Content-Type": "application/json" } : {}),
      },
      body: body ? JSON.stringify(body) : undefined,
      cache: "no-store",
    });
    if (blob && res.ok) return res.blob();
    let data = null;
    try {
      data = await res.json();
    } catch {
      /* not JSON: the status says it */
    }
    if (!res.ok || !data?.ok) throw new Error(data?.error ?? `The request failed (${res.status}).`);
    return data;
  }

  const baseName = (p) => String(p).split(/[\\/]/).pop();
  const plural = (n, word) => `${n} ${word}${n === 1 ? "" : "s"}`;
  function sizeText(bytes) {
    if (bytes < 1024) return `${bytes} bytes`;
    if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }
  const commentTitle = (c) => {
    const n = c.ref.slice(c.ref.lastIndexOf(".") + 1);
    return `${c.reviewer === 0 ? "Editor" : `Reviewer ${c.reviewer}`}, comment ${n}`;
  };

  function notice(text, bad = false) {
    const box = $("notice");
    box.hidden = !text;
    box.textContent = text || "";
    box.className = bad ? "notice bad" : "notice";
  }

  // ---- loading ---------------------------------------------------------------

  async function loadManuscripts() {
    const data = await api("manuscripts");
    manuscripts = data.manuscripts;
    const pick = $("manuscript");
    pick.replaceChildren(
      ...manuscripts.map((m, i) => el("option", { value: String(i) }, baseName(m.path))),
    );
    const at = manuscripts.findIndex((m) => m.doc === doc);
    doc = manuscripts.length ? manuscripts[Math.max(at, 0)].doc : null;
    pick.value = String(Math.max(at, 0));
    pick.disabled = manuscripts.length < 2;
    pick.title = manuscripts.find((m) => m.doc === doc)?.path ?? "";
  }

  async function loadState() {
    if (!doc) {
      state = null;
      return;
    }
    [state, files] = await Promise.all([
      api("state", { query: { doc } }),
      api("files", { query: { doc } }),
    ]);
  }

  async function reload({ list = true } = {}) {
    try {
      notice("");
      if (list) await loadManuscripts();
      await loadState();
    } catch (e) {
      state = null;
      doc = null;
      notice(e.message, true);
    }
    draw();
  }

  // One request that changes something, then the list again.
  async function act(button, name, body, errorBox) {
    button.disabled = true;
    try {
      const out = await api(name, { body: { doc, ...body } });
      await reload({ list: false });
      return out;
    } catch (e) {
      if (errorBox) errorBox.textContent = e.message;
      else notice(e.message, true);
      button.disabled = false;
      return null;
    }
  }

  // ---- pictures --------------------------------------------------------------

  function picture(c, f, side) {
    const file = f[side].view ?? f[side].source;
    const key = `${f.id}:${side}:${file.sha256}`;
    if (!pictures.has(key)) {
      pictures.set(
        key,
        (async () => {
          try {
            const blob = await api("render", {
              query: { doc, ref: c.ref, id: f.id, side },
              blob: true,
            });
            const url = URL.createObjectURL(blob);
            const img = new Image();
            img.src = url;
            await img.decode();
            return { url, width: img.naturalWidth, height: img.naturalHeight };
          } catch (e) {
            pictures.delete(key);
            return { error: e.message };
          }
        })(),
      );
    }
    return pictures.get(key);
  }

  const SIDE_LABEL = { before: "Old file", after: "New file" };

  function compare(c, f) {
    const sides = f.before ? ["before", "after"] : ["after"];
    const two = sides.length === 2;
    const body = el("div", {}, el("div", { class: "no-picture" }, "Loading the pictures..."));
    const bar = el("div", { class: "compare-bar" });
    const box = el("div", { class: "compare" }, bar, body);

    Promise.all(sides.map((s) => picture(c, f, s))).then((got) => {
      const failed = got.find((p) => p.error);
      if (failed) {
        body.replaceChildren(el("div", { class: "no-picture" }, `No picture. ${failed.error}`));
        return;
      }
      const pic = Object.fromEntries(sides.map((s, i) => [s, got[i]]));
      const maxW = Math.max(...got.map((p) => p.width));
      const maxH = Math.max(...got.map((p) => p.height));
      const mode = () => (two ? (modeOf.get(f.id) ?? "side") : "side");
      // Fit the width of the card until the reader sets a zoom.
      const fit = () => {
        const room = (box.clientWidth || 900) - 4;
        const each = mode() === "side" && two ? (room - 10) / 2 : room;
        return Math.min(1, Math.max(0.1, each / maxW));
      };
      const zoom = () => zoomOf.get(f.id) ?? fit();
      const sized = (side, cls) => {
        const img = el("img", { src: pic[side].url, alt: SIDE_LABEL[side], class: cls });
        img.style.width = `${Math.round(pic[side].width * zoom())}px`;
        img.style.height = `${Math.round(pic[side].height * zoom())}px`;
        return img;
      };

      const zoomText = el("span", {});
      const zoomIn = el("input", {
        type: "range",
        min: "10",
        max: "300",
        step: "5",
        oninput: () => {
          zoomOf.set(f.id, Number(zoomIn.value) / 100);
          paint();
        },
      });
      const split = el("input", {
        type: "range",
        min: "0",
        max: "100",
        value: "50",
        oninput: () => paint(),
      });
      const splitLabel = el("label", {}, "Old", split, "New");
      const modeButtons = [
        ["side", "Side by side"],
        ["overlay", "Overlay"],
        ["difference", "Difference"],
      ].map(([m, text]) =>
        el(
          "button",
          {
            type: "button",
            "data-mode": m,
            onclick: () => {
              modeOf.set(f.id, m);
              paint();
            },
          },
          text,
        ),
      );
      bar.replaceChildren(
        two ? el("div", { class: "modes" }, modeButtons) : null,
        el("label", {}, "Zoom", zoomIn, zoomText),
        two ? splitLabel : null,
      );

      function paint() {
        const m = mode();
        for (const b of modeButtons) b.setAttribute("aria-pressed", String(b.dataset.mode === m));
        splitLabel.hidden = m !== "overlay";
        zoomIn.value = String(Math.round(zoom() * 100));
        zoomText.textContent = `${Math.round(zoom() * 100)}%`;
        if (m === "side") {
          const views = sides.map((s) => el("div", { class: "view" }, sized(s)));
          // Linked scrolling: both pictures show the same place.
          let moving = false;
          for (const v of views)
            v.addEventListener("scroll", () => {
              if (moving) return;
              moving = true;
              for (const o of views)
                if (o !== v) {
                  o.scrollLeft = v.scrollLeft;
                  o.scrollTop = v.scrollTop;
                }
              moving = false;
            });
          body.replaceChildren(
            el(
              "div",
              { class: two ? "panes" : "panes one" },
              sides.map((s, i) =>
                el(
                  "div",
                  {},
                  el(
                    "div",
                    { class: "pane-label" },
                    `${SIDE_LABEL[s]}: ${baseName((f[s].view ?? f[s].source).path)}`,
                  ),
                  views[i],
                ),
              ),
            ),
          );
          return;
        }
        // Both pictures at the same scale, top left corners together.
        const stage = el("div", { class: `stage ${m}` }, sized("before"), sized("after", "top"));
        stage.style.width = `${Math.round(maxW * zoom())}px`;
        stage.style.height = `${Math.round(maxH * zoom())}px`;
        if (m === "overlay") stage.lastChild.style.clipPath = `inset(0 0 0 ${100 - split.value}%)`;
        body.replaceChildren(
          el(
            "div",
            { class: "pane-label" },
            m === "overlay"
              ? "Old file on the left of the slider position, new file on the right"
              : "Black where the two files are the same; anything else is a difference",
          ),
          el("div", { class: "view" }, stage),
        );
      }
      paint();
    });
    return box;
  }

  // ---- one figure change -----------------------------------------------------

  function factsTable(f) {
    const sides = f.before ? ["before", "after"] : ["after"];
    const cell = (side, fn) => {
      const facts = f.facts?.[side];
      return el("td", {}, facts ? fn(facts, f.renders?.[side]) : "Not read");
    };
    const row = (label, fn) => el("tr", {}, el("th", {}, label), ...sides.map((s) => cell(s, fn)));
    const mm = (pt) => Math.round((pt / 72) * 25.4);
    return el(
      "table",
      { class: "facts" },
      el("tr", {}, el("th", {}, ""), ...sides.map((s) => el("th", {}, SIDE_LABEL[s]))),
      el(
        "tr",
        {},
        el("th", {}, "File"),
        ...sides.map((s) => {
          const file = f[s].view ?? f[s].source;
          return el("td", {}, `${baseName(file.path)} (${sizeText(file.bytes)})`);
        }),
      ),
      row("File type", (x) => x.type),
      row("Page size", (x) =>
        x.width_pt == null
          ? "A picture file has no page size"
          : `${mm(x.width_pt)} x ${mm(x.height_pt)} mm`,
      ),
      row("Picture shown", (x, r) =>
        r
          ? `${r.width} x ${r.height} pixels${x.width_pt == null ? " (its own size)" : ` at ${r.dpi} dpi`}`
          : "None",
      ),
      row("Live text", (x) =>
        x.words > 0 ? `Yes, ${plural(x.words, "word")}` : "No (outlines or pixels)",
      ),
      row("Fonts", (x) =>
        x.fonts.length
          ? x.fonts.map((t) => `${t.name} (${t.embedded ? "embedded" : "not embedded"})`).join(", ")
          : "None named in the file",
      ),
    );
  }

  function details(f) {
    const list = el("dl", {});
    const add = (label, value) =>
      list.append(el("dt", {}, label), el("dd", {}, el("code", {}, value)));
    add("Record", f.id);
    for (const side of ["before", "after"]) {
      if (!f[side]) continue;
      add(`${SIDE_LABEL[side]}`, f[side].source.path);
      add(`${SIDE_LABEL[side]} SHA-256`, f[side].source.sha256);
      if (f[side].view) {
        add(`${SIDE_LABEL[side]}, shown and submitted as`, f[side].view.path);
        add(`${SIDE_LABEL[side]}, shown and submitted as, SHA-256`, f[side].view.sha256);
      }
    }
    if (f.adopted) {
      add("Copy for submission", f.adopted.path);
      add("Copy for submission SHA-256", f.adopted.sha256);
      add("Adopted", f.adopted.at);
    }
    add("Recorded", f.created_at);
    return el("details", {}, el("summary", {}, "Details"), list);
  }

  function requestChecks(c, ticked) {
    const boxes = c.requirements.map((q) => {
      const input = el("input", { type: "checkbox", value: q.id, checked: ticked.includes(q.id) });
      return [
        input,
        el(
          "label",
          { class: "check" },
          input,
          el(
            "span",
            {},
            q.text,
            q.outside ? el("span", { class: "badge outside" }, "Outside the Word file") : null,
          ),
        ),
      ];
    });
    return {
      nodes: boxes.length
        ? boxes.map((b) => b[1])
        : [
            el(
              "div",
              { class: "hint" },
              "This comment has no listed requests. List them on the Review tab in Word to tick what the figure answers.",
            ),
          ],
      value: () => boxes.filter((b) => b[0].checked).map((b) => b[0].value),
    };
  }

  function editForm(c, f, close) {
    const label = el("input", { type: "text", value: f.figure });
    const summary = el("input", { type: "text", value: f.summary });
    const checks = requestChecks(c, f.requirement_ids);
    const error = el("p", { class: "error" });
    const save = el(
      "button",
      {
        type: "button",
        class: "primary",
        onclick: () =>
          act(
            save,
            "update",
            {
              ref: c.ref,
              id: f.id,
              patch: {
                figure: label.value,
                summary: summary.value,
                requirement_ids: checks.value(),
              },
            },
            error,
          ),
      },
      "Save",
    );
    return el(
      "div",
      { class: "form" },
      el("div", { class: "row" }, el("label", {}, "Figure label"), label),
      el(
        "div",
        { class: "row" },
        el("div", { class: "label" }, "Requests it answers"),
        checks.nodes,
      ),
      el("div", { class: "row" }, el("label", {}, "What changed, in one sentence"), summary),
      el(
        "div",
        { class: "actions" },
        save,
        el("button", { type: "button", onclick: close }, "Cancel"),
      ),
      error,
    );
  }

  function figureCard(c, f) {
    const answers = c.requirements.filter((q) => f.requirement_ids.includes(q.id));
    const stateText = {
      proposed: "Waiting for your decision",
      adopted: "Adopted",
      rejected: "Rejected",
    };
    const card = el(
      "div",
      { class: "figure", "data-figure": f.id },
      el(
        "div",
        { class: "figure-head" },
        el("h3", {}, f.figure, f.kind === "new" ? " (new figure)" : " (revised figure)"),
        el("span", { class: `badge ${f.state}` }, stateText[f.state]),
      ),
      f.summary ? el("p", {}, f.summary) : null,
      el(
        "p",
        { class: "answers" },
        answers.length
          ? `Answers: ${answers.map((q) => q.text).join("; ")}`
          : "Answers no listed request",
      ),
      f.problems.length
        ? el(
            "ul",
            { class: "problems" },
            f.problems.map((p) => el("li", {}, p)),
          )
        : null,
      compare(c, f),
      factsTable(f),
      f.findings.length
        ? el(
            "ul",
            { class: "findings" },
            f.findings.map((p) => el("li", {}, p)),
          )
        : null,
    );
    const actions = el("div", { class: "actions" });
    const button = (text, name, cls) => {
      const b = el(
        "button",
        { type: "button", class: cls, onclick: () => act(b, name, { ref: c.ref, id: f.id }) },
        text,
      );
      return b;
    };
    if (f.state === "proposed") {
      const edit = el(
        "button",
        {
          type: "button",
          class: "quiet",
          onclick: () => {
            const form = editForm(c, f, () => form.replaceWith(actions));
            actions.replaceWith(form);
          },
        },
        "Edit the text",
      );
      actions.append(button("Adopt", "adopt", "primary"), button("Reject", "reject"), edit);
    } else if (f.state === "adopted") {
      actions.append(
        el("span", { class: "adopted-line" }, `Copy for submission: ${f.adopted.path}`),
        button("Reopen", "reopen"),
      );
    } else actions.append(button("Reopen", "reopen"));
    card.append(actions, details(f));
    return card;
  }

  // ---- the form "Add figure change" --------------------------------------------

  const OTHER = "::other";
  function filePicker() {
    const select = el(
      "select",
      {},
      el("option", { value: "" }, "Choose a file"),
      files.files.map((x) =>
        el("option", { value: x.path }, `${x.path}${x.submitted ? " (copy for submission)" : ""}`),
      ),
      el("option", { value: OTHER }, "Another file: paste its path"),
    );
    const path = el("input", {
      type: "text",
      hidden: true,
      placeholder: "Full path of the file",
    });
    const listeners = [];
    select.addEventListener("change", () => {
      path.hidden = select.value !== OTHER;
      listeners.forEach((fn) => fn());
    });
    path.addEventListener("input", () => listeners.forEach((fn) => fn()));
    return {
      nodes: [select, path],
      value: () => (select.value === OTHER ? path.value.trim() : select.value),
      onchange: (fn) => listeners.push(fn),
    };
  }

  function addForm(c) {
    const kind = (value, text, checked) =>
      el(
        "label",
        { class: "choice" },
        el("input", { type: "radio", name: `kind-${c.ref}`, value, checked }),
        text,
      );
    const revised = kind("revised", "A figure that was revised", true);
    const fresh = kind("new", "A new figure", false);
    const label = el("input", { type: "text", placeholder: "Figure 2" });
    const checks = requestChecks(
      c,
      c.requirements.filter((q) => q.outside).map((q) => q.id),
    );
    const before = filePicker();
    const after = filePicker();
    const beforeView = filePicker();
    const afterView = filePicker();
    const summary = el("input", {
      type: "text",
      placeholder: "The control curve was added and the time axis now gives hours.",
    });
    const error = el("p", { class: "error" });

    const viewRow = (picker) =>
      el(
        "div",
        { class: "row", hidden: true },
        el("label", {}, "PDF or TIFF saved from the .ai file"),
        picker.nodes,
        el(
          "div",
          { class: "hint" },
          "Optional. This file is shown here and copied for submission; the .ai file stays the source.",
        ),
      );
    const beforeViewRow = viewRow(beforeView);
    const afterViewRow = viewRow(afterView);
    const isAi = (p) => /\.ai$/i.test(p.value());
    before.onchange(() => (beforeViewRow.hidden = !isAi(before)));
    after.onchange(() => (afterViewRow.hidden = !isAi(after)));

    const beforeCol = el(
      "div",
      {},
      el("div", { class: "row" }, el("label", {}, "Old file"), before.nodes),
      beforeViewRow,
    );
    const isNew = () => fresh.firstChild.checked;
    const kindChanged = () => (beforeCol.hidden = isNew());
    revised.firstChild.addEventListener("change", kindChanged);
    fresh.firstChild.addEventListener("change", kindChanged);

    const save = el(
      "button",
      {
        type: "button",
        class: "primary",
        onclick: async () => {
          error.textContent = "";
          const out = await act(
            save,
            "propose",
            {
              ref: c.ref,
              kind: isNew() ? "new" : "revised",
              figure: label.value,
              requirement_ids: checks.value(),
              summary: summary.value,
              ...(isNew()
                ? {}
                : {
                    before: before.value(),
                    ...(isAi(before) ? { before_view: beforeView.value() } : {}),
                  }),
              after: after.value(),
              ...(isAi(after) ? { after_view: afterView.value() } : {}),
            },
            error,
          );
          if (out) {
            forms.delete(c.ref);
            draw();
          }
        },
      },
      "Record the figure change",
    );
    return el(
      "div",
      { class: "form" },
      el("h3", {}, "Add figure change"),
      el("div", { class: "row" }, revised, fresh),
      el("div", { class: "row" }, el("label", {}, "Figure label in the manuscript"), label),
      el(
        "div",
        { class: "row" },
        el("div", { class: "label" }, "Requests it answers"),
        checks.nodes,
      ),
      el(
        "div",
        { class: "row pair" },
        beforeCol,
        el(
          "div",
          {},
          el("div", { class: "row" }, el("label", {}, "New file"), after.nodes),
          afterViewRow,
        ),
      ),
      el(
        "div",
        { class: "hint row" },
        files.files.length
          ? `The lists show the figure files in the folder "${files.folder}" next to the manuscript. Neither file is changed.`
          : `No figure files were found in the folder "${files.folder}" next to the manuscript. Put the figure files there and reload, or paste a path.`,
      ),
      el("div", { class: "row" }, el("label", {}, "What changed, in one sentence"), summary),
      el(
        "div",
        { class: "actions" },
        save,
        el(
          "button",
          {
            type: "button",
            onclick: () => {
              forms.delete(c.ref);
              draw();
            },
          },
          "Cancel",
        ),
      ),
      error,
    );
  }

  // ---- the list --------------------------------------------------------------

  function commentCard(c) {
    const adopted = new Set(
      c.figures.filter((f) => f.state === "adopted").flatMap((f) => f.requirement_ids),
    );
    const proposed = new Set(
      c.figures.filter((f) => f.state === "proposed").flatMap((f) => f.requirement_ids),
    );
    const outside = c.requirements.some((q) => q.outside);
    return el(
      "section",
      { class: outside ? "comment outside" : "comment", "data-ref": c.ref },
      el("h2", {}, commentTitle(c), c.section ? ` (${c.section})` : ""),
      el("p", { class: "comment-text" }, c.text),
      c.requirements.length
        ? el(
            "ul",
            { class: "requests" },
            c.requirements.map((q) =>
              el(
                "li",
                {},
                q.text,
                q.outside ? el("span", { class: "badge outside" }, "Outside the Word file") : null,
                adopted.has(q.id)
                  ? el("span", { class: "badge adopted" }, "Answered by an adopted figure")
                  : proposed.has(q.id)
                    ? el("span", { class: "badge proposed" }, "Figure waits for a decision")
                    : null,
                q.outside && q.why ? el("span", { class: "why" }, q.why) : null,
              ),
            ),
          )
        : null,
      c.figures.map((f) => figureCard(c, f)),
      forms.get(c.ref) ??
        el(
          "button",
          {
            type: "button",
            "data-add": c.ref,
            onclick: () => {
              forms.set(c.ref, addForm(c));
              draw();
            },
          },
          "Add figure change",
        ),
    );
  }

  function draw() {
    const list = $("comments");
    $("empty").hidden = !!doc;
    if (!state) {
      list.replaceChildren();
      return;
    }
    if (!state.renderer.available) notice(state.renderer.reason);
    if (!state.comments.length) {
      list.replaceChildren(
        el(
          "div",
          { class: "empty" },
          el("h2", {}, "No reviewer comments yet"),
          el(
            "p",
            {},
            "Paste the review on the Review tab of the Sidewrite pane in Word, then reload.",
          ),
        ),
      );
      return;
    }
    const first = state.comments.filter(
      (c) => c.requirements.some((q) => q.outside) || c.figures.length,
    );
    const rest = state.comments.filter((c) => !first.includes(c));
    list.replaceChildren(
      first.length
        ? el("div", { class: "group" }, "Requests outside the Word file, and figure changes")
        : null,
      ...first.map(commentCard),
      rest.length
        ? el("div", { class: "group" }, first.length ? "Other comments" : "Comments")
        : null,
      ...rest.map(commentCard),
    );
  }

  $("reload").addEventListener("click", () => reload());
  $("manuscript").addEventListener("change", (e) => {
    doc = manuscripts[Number(e.target.value)]?.doc ?? null;
    forms.clear();
    reload({ list: false });
  });

  (async () => {
    try {
      token = (await (await fetch("/bridge-token", { cache: "no-store" })).text()).trim();
    } catch {
      notice("The Sidewrite daemon did not answer. Start it, then reload.", true);
      return;
    }
    await reload();
  })();
})();
