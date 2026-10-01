// Tools tab (Word only). DOM glue: every result is computed by the daemon
// (daemon/aux-service.mjs); the wording of the screen is in pure helpers
// (tools-view.js). The pane reads Word (the selection, the snapshot, the
// headings, the formatting) and never writes to it: a copy is a file of its
// own, made by the daemon from the saved file.
// Text from the daemon or the document is inserted with textContent only.

import { docLine } from "./tools-view.js";

const CONNECTION_LOST =
  "The connection to the daemon was lost before the result came. Run it again.";

export function initToolsPane({
  sendRequest,
  word,
  // () => the request that confirms the stored providers while it is under
  // way (providers-pane.js whenRestored), else null.
  providersReady = () => null,
}) {
  const root = document.getElementById("tools-root");
  if (!root) return { onEvent() {}, reload: async () => {} }; // Excel pane: no Tools tab
  const $ = (id) => document.getElementById(id);
  const state = {
    doc: null, // the tools_state answer, or null when the document is not saved
    busy: false,
    asking: 0, // job requests sent whose { accepted, job_id } has not come yet
    jobs: new Map(), // job_id -> { resolve, reject } of an askJob waiting for job_done
    jobsDone: new Map(), // job_id -> result of a job_done that came before its answer
  };

  function el(tag, props = {}, ...kids) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(props)) {
      if (k === "text") n.textContent = v;
      else n.setAttribute(k, v);
    }
    for (const k of kids) if (k) n.append(k);
    return n;
  }
  function say(text, bad = false) {
    const m = $("tools-message");
    m.textContent = text;
    m.classList.toggle("form-error", bad);
    m.hidden = !text;
  }
  async function ask(type, payload = {}) {
    const r = await sendRequest(type, payload);
    if (!r.ok) throw new Error(r.error || `${type} failed`);
    return r;
  }
  // A job answers { accepted, job_id } at once; its result comes in a
  // job_done tools_event, possibly before the answer.
  async function askJob(type, payload) {
    state.asking += 1;
    let job_id;
    try {
      ({ job_id } = await ask(type, payload));
    } finally {
      state.asking -= 1;
    }
    let result = state.jobsDone.get(job_id);
    state.jobsDone.delete(job_id);
    result ??= await new Promise((resolve, reject) => state.jobs.set(job_id, { resolve, reject }));
    return result;
  }
  function jobDone(msg) {
    const w = state.jobs.get(msg.job_id);
    if (w) {
      state.jobs.delete(msg.job_id);
      return w.resolve(msg.result);
    }
    // A job of the connection before this one is not waited for: its result
    // is a file next to the manuscript or nothing at all.
    if (state.asking) state.jobsDone.set(msg.job_id, msg.result);
  }
  // A new connection never brings the answers of the old one: every wait
  // ends now, with an error.
  function dropJobs() {
    for (const w of state.jobs.values()) w.reject(new Error(CONNECTION_LOST));
    state.jobs.clear();
    state.jobsDone.clear();
  }

  // Every action of the tab goes through guard(): one at a time, in the order
  // pressed, and only once the daemon has confirmed the stored provider.
  let chain = Promise.resolve();
  const buttons = [];
  function setBusy(busy) {
    state.busy = busy;
    for (const b of buttons) b.disabled = busy;
  }
  async function run(fn) {
    try {
      const confirming = providersReady();
      if (confirming) {
        say("Waiting for the daemon to confirm the stored provider choice...");
        await confirming;
      }
      say("");
      setBusy(true);
      await fn();
    } catch (e) {
      say(e.message, true);
    } finally {
      setBusy(false);
    }
  }
  const guard = (fn) => () => (chain = chain.then(() => run(fn)));
  function onClick(id, fn) {
    const b = $(id);
    b.addEventListener("click", guard(fn));
    buttons.push(b);
  }

  function renderDoc() {
    $("tools-doc").textContent = docLine(state.doc);
  }

  // ---- sections -------------------------------------------------------------

  renderDoc(); // loaded on "welcome" (taskpane.js calls reload once connected)

  return {
    onEvent(msg) {
      if (msg.event === "job_done") jobDone(msg);
    },
    // A new connection (welcome): the waits of the old one end with an error.
    async reload() {
      dropJobs();
      try {
        state.doc = await ask("tools_state");
      } catch (e) {
        state.doc = null;
        say(e.message, true);
      }
      renderDoc();
    },
  };
}
