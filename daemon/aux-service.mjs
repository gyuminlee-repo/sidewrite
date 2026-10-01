// Tools tab: task-pane requests beside the review and the response letter.
//
// Requests (bridge.mjs extraHandlers): { type, request_id, ...payload } and the
// reply { type: "<type>_result", request_id, ok, ... }.
//   tools_state
// Jobs (AUX_JOBS; they call a model or a Python tool): accepted at once, the
// result follows as { type: "tools_event", event: "job_done", job_id, result }.
//
// Nothing here writes to the open document. A copy (marked, anonymized) is a
// new file next to the manuscript, made from the file as it was last saved.
//
// Handler rules (as in response-service.mjs): the document of a pane is read
// once per request (at(key)); a failure is an error reply, never a silent
// default.

import { basename, dirname } from "node:path";
import { localDocPath, revisionPathFor, loadRevision } from "./revision-store.mjs";
import { wrapHandlers } from "./handler-wrap.mjs";
import { effectiveProfile } from "./journal-profile.mjs";
import { SAVE_FIRST } from "./doc-bytes.mjs";

export const AUX_JOBS = new Set([]);

export function createAuxService({
  getContext,
  // Default language of the settings file: () => string | null.
  langDefault = null,
  sendEvent = () => {},
  log = console.log,
  env = process.env,
}) {
  // The saved file of a pane, read from its live context once.
  function at(key) {
    const doc = localDocPath(getContext(key)?.activeDoc, env);
    const need = () => {
      if (!doc) throw new Error(SAVE_FIRST);
      return doc;
    };
    return {
      key,
      get doc() {
        return need();
      },
      get path() {
        return revisionPathFor(need());
      },
    };
  }
  const load = (d) => loadRevision(d.path, d.doc);
  const profileOf = (d, rev) =>
    effectiveProfile(rev, {
      docDir: dirname(d.doc),
      env,
      settingsDefault: langDefault?.() ?? null,
    });

  const impl = {
    // Which saved file the copies are made from, and the profile in force.
    async tools_state(_msg, key) {
      const d = at(key);
      const { name, profile } = await profileOf(d, await load(d));
      return { doc_name: basename(d.doc), profile_name: name, policy_packs: profile.policy_packs };
    },
  };

  return {
    handlers: wrapHandlers(impl, {
      jobTypes: AUX_JOBS,
      sendEvent,
      eventType: "tools_event",
      log,
    }),
  };
}
