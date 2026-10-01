// Turn the functions of a service into bridge handlers.
//
// impl: { <type>: async (msg, key, host) => result object }
// A plain request is answered with { type: "<type>_result", request_id, ok,
// ...result } or { ok: false, error }.
// A job (its type is in jobTypes) is answered at once with { accepted: true,
// job_id }; it runs on, and its result goes out as an event
//   { type: eventType, event: "job_done", job_id, result: { ok, ... } }
// through sendEvent(key, event). The task pane waits 10 s for a reply, so
// anything that calls a model or a Python tool is a job.
//
// eventType keeps the jobs of one service apart from those of another
// ("review_event", "response_event").

import { randomUUID } from "node:crypto";

export function wrapHandlers(
  impl,
  {
    jobTypes = new Set(),
    sendEvent = () => {},
    eventType = "review_event",
    log = console.log,
  } = {},
) {
  const tag = `[${eventType.replace(/_event$/, "")}]`;
  const handlers = {};
  for (const [type, fn] of Object.entries(impl)) {
    if (jobTypes.has(type)) {
      // Returns once accepted; the job runs on and ends in job_done.
      handlers[type] = async (msg, reply, key, host) => {
        const job_id = randomUUID();
        reply({
          type: `${type}_result`,
          request_id: msg.request_id,
          ok: true,
          accepted: true,
          job_id,
        });
        const done = (result) =>
          sendEvent(key, { type: eventType, event: "job_done", job_id, result });
        // The job starts before this handler returns, as it always did; a
        // function that throws before its first await is a failed job too.
        let job;
        try {
          job = Promise.resolve(fn(msg, key, host));
        } catch (e) {
          job = Promise.reject(e);
        }
        job
          .then(
            (out) => done({ ok: true, ...out }),
            (e) => done({ ok: false, error: e?.message ?? String(e) }),
          )
          // Nobody awaits the job: a failed send must not end the daemon.
          .catch((e) => log(`${tag} job_done not delivered: ${e?.message ?? e}`));
      };
      continue;
    }
    handlers[type] = async (msg, reply, key, host) => {
      try {
        const out = await fn(msg, key, host);
        reply({ type: `${type}_result`, request_id: msg.request_id, ok: true, ...out });
      } catch (e) {
        reply({
          type: `${type}_result`,
          request_id: msg.request_id,
          ok: false,
          error: e?.message ?? String(e),
        });
      }
    };
  }
  return handlers;
}
