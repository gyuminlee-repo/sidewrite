// One time limit for a whole exchange with a server: the request, the
// headers and the body. The signal goes into the fetch call; fetch then also
// ends the reading of the body when the time is up, so a server that sends
// its headers and then nothing cannot hold the caller for ever.

export const FETCH_TIMEOUT_MS = 30_000;

export const deadline = (ms = FETCH_TIMEOUT_MS) => AbortSignal.timeout(ms);
