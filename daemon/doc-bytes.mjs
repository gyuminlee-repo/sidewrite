// The bytes of a docx as they travel between the task pane and the daemon:
// base64 chunks { seq, last, data_b64 }, numbered from 0, the last one
// flagged. Pure assembly; the pane call and the file read are injected.
//
// getDocBytes asks the pane first (it has the unsaved edits). When the pane
// cannot give the file, the saved file is read from disk and the result
// says that only what was saved is included.
//
// Not wired yet. The caller passes fromPane (a call of the pane's file
// export that resolves to { ok, chunks }), readFile and the local path of
// the document.

export const DOC_MAX_BYTES = 100 * 1024 * 1024;
export const SAVE_FIRST = "Save the document to a local folder first.";
export const DISK_NOTICE =
  "The file was read from disk: only what was saved is included, not unsaved edits.";

const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;

export function toChunks(bytes, chunkBytes = 1024 * 1024) {
  const data = Buffer.from(bytes);
  const out = [];
  for (let at = 0, seq = 0; at < data.length || seq === 0; at += chunkBytes, seq++) {
    out.push({
      seq,
      last: at + chunkBytes >= data.length,
      data_b64: data.subarray(at, at + chunkBytes).toString("base64"),
    });
  }
  return out;
}

// add(chunk) -> { ok, done, received, bytes? } or { ok: false, error }. A
// refused chunk ends the transfer: nothing more is taken.
export function createAssembler({ maxBytes = DOC_MAX_BYTES } = {}) {
  let parts = [];
  let received = 0;
  let next = 0;
  let closed = null;
  const refuse = (error) => {
    parts = [];
    closed = error;
    return { ok: false, error };
  };
  return {
    add(chunk) {
      if (closed) return { ok: false, error: closed };
      if (chunk === null || typeof chunk !== "object") return refuse("not a chunk");
      if (chunk.seq !== next) return refuse(`chunk ${chunk.seq} came when ${next} was expected`);
      const text = chunk.data_b64;
      if (typeof text !== "string" || !BASE64.test(text)) return refuse("chunk data is not base64");
      const size = (text.length / 4) * 3 - (text.endsWith("==") ? 2 : text.endsWith("=") ? 1 : 0);
      if (received + size > maxBytes) return refuse(`the file is larger than ${maxBytes} bytes`);
      parts.push(Buffer.from(text, "base64"));
      received += size;
      next++;
      if (!chunk.last) return { ok: true, done: false, received };
      const bytes = Buffer.concat(parts);
      parts = [];
      closed = "the transfer is already complete";
      return { ok: true, done: true, received, bytes };
    },
  };
}

function assemble(chunks, maxBytes) {
  if (!Array.isArray(chunks) || !chunks.length) return { ok: false, error: "no chunks" };
  const asm = createAssembler({ maxBytes });
  let r;
  for (const c of chunks) {
    r = asm.add(c);
    if (!r.ok) return r;
  }
  return r.done ? r : { ok: false, error: "the last chunk is missing" };
}

export async function getDocBytes({ fromPane, readFile, docPath, maxBytes = DOC_MAX_BYTES } = {}) {
  if (typeof fromPane === "function") {
    try {
      const pane = await fromPane();
      const r = pane?.ok ? assemble(pane.chunks, maxBytes) : null;
      if (r?.ok) return { ok: true, bytes: r.bytes, source: "pane" };
    } catch {
      // The pane could not give the file; the saved file is read instead.
    }
  }
  if (!docPath) return { ok: false, error: SAVE_FIRST };
  try {
    const bytes = Buffer.from(await readFile(docPath));
    if (bytes.length > maxBytes) {
      return { ok: false, error: `the file is larger than ${maxBytes} bytes` };
    }
    return { ok: true, bytes, source: "disk", notice: DISK_NOTICE };
  } catch (e) {
    return { ok: false, error: `The saved file could not be read: ${String(e?.message ?? e)}` };
  }
}
