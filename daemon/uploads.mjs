// Files the task pane sends to the daemon in chunks (the baseline for a
// marked copy). One message per chunk: { upload_id, seq, last, data_b64 }.
// The chunks are held in memory and written to <dir>/<upload_id>.docx when
// the last one arrives; the answer to the last chunk carries the path.
//
// A chunk out of order, bad base64 or a file over the size cap ends that
// upload and nothing is written. An upload id is a plain name (letters,
// digits, "-" and "_"), so a file cannot be written outside the folder.
// Only a few uploads are held at a time.
//
// A finished upload keeps its record. When its chunks come again (the answer
// to the last chunk was lost), the same bytes are answered like the first
// time and other bytes are refused; the file and the record are never
// replaced.
//
// Not wired yet. The caller passes a folder of its own (a temporary one)
// and answers the pane's file_upload message with accept().

import {
  mkdir as fsMkdir,
  readFile as fsReadFile,
  writeFile as fsWriteFile,
} from "node:fs/promises";
import { join } from "node:path";
import { createAssembler, DOC_MAX_BYTES } from "./doc-bytes.mjs";

const ID = /^[A-Za-z0-9_-]{1,64}$/;

export function createUploads({
  dir,
  maxBytes = DOC_MAX_BYTES,
  maxOpen = 4,
  writeFile = fsWriteFile,
  readFile = fsReadFile,
  mkdir = fsMkdir,
} = {}) {
  const open = new Map();
  const done = new Map();
  return {
    async accept(message) {
      const id = message?.upload_id;
      if (typeof id !== "string" || !ID.test(id)) return { ok: false, error: "bad upload id" };
      if (!open.has(id)) {
        if (open.size >= maxOpen) return { ok: false, error: "too many uploads at once" };
        open.set(id, createAssembler({ maxBytes }));
      }
      const r = open.get(id).add(message);
      if (!r.ok) {
        open.delete(id);
        return r;
      }
      if (!r.done) return { ok: true, received: r.received };
      open.delete(id);
      const path = join(dir, `${id}.docx`);
      try {
        await mkdir(dir, { recursive: true, mode: 0o700 });
        // "wx": the file is created, never replaced.
        await writeFile(path, r.bytes, { mode: 0o600, flag: "wx" });
      } catch (e) {
        if (e?.code === "EEXIST") {
          // The same file sent again is the same upload; anything else is refused.
          const there = await readFile(path).catch(() => null);
          if (there && Buffer.from(there).equals(r.bytes)) {
            done.set(id, path);
            return { ok: true, received: r.received, path };
          }
          return {
            ok: false,
            error: `A file for the upload id ${id} already exists with different content.`,
          };
        }
        return { ok: false, error: `The upload could not be saved: ${String(e?.message ?? e)}` };
      }
      done.set(id, path);
      return { ok: true, received: r.received, path };
    },
    // The path of a finished upload, or null.
    pathOf: (id) => done.get(id) ?? null,
    drop(id) {
      open.delete(id);
      done.delete(id);
    },
  };
}
