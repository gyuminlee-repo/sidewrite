// Pure helpers of the Response tab: no DOM, no Office.js, so they run under
// node --test.

// The daemon's WebSocket server keeps the `ws` default of 100 MiB per frame
// (daemon/bridge.mjs does not change it). A PDF goes to the daemon as base64
// inside one JSON frame.
export const PDF_FRAME_LIMIT = 100 * 1024 * 1024;
const ENVELOPE = 4096; // room for the rest of the JSON message

// base64 of a Uint8Array. String.fromCharCode takes a limited number of
// arguments, so the bytes go in chunks.
export function bytesToBase64(bytes) {
  let bin = "";
  const CHUNK = 8192;
  for (let i = 0; i < bytes.length; i += CHUNK)
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  return btoa(bin);
}

// Would a PDF of this base64 length not fit one frame? Checked before sending:
// a frame that is too large is dropped without an answer.
export const pdfTooLarge = (base64Length) => base64Length + ENVELOPE > PDF_FRAME_LIMIT;
