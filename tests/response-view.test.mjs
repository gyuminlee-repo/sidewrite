import { test } from "node:test";
import assert from "node:assert/strict";
import { bytesToBase64, pdfTooLarge, PDF_FRAME_LIMIT } from "../taskpane/shared/response-view.js";

test("base64 of bytes in chunks equals Buffer's", () => {
  const b = Uint8Array.from({ length: 20000 }, (_, i) => i % 251);
  assert.equal(bytesToBase64(b), Buffer.from(b).toString("base64"));
  assert.equal(bytesToBase64(new Uint8Array(0)), "");
  // A length that is no multiple of the chunk size, and one byte.
  const odd = Uint8Array.from({ length: 8193 }, (_, i) => 255 - (i % 256));
  assert.equal(bytesToBase64(odd), Buffer.from(odd).toString("base64"));
  assert.equal(bytesToBase64(Uint8Array.of(0)), "AA==");
});

test("a PDF that cannot fit one WebSocket frame is refused before sending", () => {
  assert.equal(PDF_FRAME_LIMIT, 100 * 1024 * 1024);
  assert.equal(pdfTooLarge(10), false);
  assert.equal(pdfTooLarge(100 * 1024 * 1024), true);
  assert.deepEqual(
    [pdfTooLarge(PDF_FRAME_LIMIT - 4096), pdfTooLarge(PDF_FRAME_LIMIT - 4095)],
    [false, true],
  );
});
