// Shipped Draftspect manifests immediately before the rename (PR #15 base:
// eb0f2cfc52b055d6ba6129895ad165b32da09f76). Match content, not just a filename
// or GUID: customized manifests at the old paths belong to the user.
import { createHash } from "node:crypto";

export const LEGACY_MANIFEST_HASHES = {
  // Word Id: 27ee7761-18ca-4fae-b482-4171d5afa6cd
  "word.xml": "e4c72462952e21aad24517c5fe474dfb370d05ff99b279f81bb3a44b9fbc834d",
  // Excel Id: 4b5861d6-18ee-4d4a-b9ee-d43d7da80ff4
  "excel.xml": "4573041c6fcd00e7de42cf5ec7c39424cd963bceb936afde98ea298e8a91fd2f",
};

export function manifestHash(text) {
  return createHash("sha256")
    .update(text.replace(/^\uFEFF/, "").replace(/\r\n/g, "\n"), "utf8")
    .digest("hex");
}

export function isLegacyManifest(file, text) {
  return (
    Object.hasOwn(LEGACY_MANIFEST_HASHES, file) &&
    manifestHash(text) === LEGACY_MANIFEST_HASHES[file]
  );
}
