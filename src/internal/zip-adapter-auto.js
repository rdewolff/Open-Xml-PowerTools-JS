import { OpenXmlPowerToolsError } from "../open-xml-powertools-error.js";

export async function getDefaultZipAdapter() {
  // Node.js: prefer raw deflate/inflate via node:zlib, but avoid top-level imports
  // so that the main package can be imported in browsers without bundler shims.
  if (isNodeLike()) {
    try {
      const zlib = await import("node:zlib");
      return {
        inflateRaw(data) {
          return new Uint8Array(zlib.inflateRawSync(data));
        },
        deflateRaw(data, options = {}) {
          return new Uint8Array(zlib.deflateRawSync(data, { level: options.level ?? 6 }));
        },
      };
    } catch (e) {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "Failed to load Node zlib adapter", e);
    }
  }

  throw new OpenXmlPowerToolsError(
    "OXPT_ZIP_UNSUPPORTED",
    "No built-in ZIP adapter available in this runtime; pass an adapter with inflateRaw/deflateRaw",
  );
}

function isNodeLike() {
  return (
    typeof process !== "undefined" &&
    process &&
    typeof process === "object" &&
    process.versions &&
    typeof process.versions === "object" &&
    typeof process.versions.node === "string"
  );
}

