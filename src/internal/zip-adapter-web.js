import { OpenXmlPowerToolsError } from "../open-xml-powertools-error.js";
import { inflateRaw as inflateRawJs } from "./inflate-raw.js";

export const ZipAdapterWeb = {
  async inflateRaw(data, options = {}) {
    // Prefer built-in raw deflate (when available), otherwise fall back to a pure-JS inflater.
    if (typeof DecompressionStream !== "undefined") {
      const format = pickDecompressionFormat();
      if (format === "deflate-raw") {
        const stream = new DecompressionStream("deflate-raw");
        const bytes = await streamToBytes(new Blob([data]).stream().pipeThrough(stream));
        if (options.expectedSize != null && bytes.length !== options.expectedSize) {
          throw new OpenXmlPowerToolsError(
            "OXPT_ZIP_INVALID",
            `Inflated size mismatch (expected ${options.expectedSize}, got ${bytes.length})`,
          );
        }
        return bytes;
      }
    }

    // NOTE: DecompressionStream('deflate') is typically zlib-wrapped "deflate" and cannot
    // decode raw ZIP deflate streams. Use JS fallback for broad browser compatibility.
    return inflateRawJs(data, options);
  },

  async deflateRaw(data) {
    if (typeof CompressionStream === "undefined") {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "CompressionStream is not available");
    }
    const format = pickCompressionFormat();
    if (format === "deflate-raw") {
      const stream = new CompressionStream("deflate-raw");
      return await streamToBytes(new Blob([data]).stream().pipeThrough(stream));
    }

    // Fallback: CompressionStream('deflate') typically emits zlib-wrapped deflate.
    // ZIP requires raw deflate, so strip the 2-byte zlib header and 4-byte Adler32 trailer.
    if (format === "deflate") {
      const stream = new CompressionStream("deflate");
      const zlibBytes = await streamToBytes(new Blob([data]).stream().pipeThrough(stream));
      if (zlibBytes.length < 6) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "deflate output too small to be zlib-wrapped");
      }
      return zlibBytes.subarray(2, zlibBytes.length - 4);
    }

    throw new OpenXmlPowerToolsError(
      "OXPT_ZIP_UNSUPPORTED",
      "This browser does not support CompressionStream('deflate-raw'|'deflate'); provide a zipAdapter with inflateRaw/deflateRaw",
    );
  },
};

function pickDecompressionFormat() {
  // Prefer raw deflate for ZIP. Some browsers support "deflate" but not "deflate-raw".
  try {
    // eslint-disable-next-line no-new
    new DecompressionStream("deflate-raw");
    return "deflate-raw";
  } catch {
    // fall through
  }
  try {
    // eslint-disable-next-line no-new
    new DecompressionStream("deflate");
    return "deflate";
  } catch {
    return null;
  }
}

function pickCompressionFormat() {
  try {
    // eslint-disable-next-line no-new
    new CompressionStream("deflate-raw");
    return "deflate-raw";
  } catch {
    // fall through
  }
  try {
    // eslint-disable-next-line no-new
    new CompressionStream("deflate");
    return "deflate";
  } catch {
    return null;
  }
}

async function streamToBytes(readable) {
  const reader = readable.getReader();
  const chunks = [];
  let total = 0;
  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    const chunk = value instanceof Uint8Array ? value : new Uint8Array(value);
    chunks.push(chunk);
    total += chunk.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out;
}
