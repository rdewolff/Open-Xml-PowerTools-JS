import { OpenXmlPowerToolsError } from "../open-xml-powertools-error.js";
import { inflateRaw as inflateRawJs } from "./inflate-raw.js";

export const ZipAdapterWeb = {
  async inflateRaw(data, options = {}) {
    // Prefer built-in raw deflate (when available), otherwise fall back to a pure-JS inflater.
    if (typeof DecompressionStream !== "undefined") {
      const format = pickDecompressionFormat();
      if (format === "deflate-raw") {
        try {
          const stream = new DecompressionStream("deflate-raw");
          const bytes = await streamToBytes(new Blob([data]).stream().pipeThrough(stream));
          if (options.expectedSize != null && bytes.length !== options.expectedSize) {
            throw new OpenXmlPowerToolsError(
              "OXPT_ZIP_INVALID",
              `Inflated size mismatch (expected ${options.expectedSize}, got ${bytes.length})`,
            );
          }
          return bytes;
        } catch {
          // Some runtimes claim to support `deflate-raw` but fail at runtime (e.g. "incorrect header check").
          // Fall back to our pure-JS inflater for compatibility.
        }
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
    // ZIP requires raw deflate, so strip the zlib wrapper (if present).
    if (format === "deflate") {
      const stream = new CompressionStream("deflate");
      const zlibBytes = await streamToBytes(new Blob([data]).stream().pipeThrough(stream));
      return stripZlibWrapperIfPresent(zlibBytes);
    }

    throw new OpenXmlPowerToolsError(
      "OXPT_ZIP_UNSUPPORTED",
      "This browser does not support CompressionStream('deflate-raw'|'deflate'); provide a zipAdapter with inflateRaw/deflateRaw",
    );
  },
};

function stripZlibWrapperIfPresent(bytes) {
  // Some runtimes label a raw-deflate stream as "deflate" (no zlib wrapper).
  // Only strip when the output actually looks like zlib-wrapped deflate.
  if (bytes.length >= 2 && bytes[0] === 0x1f && bytes[1] === 0x8b) {
    throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "CompressionStream('deflate') returned gzip data");
  }
  if (!looksLikeZlibHeader(bytes)) return bytes;

  // zlib wrapper is: 2-byte header (+ optional 4-byte dict id) + raw deflate data + 4-byte Adler32
  if (bytes.length < 6) {
    throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "deflate output too small to be zlib-wrapped");
  }

  const cmf = bytes[0];
  const flg = bytes[1];
  if ((flg & 0x20) !== 0) {
    throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "zlib preset dictionary is not supported");
  }

  return bytes.subarray(2, bytes.length - 4);
}

function looksLikeZlibHeader(bytes) {
  if (bytes.length < 2) return false;
  const cmf = bytes[0];
  const flg = bytes[1];

  // CMF: compression method (low 4 bits) must be 8 for deflate.
  if ((cmf & 0x0f) !== 8) return false;
  // CMF: CINFO (window size) is (cmf >> 4), must be <= 7 for 32K window.
  if ((cmf >>> 4) > 7) return false;
  // Header check: (CMF*256 + FLG) is a multiple of 31.
  if ((((cmf << 8) | flg) % 31) !== 0) return false;
  return true;
}

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
