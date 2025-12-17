import { OpenXmlPowerToolsError } from "../open-xml-powertools-error.js";

export const ZipAdapterWeb = {
  async inflateRaw(data) {
    if (typeof DecompressionStream === "undefined") {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "DecompressionStream is not available");
    }
    const stream = new DecompressionStream("deflate");
    const out = await streamToBytes(new Blob([data]).stream().pipeThrough(stream));
    return out;
  },

  async deflateRaw(data) {
    if (typeof CompressionStream === "undefined") {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "CompressionStream is not available");
    }
    const stream = new CompressionStream("deflate");
    const out = await streamToBytes(new Blob([data]).stream().pipeThrough(stream));
    return out;
  },
};

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

