import { inflateRawSync, deflateRawSync } from "node:zlib";

export const ZipAdapterNode = {
  async inflateRaw(data) {
    return new Uint8Array(inflateRawSync(data));
  },
  async deflateRaw(data, options = {}) {
    return new Uint8Array(deflateRawSync(data, { level: options.level ?? 6 }));
  },
};
