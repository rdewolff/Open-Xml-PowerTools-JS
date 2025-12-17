import { inflateRawSync, deflateRawSync } from "node:zlib";

export const ZipAdapterNode = {
  inflateRaw(data) {
    return new Uint8Array(inflateRawSync(data));
  },
  deflateRaw(data, options = {}) {
    return new Uint8Array(deflateRawSync(data, { level: options.level ?? 6 }));
  },
};

