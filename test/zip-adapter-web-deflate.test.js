import test from "node:test";
import assert from "node:assert/strict";
import { deflateRawSync, deflateSync, inflateRawSync } from "node:zlib";
import { ZipAdapterWeb } from "../src/internal/zip-adapter-web.js";

function installCompressionStreamStub({ mode }) {
  const original = globalThis.CompressionStream;

  class StubCompressionStream {
    constructor(format) {
      if (format !== "deflate") throw new TypeError("Unsupported format");
      const chunks = [];
      const ts = new TransformStream({
        transform(chunk) {
          chunks.push(chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk));
        },
        flush(controller) {
          const total = chunks.reduce((n, c) => n + c.length, 0);
          const input = new Uint8Array(total);
          let offset = 0;
          for (const c of chunks) {
            input.set(c, offset);
            offset += c.length;
          }

          const out = mode === "raw" ? deflateRawSync(input) : deflateSync(input);
          controller.enqueue(new Uint8Array(out));
        },
      });
      this.readable = ts.readable;
      this.writable = ts.writable;
    }
  }

  globalThis.CompressionStream = StubCompressionStream;
  return () => {
    globalThis.CompressionStream = original;
  };
}

test("ZipAdapterWeb.deflateRaw: does not strip when CompressionStream('deflate') returns raw deflate", async (t) => {
  const restore = installCompressionStreamStub({ mode: "raw" });
  t.after(restore);

  const input = new TextEncoder().encode("hello hello hello");
  const compressed = await ZipAdapterWeb.deflateRaw(input);
  const inflated = new Uint8Array(inflateRawSync(compressed));
  assert.deepEqual(inflated, input);
});

test("ZipAdapterWeb.deflateRaw: strips zlib wrapper when CompressionStream('deflate') returns zlib-wrapped deflate", async (t) => {
  const restore = installCompressionStreamStub({ mode: "zlib" });
  t.after(restore);

  const input = new TextEncoder().encode("hello hello hello");
  const compressed = await ZipAdapterWeb.deflateRaw(input);
  const inflated = new Uint8Array(inflateRawSync(compressed));
  assert.deepEqual(inflated, input);
});

