import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { WmlDocument } from "../src/index.js";
import { ZipAdapterWeb } from "../src/internal/zip-adapter-web.js";

test("Local fixture: Whisperit v3 changes.docx loads via ZipAdapterWeb even if DecompressionStream fails", async (t) => {
  const fixtureUrl = new URL("./fixtures/local/Whisperit v3 changes.docx", import.meta.url);

  let bytes;
  try {
    bytes = new Uint8Array(await readFile(fixtureUrl));
  } catch (e) {
    if (e?.code === "ENOENT") {
      t.skip("Local fixture not present (see `test/fixtures/README.md`).");
      return;
    }
    throw e;
  }

  const originalDecompressionStream = globalThis.DecompressionStream;
  t.after(() => {
    globalThis.DecompressionStream = originalDecompressionStream;
  });

  let attempted = false;
  class BadDecompressionStream {
    constructor() {
      const ts = new TransformStream({
        transform() {
          attempted = true;
          throw new Error("The compressed data was not valid: incorrect header check.");
        },
      });
      this.readable = ts.readable;
      this.writable = ts.writable;
    }
  }

  globalThis.DecompressionStream = BadDecompressionStream;

  const doc = WmlDocument.fromBytes(bytes, { fileName: "Whisperit v3 changes.docx", zipAdapter: ZipAdapterWeb });
  const { text } = await doc.getMainDocumentText();
  assert.ok(attempted, "Expected DecompressionStream path to be attempted");
  assert.ok(text.includes("Whisperit version 3"));
});

