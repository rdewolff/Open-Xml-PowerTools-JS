import test from "node:test";
import assert from "node:assert/strict";
import { deflateRawSync, constants as zlibConstants } from "node:zlib";
import { inflateRaw } from "../src/internal/inflate-raw.js";

function u8(input) {
  return input instanceof Uint8Array ? input : new TextEncoder().encode(String(input));
}

test("inflateRaw (JS): round-trips stored block (level 0)", () => {
  const original = u8("Hello Hello Hello Hello Hello Hello");
  const compressed = deflateRawSync(original, { level: 0 });
  const out = inflateRaw(new Uint8Array(compressed), { expectedSize: original.length });
  assert.deepEqual(out, original);
});

test("inflateRaw (JS): round-trips fixed Huffman blocks (Z_FIXED)", () => {
  const original = u8(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 ".repeat(20) +
      "overlap-overlap-overlap-overlap".repeat(30),
  );
  const compressed = deflateRawSync(original, { level: 6, strategy: zlibConstants.Z_FIXED });
  const out = inflateRaw(new Uint8Array(compressed), { expectedSize: original.length });
  assert.deepEqual(out, original);
});

test("inflateRaw (JS): round-trips dynamic Huffman blocks (default)", () => {
  const original = u8("This is some text that should compress well. ".repeat(200));
  const compressed = deflateRawSync(original, { level: 6 });
  const out = inflateRaw(new Uint8Array(compressed), { expectedSize: original.length });
  assert.deepEqual(out, original);
});

test("inflateRaw (JS): throws on truncated input", () => {
  const original = u8("Truncated stream test ".repeat(50));
  const compressed = new Uint8Array(deflateRawSync(original, { level: 6 }));
  const truncated = compressed.subarray(0, Math.max(0, compressed.length - 4));
  assert.throws(() => inflateRaw(truncated, { expectedSize: original.length }));
});
