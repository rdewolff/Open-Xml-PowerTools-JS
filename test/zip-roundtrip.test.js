import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { base64ToBytes } from "../src/util/base64.js";
import { ZipArchive } from "../src/internal/zip.js";
import { ZipAdapterNode } from "../src/internal/zip-adapter-node.js";
import { WmlDocument } from "../src/wml-document.js";

test("Phase 1: can read and rebuild DOCX zip, preserving part bytes", async () => {
  const base64 = await readFile(new URL("./fixtures/minimal.docx.base64", import.meta.url), "utf8");
  const bytes = base64ToBytes(base64);

  const zip = await ZipArchive.fromBytes(bytes, { adapter: ZipAdapterNode });
  const names = zip.entries.map((e) => e.name).sort();
  assert.ok(names.includes("[Content_Types].xml"));
  assert.ok(names.includes("word/document.xml"));

  const originalDocXml = await zip.getEntryBytes("word/document.xml");

  const rebuilt2 = await ZipArchive.build(
    await Promise.all(
      zip.entries.map(async (e) => ({
        name: e.name,
        bytes: await e.getBytes(),
        compressionMethod: 8,
      })),
    ),
    { adapter: ZipAdapterNode, level: 6 },
  );

  const zip2 = await ZipArchive.fromBytes(rebuilt2, { adapter: ZipAdapterNode });
  const doc2 = new WmlDocument(rebuilt2, { fileName: "rebuilt.docx" });
  const roundtrippedDocXml = await zip2.getEntryBytes("word/document.xml");

  assert.deepEqual(roundtrippedDocXml, originalDocXml);
  const { text } = await doc2.getMainDocumentText();
  assert.equal(text, "Hello OpenXmlPowerTools-JS");
});
