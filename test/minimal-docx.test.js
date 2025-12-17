import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { WmlDocument } from "../src/index.js";
import { base64ToBytes } from "../src/util/base64.js";
import { OpcPackage } from "../src/internal/opc.js";

test("Phase 0: parse minimal DOCX and extract expected text", async () => {
  const base64 = await readFile(new URL("./fixtures/minimal.docx.base64", import.meta.url), "utf8");
  const bytes = base64ToBytes(base64);
  const doc = WmlDocument.fromBytes(bytes, { fileName: "minimal.docx" });

  const pkg = await OpcPackage.fromBytes(doc.toBytes());
  await pkg.assertIsValidOpc();
  const officeUri = await pkg.getOfficeDocumentPartUri();
  assert.equal(officeUri, "/word/document.xml");

  const { paragraphs, text } = await doc.getMainDocumentText();
  assert.deepEqual(paragraphs, ["Hello OpenXmlPowerTools-JS"]);
  assert.equal(text, "Hello OpenXmlPowerTools-JS");
});

