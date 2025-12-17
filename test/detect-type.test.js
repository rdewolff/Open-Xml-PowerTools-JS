import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { OpenXmlPowerToolsDocument } from "../src/index.js";
import { base64ToBytes } from "../src/util/base64.js";

test("OpenXmlPowerToolsDocument.detectType: detects DOCX", async () => {
  const base64 = await readFile(new URL("./fixtures/minimal.docx.base64", import.meta.url), "utf8");
  const bytes = base64ToBytes(base64);
  const doc = new OpenXmlPowerToolsDocument(bytes);
  assert.equal(await doc.detectType(), "docx");
});

