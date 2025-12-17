import test from "node:test";
import assert from "node:assert/strict";
import { HtmlToWmlConverter } from "../src/index.js";

test("HtmlToWmlConverter: converts simple XHTML into a DOCX with expected text and formatting", async () => {
  const xhtml = `<?xml version="1.0" encoding="UTF-8"?>
<html>
  <body>
    <p>Hello <strong>World</strong><br/>Line2</p>
  </body>
</html>`;

  const doc = await HtmlToWmlConverter.convertHtmlToWml("", "", "", xhtml, {});
  const { paragraphs, text } = await doc.getMainDocumentText();
  assert.deepEqual(paragraphs, ["Hello WorldLine2"]);
  assert.equal(text, "Hello WorldLine2");

  const docXml = await doc.getPartText("/word/document.xml");
  assert.match(docXml, /<w:b\s*\/>/);
  assert.match(docXml, /<w:br\s*\/>/);
});

