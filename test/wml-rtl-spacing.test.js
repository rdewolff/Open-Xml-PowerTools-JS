import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: bidi paragraphs map to RTL + line-height", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:bidi/>
        <w:spacing w:line="360" w:lineRule="auto"/>
      </w:pPr>
      <w:r><w:t>RTL para</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr><w:rtl/></w:rPr>
        <w:t>RTL run</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);

  assert.match(res.html, /<p[^>]*dir="rtl"[^>]*>/);
  assert.match(res.html, /line-height:1\.5/);
  assert.match(res.html, /<span[^>]*dir="rtl"[^>]*>[^<]*RTL run/);
});
