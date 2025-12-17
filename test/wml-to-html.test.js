import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: converts paragraphs and basic run formatting", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>World</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);

  const res = await WmlToHtmlConverter.convertToHtml(doc, { additionalCss: "body { max-width: 20cm; }" });
  assert.match(res.html, /<p>/);
  assert.match(res.html, />Hello /);
  assert.match(res.html, /<strong>World<\/strong>/);
  assert.match(res.html, /max-width:\s*20cm/);
});

