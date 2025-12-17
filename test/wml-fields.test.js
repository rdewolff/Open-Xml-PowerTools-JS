import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: ignores field code plumbing and renders field result text", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText>DATE</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>2025-01-01</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);

  assert.match(res.html, /2025-01-01/);
  assert.equal(res.warnings.some((w) => w.code === "OXPT_HTML_UNSUPPORTED_RUN_CHILD"), false);
});

