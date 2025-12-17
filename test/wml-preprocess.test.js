import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: preprocess options can disable acceptRevisions", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:ins>
        <w:r><w:t>Inserted</w:t></w:r>
      </w:ins>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);

  const withDefault = await WmlToHtmlConverter.convertToHtml(doc);
  assert.match(withDefault.html, /Inserted/);

  const withoutAccept = await WmlToHtmlConverter.convertToHtml(doc, { preprocess: { acceptRevisions: false } });
  assert.doesNotMatch(withoutAccept.html, /Inserted/);
});

