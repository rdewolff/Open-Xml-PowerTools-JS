import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, MarkupSimplifier } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("MarkupSimplifier: remove content controls, goBack bookmark, soft hyphens, rsid attrs", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p w:rsidR="00112233">
      <w:bookmarkStart w:id="0" w:name="_GoBack"/>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>he\u00ADllo</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);

  const simplified = await MarkupSimplifier.simplifyMarkup(doc, {
    removeContentControls: true,
    removeGoBackBookmark: true,
    removeSoftHyphens: true,
    removeRsidInfo: true,
  });

  const { text } = await simplified.getMainDocumentText();
  assert.equal(text, "hello");

  const newDocXmlText = await simplified.getPartText("/word/document.xml");
  assert.doesNotMatch(newDocXmlText, /<w:sdt>/);
  assert.doesNotMatch(newDocXmlText, /_GoBack/);
  assert.doesNotMatch(newDocXmlText, /rsidR/);
});

