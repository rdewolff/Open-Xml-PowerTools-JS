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

test("MarkupSimplifier: remove hyperlinks, field codes, and note references", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rIdLink">
        <w:r><w:t>Click</w:t></w:r>
      </w:hyperlink>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText>HYPERLINK \"https://example.com\"</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t>Shown</w:t></w:r>
      <w:r><w:footnoteReference w:id="1"/></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({
    documentXml,
    documentRelationships: [
      {
        Id: "rIdLink",
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        Target: "https://example.com/",
        TargetMode: "External",
      },
    ],
  });
  const doc = new WmlDocument(bytes);

  const simplified = await MarkupSimplifier.simplifyMarkup(doc, {
    removeHyperlinks: true,
    removeFieldCodes: true,
    removeEndAndFootNotes: true,
  });

  const { text } = await simplified.getMainDocumentText();
  assert.equal(text, "ClickShown");

  const xml = await simplified.getPartText("/word/document.xml");
  assert.doesNotMatch(xml, /<w:hyperlink/);
  assert.doesNotMatch(xml, /<w:instrText/);
  assert.doesNotMatch(xml, /<w:fldChar/);
  assert.doesNotMatch(xml, /footnoteReference/);
});
