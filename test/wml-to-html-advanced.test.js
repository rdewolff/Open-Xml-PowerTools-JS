import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

const PNG_1X1_TRANSPARENT = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

test("WmlToHtmlConverter: headings, hyperlinks, tables, lists, images", async () => {
  const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="9">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>
`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Title</w:t></w:r>
    </w:p>

    <w:p>
      <w:hyperlink r:id="rIdLink">
        <w:r><w:t>Example</w:t></w:r>
      </w:hyperlink>
    </w:p>

    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="9"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>Item 1</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="9"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>Item 2</w:t></w:r>
    </w:p>

    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>

    <w:p>
      <w:r>
        <w:drawing>
          <a:blip r:embed="rIdImg"/>
        </w:drawing>
      </w:r>
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
      {
        Id: "rIdImg",
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Target: "media/image1.png",
      },
    ],
    contentTypes: {
      defaults: [{ Extension: "png", ContentType: "image/png" }],
      overrides: [
        {
          PartName: "/word/numbering.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        },
      ],
    },
    extraEntries: [
      { name: "word/numbering.xml", text: numberingXml },
      { name: "word/media/image1.png", bytes: PNG_1X1_TRANSPARENT, compressionMethod: 0 },
    ],
  });

  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc, { output: { format: "xml" } });

  assert.match(res.html, /<h1[^>]*>.*Title.*<\/h1>/);
  assert.match(res.html, /<a href="https:\/\/example\.com\/">/);
  assert.match(res.html, /<ol[^>]*>/);
  assert.match(res.html, /<li>.*Item 1.*<\/li>/);
  assert.match(res.html, /<table>/);
  assert.match(res.html, /<td>.*A1.*<\/td>/);
  assert.match(res.html, /<img src="data:image\/png;base64,/);
  assert.ok(res.htmlElement, "expected htmlElement when output.format=xml");
});
