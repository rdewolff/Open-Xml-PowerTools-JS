import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: renders header and footer from sectPr references", async () => {
  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Header text</w:t></w:r></w:p>
</w:hdr>
`;

  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Footer text</w:t></w:r></w:p>
</w:ftr>
`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId10"/>
      <w:footerReference w:type="default" r:id="rId11"/>
    </w:sectPr>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({
    documentXml,
    documentRelationships: [
      {
        Id: "rId10",
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
        Target: "header1.xml",
      },
      {
        Id: "rId11",
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        Target: "footer1.xml",
      },
    ],
    contentTypes: {
      overrides: [
        {
          PartName: "/word/header1.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        },
        {
          PartName: "/word/footer1.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        },
      ],
    },
    extraEntries: [
      { name: "word/header1.xml", text: headerXml },
      { name: "word/footer1.xml", text: footerXml },
    ],
  });

  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);

  assert.match(res.html, /class="pt-header"/);
  assert.match(res.html, /Header text/);
  assert.match(res.html, /Body/);
  assert.match(res.html, /class="pt-footer"/);
  assert.match(res.html, /Footer text/);
});

