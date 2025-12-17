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

test("WmlToHtmlConverter: renders different headers/footers per section", async () => {
  const header1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Header 1</w:t></w:r></w:p>
</w:hdr>
`;

  const footer1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Footer 1</w:t></w:r></w:p>
</w:ftr>
`;

  const header2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Header 2</w:t></w:r></w:p>
</w:hdr>
`;

  const footer2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Footer 2</w:t></w:r></w:p>
</w:ftr>
`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>S1-A</w:t></w:r></w:p>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:headerReference w:type="default" r:id="rId10"/>
          <w:footerReference w:type="default" r:id="rId11"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>S1-B</w:t></w:r>
    </w:p>

    <w:p><w:r><w:t>S2</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId12"/>
      <w:footerReference w:type="default" r:id="rId13"/>
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
      {
        Id: "rId12",
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
        Target: "header2.xml",
      },
      {
        Id: "rId13",
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        Target: "footer2.xml",
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
        {
          PartName: "/word/header2.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        },
        {
          PartName: "/word/footer2.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        },
      ],
    },
    extraEntries: [
      { name: "word/header1.xml", text: header1Xml },
      { name: "word/footer1.xml", text: footer1Xml },
      { name: "word/header2.xml", text: header2Xml },
      { name: "word/footer2.xml", text: footer2Xml },
    ],
  });

  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);

  assert.match(res.html, /Header 1/);
  assert.match(res.html, /Footer 1/);
  assert.match(res.html, /Header 2/);
  assert.match(res.html, /Footer 2/);

  const idxHeader1 = res.html.indexOf("Header 1");
  const idxS1A = res.html.indexOf("S1-A");
  const idxHeader2 = res.html.indexOf("Header 2");
  const idxS2 = res.html.indexOf("S2");
  assert.ok(idxHeader1 !== -1 && idxS1A !== -1 && idxHeader2 !== -1 && idxS2 !== -1);
  assert.ok(idxHeader1 < idxS1A);
  assert.ok(idxHeader2 < idxS2);
  assert.ok(idxS1A < idxHeader2);
});
