import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, RevisionAccepter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("RevisionAccepter: accept <w:ins> and drop <w:del>", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:ins>
        <w:r><w:t>Inserted</w:t></w:r>
      </w:ins>
      <w:del>
        <w:r><w:delText>Deleted</w:delText></w:r>
      </w:del>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);

  assert.equal(await RevisionAccepter.hasTrackedRevisions(doc), true);

  const accepted = await RevisionAccepter.acceptRevisions(doc);
  assert.equal(await RevisionAccepter.hasTrackedRevisions(accepted), false);

  const { text } = await accepted.getMainDocumentText();
  assert.equal(text, "Inserted");
});

test("RevisionAccepter: accepts tracked revisions in header/footer parts too", async () => {
  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:ins><w:r><w:t>H+</w:t></w:r></w:ins>
    <w:del><w:r><w:delText>H-</w:delText></w:r></w:del>
  </w:p>
</w:hdr>
`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId10"/>
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
    ],
    contentTypes: {
      overrides: [
        {
          PartName: "/word/header1.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        },
      ],
    },
    extraEntries: [{ name: "word/header1.xml", text: headerXml }],
  });

  const doc = new WmlDocument(bytes);
  assert.equal(await RevisionAccepter.hasTrackedRevisions(doc), true);

  const accepted = await RevisionAccepter.acceptRevisions(doc);
  const newHeader = await accepted.getPartText("/word/header1.xml");
  assert.equal(newHeader.includes("<w:ins"), false);
  assert.equal(newHeader.includes("<w:del"), false);
  assert.match(newHeader, /H\+/);
  assert.equal(newHeader.includes("H-"), false);
});
