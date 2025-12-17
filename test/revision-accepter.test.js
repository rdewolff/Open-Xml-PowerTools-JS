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

