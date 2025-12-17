import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, TextReplacer } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("TextReplacer: replaces text across run boundaries and preserves first-run formatting", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Hello </w:t>
      </w:r>
      <w:r><w:t>World</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes, { fileName: "in.docx" });

  const replaced = await TextReplacer.searchAndReplace(doc, "Hello World", "Hi", { matchCase: true });
  const { text } = await replaced.getMainDocumentText();
  assert.equal(text, "Hi");

  const newDocXmlText = await replaced.getPartText("/word/document.xml");
  assert.match(newDocXmlText, /<w:b\s*\/>/);
});

test("TextReplacer: replaces text in header parts too", async () => {
  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>He</w:t></w:r>
    <w:r><w:t>llo</w:t></w:r>
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

  const replaced = await TextReplacer.searchAndReplace(doc, "Hello", "Hi", { matchCase: true });
  const newHeader = await replaced.getPartText("/word/header1.xml");
  assert.match(newHeader, />Hi</);
  assert.doesNotMatch(newHeader, /Hello/);
});
