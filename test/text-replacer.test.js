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
