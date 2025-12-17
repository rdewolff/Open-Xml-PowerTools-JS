import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: maps strike and vertAlign (sup/sub)", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:rPr><w:strike/></w:rPr><w:t>Strike</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>Sup</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>Sub</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);

  assert.match(res.html, /<s>Strike<\/s>/);
  assert.match(res.html, /<sup>Sup<\/sup>/);
  assert.match(res.html, /<sub>Sub<\/sub>/);
});

test("WmlToHtmlConverter: maps caps and smallCaps to CSS", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:rPr><w:caps/></w:rPr><w:t>Caps</w:t></w:r>
      <w:r><w:rPr><w:smallCaps/></w:rPr><w:t>SmallCaps</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);

  assert.match(res.html, /text-transform:uppercase/);
  assert.match(res.html, /font-variant:small-caps/);
});

