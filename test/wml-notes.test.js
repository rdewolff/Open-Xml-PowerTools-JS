import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: renders footnote references and footnotes section", async () => {
  const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p><w:r><w:t>Footnote text</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>
`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello</w:t></w:r>
      <w:r><w:footnoteReference w:id="1"/></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({
    documentXml,
    contentTypes: {
      overrides: [
        {
          PartName: "/word/footnotes.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        },
      ],
    },
    extraEntries: [{ name: "word/footnotes.xml", text: footnotesXml }],
  });

  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);
  assert.match(res.html, /<sup><a href="#pt-footnote-1">1<\/a><\/sup>/);
  assert.match(res.html, /<ol class="pt-footnotes">/);
  assert.match(res.html, /id="pt-footnote-1"/);
  assert.match(res.html, /Footnote text/);
});
