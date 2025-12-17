import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: restrictToSupportedLanguages warns when no locale-specific list implementation", async () => {
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr>
      </w:pPr>
      <w:r>
        <w:rPr><w:lang w:val="de-DE"/></w:rPr>
        <w:t>Item</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({
    documentXml,
    contentTypes: {
      overrides: [
        {
          PartName: "/word/numbering.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        },
      ],
    },
    extraEntries: [
      { name: "word/numbering.xml", text: numberingXml },
    ],
  });

  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc, {
    restrictToSupportedLanguages: true,
    listItemImplementations: {
      "en-US": (_lvlText, n) => `${n}.`,
    },
  });

  assert.ok(res.warnings.some((w) => w.code === "OXPT_LIST_LANG_UNSUPPORTED"), "expected language warning");
  assert.match(res.html, /<ol/);
});
