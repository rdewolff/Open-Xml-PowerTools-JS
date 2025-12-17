import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: applies paragraph and character styles from styles.xml", async () => {
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="MyPara">
    <w:name w:val="MyPara"/>
    <w:rPr>
      <w:b/>
      <w:color w:val="FF0000"/>
      <w:sz w:val="48"/>
    </w:rPr>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="240" w:after="240"/>
    </w:pPr>
  </w:style>

  <w:style w:type="character" w:styleId="MyChar">
    <w:name w:val="MyChar"/>
    <w:rPr>
      <w:i/>
    </w:rPr>
  </w:style>
</w:styles>
`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="MyPara"/></w:pPr>
      <w:r><w:t>Styled</w:t></w:r>
      <w:r>
        <w:rPr><w:rStyle w:val="MyChar"/></w:rPr>
        <w:t> Italic</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml, extraEntries: [{ name: "word/styles.xml", text: stylesXml }] });
  const doc = new WmlDocument(bytes);

  const res = await WmlToHtmlConverter.convertToHtml(doc, { fabricateCssClasses: true });
  assert.match(res.html, /class="pt-p-mypara"/);
  assert.match(res.cssText, /\.pt-p-mypara/);
  assert.match(res.html, /<strong>Styled<\/strong>/);
  assert.match(res.html, /style="[^"]*color:#FF0000[^"]*"/);
  assert.match(res.html, /font-size:24pt/);
  assert.match(res.html, /<em> Italic<\/em>/);
  assert.match(res.cssText, /\.pt-r-mychar/);
});

