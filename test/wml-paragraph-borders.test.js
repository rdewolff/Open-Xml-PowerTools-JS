import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: maps paragraph shading and borders to CSS", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:shd w:fill="FFFF00"/>
        <w:pBdr>
          <w:top w:val="single" w:sz="8" w:color="FF0000"/>
          <w:left w:val="dotted" w:sz="8" w:color="00FF00"/>
          <w:bottom w:val="dashed" w:sz="8" w:color="0000FF"/>
          <w:right w:val="double" w:sz="8" w:color="000000"/>
        </w:pBdr>
      </w:pPr>
      <w:r><w:t>Para</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);

  assert.match(res.html, /background-color:#FFFF00/);
  assert.match(res.html, /border-top:1pt solid #FF0000/);
  assert.match(res.html, /border-left:1pt dotted #00FF00/);
  assert.match(res.html, /border-bottom:1pt dashed #0000FF/);
  assert.match(res.html, /border-right:1pt double #000000/);
});

