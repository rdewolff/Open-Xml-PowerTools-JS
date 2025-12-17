import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: maps tblBorders to CSS borders", async () => {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:color="FF0000"/>
          <w:left w:val="single" w:sz="8" w:color="00FF00"/>
          <w:bottom w:val="single" w:sz="8" w:color="0000FF"/>
          <w:right w:val="single" w:sz="8" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({ documentXml });
  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc);

  assert.match(res.html, /border-collapse:collapse/);
  assert.match(res.html, /border-top:1pt solid #FF0000/);
  assert.match(res.html, /border-left:1pt solid #00FF00/);
  assert.match(res.html, /border-bottom:1pt solid #0000FF/);
  assert.match(res.html, /border-right:1pt solid #000000/);
});

