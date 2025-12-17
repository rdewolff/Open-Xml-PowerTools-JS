import test from "node:test";
import assert from "node:assert/strict";
import { HtmlToWmlConverter, WmlToHtmlConverter } from "../src/index.js";

const PNG_1X1_TRANSPARENT_DATA_URL =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";

test("HtmlToWmlConverter: headings, lists, tables, hyperlinks, images", async () => {
  const xhtml = `<?xml version="1.0" encoding="UTF-8"?>
<html>
  <body>
    <h1>Title</h1>
    <p><a href="https://example.com/">Link</a></p>
    <ol>
      <li>One</li>
      <li>Two<ul><li>Nested</li></ul></li>
    </ol>
    <table>
      <tr><td>A1</td><td>B1</td></tr>
    </table>
    <p><img src="${PNG_1X1_TRANSPARENT_DATA_URL}" /></p>
  </body>
</html>`;

  const doc = await HtmlToWmlConverter.convertHtmlToWml("", "", "", xhtml, {});
  const { text } = await doc.getMainDocumentText();
  assert.match(text, /Title/);
  assert.match(text, /Link/);
  assert.match(text, /One/);
  assert.match(text, /Nested/);
  assert.match(text, /A1/);

  // Round-trip to HTML to validate the generated DOCX structure is usable by our converter.
  const html = await WmlToHtmlConverter.convertToHtml(doc);
  assert.match(html.html, /<h1[^>]*>.*Title.*<\/h1>/);
  assert.match(html.html, /<a href="https:\/\/example\.com\/">/);
  assert.match(html.html, /<ol/);
  assert.match(html.html, /<table/);
  assert.match(html.html, /data:image\/png;base64/);
});

test("HtmlToWmlConverter: maps table colspan/rowspan and image width/height", async () => {
  const xhtml = `<?xml version="1.0" encoding="UTF-8"?>
<html>
  <body>
    <table>
      <tr><td colspan="2">A</td><td>B</td></tr>
      <tr><td rowspan="2">C</td><td>D</td><td>E</td></tr>
      <tr><td>F</td><td>G</td></tr>
    </table>
    <p><img width="96" height="48" src="${PNG_1X1_TRANSPARENT_DATA_URL}" /></p>
  </body>
</html>`;

  const doc = await HtmlToWmlConverter.convertHtmlToWml("", "", "", xhtml, {});
  const mainXml = await doc.getPartText("/word/document.xml");
  assert.match(mainXml, /<w:gridSpan[^>]*w:val="2"/);
  assert.match(mainXml, /<w:vMerge[^>]*w:val="restart"/);
  assert.match(mainXml, /<w:vMerge\s*\/>/);
  assert.match(mainXml, /<wp:extent[^>]*cx="914400"[^>]*cy="457200"/);

  const html = await WmlToHtmlConverter.convertToHtml(doc);
  assert.match(html.html, /colspan="2"/);
  assert.match(html.html, /rowspan="2"/);
});

test("HtmlToWmlConverter: supports thead/th headers and sup/sub/strike", async () => {
  const xhtml = `<?xml version="1.0" encoding="UTF-8"?>
<html>
  <body>
    <p>H<sup>2</sup>O <sub>x</sub> <s>gone</s></p>
    <table>
      <thead>
        <tr><th>Head</th></tr>
      </thead>
      <tbody>
        <tr><td>Body</td></tr>
      </tbody>
    </table>
  </body>
</html>`;

  const doc = await HtmlToWmlConverter.convertHtmlToWml("", "", "", xhtml, {});
  const mainXml = await doc.getPartText("/word/document.xml");
  assert.match(mainXml, /<w:tblHeader\s*\/>/);
  assert.match(mainXml, /<w:vertAlign[^>]*w:val="superscript"/);
  assert.match(mainXml, /<w:vertAlign[^>]*w:val="subscript"/);
  assert.match(mainXml, /<w:strike\s*\/>/);

  const html = await WmlToHtmlConverter.convertToHtml(doc);
  assert.match(html.html, /<thead>/);
  assert.match(html.html, /<th>.*Head.*<\/th>/);
});

test("HtmlToWmlConverter: supports internal hyperlinks via bookmarks", async () => {
  const xhtml = `<?xml version="1.0" encoding="UTF-8"?>
<html>
  <body>
    <p><a id="BM1"></a>Target</p>
    <p><a href="#BM1">Jump</a></p>
  </body>
</html>`;

  const doc = await HtmlToWmlConverter.convertHtmlToWml("", "", "", xhtml, {});
  const mainXml = await doc.getPartText("/word/document.xml");
  assert.match(mainXml, /<w:bookmarkStart[^>]*w:name="BM1"/);
  assert.match(mainXml, /<w:bookmarkEnd[^>]*\/>/);
  assert.match(mainXml, /<w:hyperlink[^>]*w:anchor="BM1"/);

  const html = await WmlToHtmlConverter.convertToHtml(doc);
  assert.match(html.html, /<a id="BM1"\/?>/);
  assert.match(html.html, /<a href="#BM1">/);
});

test("HtmlToWmlConverter: honors <ol start> without affecting other lists", async () => {
  const xhtml = `<?xml version="1.0" encoding="UTF-8"?>
<html>
  <body>
    <ol start="5"><li>Five</li></ol>
    <ol><li>One</li></ol>
  </body>
</html>`;

  const doc = await HtmlToWmlConverter.convertHtmlToWml("", "", "", xhtml, {});
  const mainXml = await doc.getPartText("/word/document.xml");
  // Expect at least two distinct numIds used (separate list instances).
  assert.match(mainXml, /<w:numId[^>]*w:val="1"/);
  assert.match(mainXml, /<w:numId[^>]*w:val="2"/);

  const numberingXml = await doc.getPartText("/word/numbering.xml");
  assert.match(numberingXml, /<w:num w:numId="1">[\s\S]*<w:startOverride w:val="5"/);
  assert.doesNotMatch(numberingXml, /<w:num w:numId="2">[\s\S]*<w:startOverride/);

  const html = await WmlToHtmlConverter.convertToHtml(doc);
  assert.match(html.html, /<ol[^>]*start="5"/);
  assert.match(html.html, /Five/);
  assert.match(html.html, /One/);
});
