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

