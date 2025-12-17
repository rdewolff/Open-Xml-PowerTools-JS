import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { buildDocx } from "./helpers/build-docx.js";

test("WmlToHtmlConverter: includeComments renders references + comment list", async () => {
  const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1">
    <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
  </w:comment>
</w:comments>
`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Body</w:t></w:r>
      <w:r><w:commentReference w:id="1"/></w:r>
    </w:p>
  </w:body>
</w:document>
`;

  const bytes = await buildDocx({
    documentXml,
    documentRelationships: [
      {
        Id: "rIdC1",
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        Target: "comments.xml",
      },
    ],
    contentTypes: {
      overrides: [
        {
          PartName: "/word/comments.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        },
      ],
    },
    extraEntries: [{ name: "word/comments.xml", text: commentsXml }],
  });

  const doc = new WmlDocument(bytes);
  const res = await WmlToHtmlConverter.convertToHtml(doc, { includeComments: true });

  assert.match(res.html, /href="#pt-comment-1"/);
  assert.match(res.html, /class="pt-comments"/);
  assert.match(res.html, /Comment text/);
});

