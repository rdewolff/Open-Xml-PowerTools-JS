import test from "node:test";
import assert from "node:assert/strict";
import { buildDocx } from "./helpers/build-docx.js";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";

test("WmlToHtmlConverter: resolves images in header/footer parts via their own .rels", async () => {
  // 1x1 transparent PNG
  const pngBytes = new Uint8Array(
    Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAOq6Qv4AAAAASUVORK5CYII=", "base64"),
  );

  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:p>
    <w:r>
      <w:drawing>
        <wp:inline>
          <wp:extent cx="9525" cy="9525"/>
          <wp:docPr id="1" name="HeaderImage" descr="HeaderImage"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:blipFill>
                  <a:blip r:embed="rIdImg1"/>
                </pic:blipFill>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>
</w:hdr>`;

  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:p>
    <w:r>
      <w:drawing>
        <wp:inline>
          <wp:extent cx="9525" cy="9525"/>
          <wp:docPr id="2" name="FooterImage" descr="FooterImage"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:blipFill>
                  <a:blip r:embed="rIdImg2"/>
                </pic:blipFill>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>
</w:ftr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader1"/>
      <w:footerReference w:type="default" r:id="rIdFooter1"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  const bytes = await buildDocx({
    documentXml,
    contentTypes: {
      defaults: [{ Extension: "png", ContentType: "image/png" }],
      overrides: [
        { PartName: "/word/header1.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" },
        { PartName: "/word/footer1.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml" },
      ],
    },
    documentRelationships: [
      { Id: "rIdHeader1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", Target: "header1.xml" },
      { Id: "rIdFooter1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", Target: "footer1.xml" },
    ],
    extraEntries: [
      { name: "word/header1.xml", text: headerXml, compressionMethod: 8 },
      {
        name: "word/_rels/header1.xml.rels",
        text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`,
        compressionMethod: 8,
      },
      { name: "word/footer1.xml", text: footerXml, compressionMethod: 8 },
      {
        name: "word/_rels/footer1.xml.rels",
        text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
</Relationships>`,
        compressionMethod: 8,
      },
      { name: "word/media/image1.png", bytes: pngBytes, compressionMethod: 0 },
      { name: "word/media/image2.png", bytes: pngBytes, compressionMethod: 0 },
    ],
});

test("WmlToHtmlConverter: prefers svgBlip over raster fallback in a:blip", async () => {
  // 1x1 transparent PNG (fallback)
  const pngBytes = new Uint8Array(
    Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAOq6Qv4AAAAASUVORK5CYII=", "base64"),
  );
  const svgText = `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1\" height=\"1\"></svg>`;
  const svgBytes = new TextEncoder().encode(svgText);

  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <w:p>
    <w:r>
      <w:drawing>
        <wp:inline>
          <wp:extent cx="9525" cy="9525"/>
          <wp:docPr id="1" name="SvgPreferred" descr="SvgPreferred"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:blipFill>
                  <a:blip r:embed="rIdPng">
                    <a:extLst>
                      <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
                        <asvg:svgBlip r:embed="rIdSvg"/>
                      </a:ext>
                    </a:extLst>
                  </a:blip>
                </pic:blipFill>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>
</w:hdr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader1"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  const bytes = await buildDocx({
    documentXml,
    contentTypes: {
      defaults: [{ Extension: "png", ContentType: "image/png" }],
      overrides: [
        { PartName: "/word/header1.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" },
        { PartName: "/word/media/image2.svg", ContentType: "image/svg+xml" },
      ],
    },
    documentRelationships: [
      { Id: "rIdHeader1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", Target: "header1.xml" },
    ],
    extraEntries: [
      { name: "word/header1.xml", text: headerXml, compressionMethod: 8 },
      {
        name: "word/_rels/header1.xml.rels",
        text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.svg"/>
  <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`,
        compressionMethod: 8,
      },
      { name: "word/media/image1.png", bytes: pngBytes, compressionMethod: 0 },
      { name: "word/media/image2.svg", bytes: svgBytes, compressionMethod: 0 },
    ],
  });

  const doc = WmlDocument.fromBytes(bytes, { fileName: "svg-preferred.docx" });
  const seen = [];
  await WmlToHtmlConverter.convertToHtml(doc, {
    additionalCss: "",
    imageHandler(info) {
      seen.push(info.contentType);
      return { src: "x" };
    },
  });

  assert.deepEqual(seen, ["image/svg+xml"]);
});

  const doc = WmlDocument.fromBytes(bytes, { fileName: "hf-images.docx" });

  const seen = [];
  const res = await WmlToHtmlConverter.convertToHtml(doc, {
    additionalCss: "",
    imageHandler(info) {
      seen.push({ contentType: info.contentType, bytes: info.bytes.length, altText: info.altText });
      return { src: "x" };
    },
  });

  assert.ok(res.html.includes("pt-header"), "Expected header wrapper");
  assert.ok(res.html.includes("pt-footer"), "Expected footer wrapper");
  assert.equal(seen.filter((s) => s.altText === "HeaderImage").length, 1);
  assert.equal(seen.filter((s) => s.altText === "FooterImage").length, 1);
  assert.equal(seen[0].contentType, "image/png");
});
