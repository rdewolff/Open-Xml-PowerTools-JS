import test from "node:test";
import assert from "node:assert/strict";
import { ZipArchive } from "../src/internal/zip.js";
import { ZipAdapterNode } from "../src/internal/zip-adapter-node.js";
import { OpenXmlPowerToolsDocument } from "../src/index.js";

test("OpenXmlPowerToolsDocument.detectType: returns opc for a valid OPC package without officeDocument rel", async () => {
  const enc = new TextEncoder();
  const bytes = await ZipArchive.build(
    [
      {
        name: "[Content_Types].xml",
        bytes: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`),
        compressionMethod: 8,
      },
      {
        name: "_rels/.rels",
        bytes: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`),
        compressionMethod: 8,
      },
    ],
    { adapter: ZipAdapterNode, level: 6 },
  );

  const doc = new OpenXmlPowerToolsDocument(bytes);
  assert.equal(await doc.detectType(), "opc");
});

