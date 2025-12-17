import { ZipArchive } from "../../src/internal/zip.js";
import { ZipAdapterNode } from "../../src/internal/zip-adapter-node.js";

function buildContentTypesXml({ defaults = [], overrides = [] } = {}) {
  const baseDefaults = [
    { Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" },
    { Extension: "xml", ContentType: "application/xml" },
  ];
  const baseOverrides = [
    {
      PartName: "/word/document.xml",
      ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    },
    { PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" },
    { PartName: "/docProps/core.xml", ContentType: "application/vnd.openxmlformats-package.core-properties+xml" },
    { PartName: "/docProps/app.xml", ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml" },
  ];

  const allDefaults = [...baseDefaults, ...defaults];
  const allOverrides = [...baseOverrides, ...overrides];

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
${allDefaults.map((d) => `  <Default Extension="${d.Extension}" ContentType="${d.ContentType}"/>`).join("\n")}
${allOverrides.map((o) => `  <Override PartName="${o.PartName}" ContentType="${o.ContentType}"/>`).join("\n")}
</Types>
`;
}

const BASE_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
`;

function buildDocumentRelsXml(relationships = []) {
  const base = [
    {
      Id: "rId1",
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      Target: "styles.xml",
    },
  ];
  const all = [...base, ...relationships];
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${all.map((r) => {
    const mode = r.TargetMode ? ` TargetMode="${r.TargetMode}"` : "";
    return `  <Relationship Id="${r.Id}" Type="${r.Type}" Target="${r.Target}"${mode}/>`;
  }).join("\n")}
</Relationships>
`;
}

const BASE_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
`;

const BASE_CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Generated</dc:title>
</cp:coreProperties>
`;

const BASE_APP = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>OpenXmlPowerTools-JS</Application>
</Properties>
`;

export async function buildDocx({
  documentXml,
  contentTypes = {},
  documentRelationships = [],
  extraEntries = [],
} = {}) {
  const enc = new TextEncoder();
  const ctXml = buildContentTypesXml(contentTypes);
  const docRelsXml = buildDocumentRelsXml(documentRelationships);

  return ZipArchive.build(
    [
      { name: "[Content_Types].xml", bytes: enc.encode(ctXml), compressionMethod: 8 },
      { name: "_rels/.rels", bytes: enc.encode(BASE_ROOT_RELS), compressionMethod: 8 },
      { name: "word/document.xml", bytes: enc.encode(documentXml), compressionMethod: 8 },
      { name: "word/_rels/document.xml.rels", bytes: enc.encode(docRelsXml), compressionMethod: 8 },
      { name: "word/styles.xml", bytes: enc.encode(BASE_STYLES), compressionMethod: 8 },
      { name: "docProps/core.xml", bytes: enc.encode(BASE_CORE), compressionMethod: 8 },
      { name: "docProps/app.xml", bytes: enc.encode(BASE_APP), compressionMethod: 8 },
      ...extraEntries.map((e) => ({
        name: e.name,
        bytes: typeof e.text === "string" ? enc.encode(e.text) : e.bytes,
        compressionMethod: e.compressionMethod ?? 8,
      })),
    ],
    { adapter: ZipAdapterNode, level: 6 },
  );
}
