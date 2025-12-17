import { XmlDocument, XmlElement, XmlText, parseXml, serializeXml } from "./internal/xml.js";
import { ZipArchive } from "./internal/zip.js";
import { getDefaultZipAdapter } from "./internal/zip-adapter-auto.js";
import { WmlDocument } from "./wml-document.js";
import { OpenXmlPowerToolsError } from "./open-xml-powertools-error.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const HtmlToWmlConverter = {
  // Minimal v0 implementation: accepts well-formed XHTML-as-XML (string or XmlElement).
  // Ignores CSS and most settings; supports <p>, <br/>, <strong>/<b>, <em>/<i>, <u>.
  async convertHtmlToWml(defaultCss, authorCss, userCss, xhtml, settings = {}, templateDoc = null) {
    const xhtmlDoc = coerceXhtml(xhtml);
    const body = findHtmlBody(xhtmlDoc.root) ?? xhtmlDoc.root;
    const paragraphs = htmlToParagraphs(body);

    const wml = buildWmlDocumentXml(paragraphs);

    if (templateDoc) {
      return templateDoc.replacePartXml("/word/document.xml", wml);
    }

    const adapter = await getDefaultZipAdapter();
    const bytes = await buildMinimalDocxPackage(wml, adapter);
    return new WmlDocument(bytes, { fileName: settings?.fileName });
  },
};

function coerceXhtml(xhtml) {
  if (xhtml instanceof XmlDocument) return xhtml;
  if (xhtml instanceof XmlElement) return new XmlDocument(xhtml);
  if (typeof xhtml === "string") return parseXml(xhtml);
  throw new OpenXmlPowerToolsError(
    "OXPT_INVALID_ARGUMENT",
    "xhtml must be an XML string or an XmlElement/XmlDocument",
  );
}

function findHtmlBody(root) {
  if (root.qname === "body") return root;
  for (const el of root.descendants()) {
    if (el instanceof XmlElement && el.qname === "body") return el;
  }
  return null;
}

function htmlToParagraphs(container) {
  const out = [];
  const ps = [];
  for (const child of container.children ?? []) {
    if (child instanceof XmlElement && child.qname.toLowerCase() === "p") ps.push(child);
  }
  if (ps.length) {
    for (const p of ps) out.push(htmlInlineToRuns(p, { bold: false, italic: false, underline: false }));
    return out;
  }
  // If no <p>, treat the body as one paragraph.
  out.push(htmlInlineToRuns(container, { bold: false, italic: false, underline: false }));
  return out;
}

function htmlInlineToRuns(el, fmt) {
  const runs = [];
  for (const child of el.children ?? []) {
    if (child instanceof XmlText) {
      const text = child.text;
      if (!text) continue;
      runs.push(makeRun(text, fmt));
      continue;
    }
    if (!(child instanceof XmlElement)) continue;
    const tag = child.qname.toLowerCase();
    if (tag === "br") {
      runs.push(makeBreakRun());
      continue;
    }
    if (tag === "strong" || tag === "b") {
      runs.push(...htmlInlineToRuns(child, { ...fmt, bold: true }));
      continue;
    }
    if (tag === "em" || tag === "i") {
      runs.push(...htmlInlineToRuns(child, { ...fmt, italic: true }));
      continue;
    }
    if (tag === "u") {
      runs.push(...htmlInlineToRuns(child, { ...fmt, underline: true }));
      continue;
    }
    // Unknown inline element: recurse without changing formatting.
    runs.push(...htmlInlineToRuns(child, fmt));
  }
  return runs;
}

function buildWmlDocumentXml(paragraphRuns) {
  const bodyChildren = [];
  for (const runs of paragraphRuns) {
    bodyChildren.push(makeElement("w:p", new Map(), runs));
  }
  bodyChildren.push(
    makeElement("w:sectPr", new Map(), [
      makeElement("w:pgSz", new Map([["w:w", "12240"], ["w:h", "15840"]]), []),
      makeElement(
        "w:pgMar",
        new Map([
          ["w:top", "1440"],
          ["w:right", "1440"],
          ["w:bottom", "1440"],
          ["w:left", "1440"],
          ["w:header", "720"],
          ["w:footer", "720"],
          ["w:gutter", "0"],
        ]),
        [],
      ),
    ]),
  );

  const root = makeElement(
    "w:document",
    new Map([["xmlns:w", W_NS]]),
    [makeElement("w:body", new Map(), bodyChildren)],
  );
  return new XmlDocument(root);
}

function makeRun(text, fmt) {
  const rPrChildren = [];
  if (fmt.bold) rPrChildren.push(makeElement("w:b", new Map(), []));
  if (fmt.italic) rPrChildren.push(makeElement("w:i", new Map(), []));
  if (fmt.underline) rPrChildren.push(makeElement("w:u", new Map([["w:val", "single"]]), []));

  const rChildren = [];
  if (rPrChildren.length) rChildren.push(makeElement("w:rPr", new Map(), rPrChildren));

  const tAttrs = new Map();
  if (text.startsWith(" ") || text.endsWith(" ")) tAttrs.set("xml:space", "preserve");
  rChildren.push(makeElement("w:t", tAttrs, [new XmlText(text)]));
  return makeElement("w:r", new Map(), rChildren);
}

function makeBreakRun() {
  return makeElement("w:r", new Map(), [makeElement("w:br", new Map(), [])]);
}

function makeElement(qname, attributes, children) {
  return new XmlElement(qname, attributes, children);
}

async function buildMinimalDocxPackage(wmlXmlDoc, adapter) {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
`;

  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
`;

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
`;

  const enc = new TextEncoder();
  const documentXmlText = serializeXml(wmlXmlDoc, { xmlDeclaration: true });

  return ZipArchive.build(
    [
      { name: "[Content_Types].xml", bytes: enc.encode(contentTypes), compressionMethod: 8 },
      { name: "_rels/.rels", bytes: enc.encode(rootRels), compressionMethod: 8 },
      { name: "word/document.xml", bytes: enc.encode(documentXmlText), compressionMethod: 8 },
      { name: "word/_rels/document.xml.rels", bytes: enc.encode(docRels), compressionMethod: 8 },
      { name: "word/styles.xml", bytes: enc.encode(styles), compressionMethod: 8 },
    ],
    { adapter, level: 6 },
  );
}

