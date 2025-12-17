import { XmlDocument, XmlElement, XmlText, parseXml, serializeXml } from "./internal/xml.js";
import { ZipArchive } from "./internal/zip.js";
import { getDefaultZipAdapter } from "./internal/zip-adapter-auto.js";
import { WmlDocument } from "./wml-document.js";
import { OpenXmlPowerToolsError } from "./open-xml-powertools-error.js";
import { base64ToBytes } from "./util/base64.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";

export const HtmlToWmlConverter = {
  // Minimal v0 implementation: accepts well-formed XHTML-as-XML (string or XmlElement).
  // Ignores CSS and most settings; supports <p>, <br/>, <strong>/<b>, <em>/<i>, <u>,
  // plus <h1>-<h6>, <ul>/<ol>/<li>, <table>/<tr>/<td>, <a href>, <img src="data:...">.
  async convertHtmlToWml(defaultCss, authorCss, userCss, xhtml, settings = {}, templateDoc = null, annotatedHtmlDumpFileName = null) {
    void defaultCss;
    void authorCss;
    void userCss;
    void annotatedHtmlDumpFileName;
    const xhtmlDoc = coerceXhtml(xhtml);
    const body = findHtmlBody(xhtmlDoc.root) ?? xhtmlDoc.root;
    const ctx = new HtmlToWmlContext();
    const bodyBlocks = htmlToBlocks(body, ctx, 0);

    const wml = buildWmlDocumentXml(bodyBlocks, ctx);

    if (templateDoc) {
      return templateDoc.replacePartXml("/word/document.xml", wml);
    }

    const adapter = await getDefaultZipAdapter();
    const bytes = await buildMinimalDocxPackage(wml, ctx, adapter);
    return new WmlDocument(bytes, { fileName: settings?.fileName });
  },
};

class HtmlToWmlContext {
  constructor() {
    this.nextRelId = 10;
    this.relationships = [];
    this.hasNumbering = false;
    this.media = []; // { name, bytes, contentType }
  }

  addExternalHyperlink(href) {
    const id = `rId${this.nextRelId++}`;
    this.relationships.push({
      Id: id,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      Target: href,
      TargetMode: "External",
    });
    return id;
  }

  addImageFromDataUrl(dataUrl) {
    const parsed = parseDataUrl(dataUrl);
    if (!parsed) return null;
    const { contentType, bytes } = parsed;
    const ext = contentTypeToExtension(contentType);
    if (!ext) return null;

    const index = this.media.length + 1;
    const fileName = `image${index}.${ext}`;
    const partName = `word/media/${fileName}`;
    this.media.push({ name: partName, bytes, contentType });

    const id = `rId${this.nextRelId++}`;
    this.relationships.push({
      Id: id,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      Target: `media/${fileName}`,
    });
    return { relId: id, contentType };
  }
}

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

function htmlToBlocks(container, ctx, listLevel) {
  const out = [];
  for (const child of container.children ?? []) {
    if (child instanceof XmlText) continue;
    if (!(child instanceof XmlElement)) continue;
    const tag = child.qname.toLowerCase();

    if (tag === "p") {
      out.push(makeParagraph(htmlInlineToRuns(child, ctx, { bold: false, italic: false, underline: false })));
      continue;
    }

    if (/^h[1-6]$/.test(tag)) {
      const level = Number(tag.slice(1));
      out.push(makeHeading(level, htmlInlineToRuns(child, ctx, { bold: false, italic: false, underline: false })));
      continue;
    }

    if (tag === "ol" || tag === "ul") {
      ctx.hasNumbering = true;
      const isOrdered = tag === "ol";
      out.push(...listToParagraphs(child, ctx, isOrdered, listLevel));
      continue;
    }

    if (tag === "table") {
      out.push(makeTable(child, ctx));
      continue;
    }
  }

  if (!out.length) {
    // If no block children, treat as one paragraph.
    out.push(makeParagraph(htmlInlineToRuns(container, ctx, { bold: false, italic: false, underline: false })));
  }

  return out;
}

function listToParagraphs(listEl, ctx, isOrdered, listLevel) {
  const out = [];
  for (const li of listEl.children ?? []) {
    if (!(li instanceof XmlElement) || li.qname.toLowerCase() !== "li") continue;

    // Split li into inline content + nested lists/tables/paras.
    const inlineContainer = new XmlElement("span", new Map(), []);
    const nestedBlocks = [];
    for (const c of li.children ?? []) {
      if (c instanceof XmlText) {
        inlineContainer.children.push(c);
        continue;
      }
      if (!(c instanceof XmlElement)) continue;
      const t = c.qname.toLowerCase();
      if (t === "ol" || t === "ul" || t === "table" || t === "p" || /^h[1-6]$/.test(t)) {
        nestedBlocks.push(c);
      } else {
        inlineContainer.children.push(c);
      }
    }

    const runs = htmlInlineToRuns(inlineContainer, ctx, { bold: false, italic: false, underline: false });
    out.push(makeListParagraph(runs, isOrdered ? 1 : 2, listLevel));

    for (const nb of nestedBlocks) {
      const nbTag = nb.qname.toLowerCase();
      if (nbTag === "ol" || nbTag === "ul") {
        ctx.hasNumbering = true;
        out.push(...listToParagraphs(nb, ctx, nbTag === "ol", listLevel + 1));
      } else if (nbTag === "table") {
        out.push(makeTable(nb, ctx));
      } else if (nbTag === "p") {
        out.push(makeParagraph(htmlInlineToRuns(nb, ctx, { bold: false, italic: false, underline: false })));
      } else if (/^h[1-6]$/.test(nbTag)) {
        out.push(makeHeading(Number(nbTag.slice(1)), htmlInlineToRuns(nb, ctx, { bold: false, italic: false, underline: false })));
      }
    }
  }
  return out;
}

function htmlInlineToRuns(el, ctx, fmt) {
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
    if (tag === "a") {
      const href = child.attributes.get("href");
      const rid = href ? ctx.addExternalHyperlink(String(href)) : null;
      const inner = htmlInlineToRuns(child, ctx, fmt);
      runs.push(...(rid ? wrapHyperlinkRuns(rid, inner) : inner));
      continue;
    }
    if (tag === "img") {
      const src = child.attributes.get("src");
      const img = src ? ctx.addImageFromDataUrl(String(src)) : null;
      if (img) {
        const dims = parseHtmlImageDims(child);
        runs.push(makeImageRun(img.relId, dims));
      }
      continue;
    }
    if (tag === "strong" || tag === "b") {
      runs.push(...htmlInlineToRuns(child, ctx, { ...fmt, bold: true }));
      continue;
    }
    if (tag === "em" || tag === "i") {
      runs.push(...htmlInlineToRuns(child, ctx, { ...fmt, italic: true }));
      continue;
    }
    if (tag === "u") {
      runs.push(...htmlInlineToRuns(child, ctx, { ...fmt, underline: true }));
      continue;
    }
    // Unknown inline element: recurse without changing formatting.
    runs.push(...htmlInlineToRuns(child, ctx, fmt));
  }
  return runs;
}

function buildWmlDocumentXml(bodyBlocks, ctx) {
  const bodyChildren = [];
  for (const block of bodyBlocks) bodyChildren.push(block);
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
    new Map([
      ["xmlns:w", W_NS],
      ["xmlns:r", R_NS],
      ["xmlns:wp", WP_NS],
      ["xmlns:a", A_NS],
      ["xmlns:pic", PIC_NS],
    ]),
    [makeElement("w:body", new Map(), bodyChildren)],
  );
  return new XmlDocument(root);
}

function makeParagraph(runs) {
  return makeElement("w:p", new Map(), runs);
}

function makeHeading(level, runs) {
  const styleId = `Heading${level}`;
  const pPr = makeElement("w:pPr", new Map(), [makeElement("w:pStyle", new Map([["w:val", styleId]]), [])]);
  return makeElement("w:p", new Map(), [pPr, ...runs]);
}

function makeListParagraph(runs, numId, ilvl) {
  const pPr = makeElement("w:pPr", new Map(), [
    makeElement("w:numPr", new Map(), [
      makeElement("w:ilvl", new Map([["w:val", String(ilvl)]]), []),
      makeElement("w:numId", new Map([["w:val", String(numId)]]), []),
    ]),
  ]);
  return makeElement("w:p", new Map(), [pPr, ...runs]);
}

function makeTable(tableEl, ctx) {
  const htmlRows = [];
  for (const tr of tableEl.children ?? []) {
    if (!(tr instanceof XmlElement) || tr.qname.toLowerCase() !== "tr") continue;
    const cells = [];
    for (const td of tr.children ?? []) {
      if (!(td instanceof XmlElement)) continue;
      const tag = td.qname.toLowerCase();
      if (tag !== "td" && tag !== "th") continue;
      const colspan = parsePositiveInt(td.attributes.get("colspan")) ?? 1;
      const rowspan = parsePositiveInt(td.attributes.get("rowspan")) ?? 1;
      cells.push({ el: td, tag, colspan, rowspan });
    }
    htmlRows.push({ el: tr, cells });
  }

  // Build a rectangular Word table with explicit vMerge continuations.
  const pendingByCol = new Map(); // col -> span
  const rows = [];

  for (const r of htmlRows) {
    const trCells = [];
    let col = 0;

    const emitPendingAtCol = () => {
      const span = pendingByCol.get(col);
      if (!span) return false;
      if (span.startCol !== col) {
        col++;
        return true;
      }

      trCells.push(makeMergedContinuationCell(span.colspan));
      span.remaining--;
      if (span.remaining <= 0) {
        for (let k = 0; k < span.colspan; k++) pendingByCol.delete(span.startCol + k);
      }
      col += span.colspan;
      return true;
    };

    while (emitPendingAtCol()) {}

    for (const c of r.cells) {
      while (emitPendingAtCol()) {}

      const cellBlocks = htmlToBlocks(c.el, ctx, 0);
      const tcChildren = cellBlocks.length ? cellBlocks : [makeParagraph([makeRun("", { bold: false, italic: false, underline: false })])];
      const tcPrChildren = [];
      if (c.colspan > 1) tcPrChildren.push(makeElement("w:gridSpan", new Map([["w:val", String(c.colspan)]]), []));
      if (c.rowspan > 1) tcPrChildren.push(makeElement("w:vMerge", new Map([["w:val", "restart"]]), []));
      const tcChildrenWithPr = tcPrChildren.length ? [makeElement("w:tcPr", new Map(), tcPrChildren), ...tcChildren] : tcChildren;
      trCells.push(makeElement("w:tc", new Map(), tcChildrenWithPr));

      if (c.rowspan > 1) {
        const span = { startCol: col, colspan: c.colspan, remaining: c.rowspan - 1 };
        for (let k = 0; k < c.colspan; k++) pendingByCol.set(col + k, span);
      }

      col += c.colspan;
    }

    // Emit any remaining pending spans after explicit HTML cells.
    while (true) {
      let nextStart = null;
      for (const span of new Set(pendingByCol.values())) {
        if (span.startCol < col) continue;
        if (nextStart == null || span.startCol < nextStart) nextStart = span.startCol;
      }
      if (nextStart == null) break;
      col = nextStart;
      if (!emitPendingAtCol()) break;
      while (emitPendingAtCol()) {}
    }

    rows.push(makeElement("w:tr", new Map(), trCells));
  }

  return makeElement("w:tbl", new Map(), rows);
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

function wrapHyperlinkRuns(rid, runs) {
  return [
    makeElement("w:hyperlink", new Map([["r:id", rid]]), runs),
  ];
}

function makeImageRun(rid, dims) {
  // Minimal inline DrawingML template.
  // Default size is 1 inch square; if width/height are provided, use those.
  const emus = imageDimsToEmus(dims);
  const cx = String(emus?.cx ?? 914400); // 1 inch
  const cy = String(emus?.cy ?? 914400);
  const blip = makeElement("a:blip", new Map([["r:embed", rid]]), []);
  const blipFill = makeElement("pic:blipFill", new Map(), [blip]);
  const pic = makeElement("pic:pic", new Map(), [blipFill]);
  const graphicData = makeElement("a:graphicData", new Map([["uri", PIC_NS]]), [pic]);
  const graphic = makeElement("a:graphic", new Map(), [graphicData]);
  const extent = makeElement("wp:extent", new Map([["cx", cx], ["cy", cy]]), []);
  const inline = makeElement("wp:inline", new Map(), [extent, graphic]);
  const drawing = makeElement("w:drawing", new Map(), [inline]);
  return makeElement("w:r", new Map(), [drawing]);
}

function makeMergedContinuationCell(colspan) {
  const tcPrChildren = [makeElement("w:vMerge", new Map(), [])];
  if (colspan > 1) tcPrChildren.unshift(makeElement("w:gridSpan", new Map([["w:val", String(colspan)]]), []));
  const tcPr = makeElement("w:tcPr", new Map(), tcPrChildren);
  const empty = makeParagraph([makeRun("", { bold: false, italic: false, underline: false })]);
  return makeElement("w:tc", new Map(), [tcPr, empty]);
}

function parsePositiveInt(v) {
  if (v == null) return null;
  const s = String(v).trim();
  if (!s) return null;
  const n = Number.parseInt(s, 10);
  if (!Number.isFinite(n) || n <= 0) return null;
  return n;
}

function parseHtmlImageDims(imgEl) {
  const w = parsePositiveInt(imgEl.attributes.get("width"));
  const h = parsePositiveInt(imgEl.attributes.get("height"));
  if (w == null && h == null) return null;
  return { widthPx: w ?? null, heightPx: h ?? null };
}

function imageDimsToEmus(dims) {
  if (!dims) return null;
  const cx = dims.widthPx != null ? pxToEmus(dims.widthPx) : null;
  const cy = dims.heightPx != null ? pxToEmus(dims.heightPx) : null;
  if (cx == null && cy == null) return null;
  // If only one dimension provided, keep 1:1 to match default (simple behavior).
  const finalCx = cx ?? 914400;
  const finalCy = cy ?? finalCx;
  return { cx: finalCx, cy: finalCy };
}

function pxToEmus(px) {
  const n = Number(px);
  if (!Number.isFinite(n) || n <= 0) return null;
  // 1 inch == 914400 EMUs; assume 96 CSS px per inch.
  return Math.round((n * 914400) / 96);
}

function makeElement(qname, attributes, children) {
  return new XmlElement(qname, attributes, children);
}

async function buildMinimalDocxPackage(wmlXmlDoc, ctx, adapter) {
  const ctDefaults = [
    { Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" },
    { Extension: "xml", ContentType: "application/xml" },
  ];
  const ctOverrides = [
    { PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" },
    { PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" },
  ];
  if (ctx.hasNumbering) {
    ctOverrides.push({ PartName: "/word/numbering.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" });
  }
  for (const m of ctx.media) {
    const ext = m.name.split(".").pop().toLowerCase();
    if (!ctDefaults.some((d) => d.Extension === ext)) ctDefaults.push({ Extension: ext, ContentType: m.contentType });
  }

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
${ctDefaults.map((d) => `  <Default Extension="${d.Extension}" ContentType="${d.ContentType}"/>`).join("\n")}
${ctOverrides.map((o) => `  <Override PartName="${o.PartName}" ContentType="${o.ContentType}"/>`).join("\n")}
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
  ${ctx.hasNumbering ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : ""}
  ${ctx.relationships.map((r) => {
    const mode = r.TargetMode ? ` TargetMode="${r.TargetMode}"` : "";
    return `<Relationship Id="${r.Id}" Type="${r.Type}" Target="${r.Target}"${mode}/>`;
  }).join("\n  ")}
</Relationships>
`;

  const styles = buildStylesXml();

  const numbering = ctx.hasNumbering ? buildNumberingXml() : null;

  const documentXmlText = serializeXml(wmlXmlDoc, { xmlDeclaration: true });

  const enc = new TextEncoder();
  const entries = [
    { name: "[Content_Types].xml", bytes: enc.encode(contentTypes), compressionMethod: 8 },
    { name: "_rels/.rels", bytes: enc.encode(rootRels), compressionMethod: 8 },
    { name: "word/document.xml", bytes: enc.encode(documentXmlText), compressionMethod: 8 },
    { name: "word/_rels/document.xml.rels", bytes: enc.encode(docRels), compressionMethod: 8 },
    { name: "word/styles.xml", bytes: enc.encode(styles), compressionMethod: 8 },
  ];

  if (numbering) entries.push({ name: "word/numbering.xml", bytes: enc.encode(numbering), compressionMethod: 8 });
  for (const m of ctx.media) entries.push({ name: m.name, bytes: m.bytes, compressionMethod: 0 });

  return ZipArchive.build(entries, { adapter, level: 6 });
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  ${[1, 2, 3, 4, 5, 6].map((n) => `
  <w:style w:type="paragraph" w:styleId="Heading${n}">
    <w:name w:val="heading ${n}"/>
    <w:qFormat/>
    <w:rPr><w:b/><w:sz w:val="${Math.max(28, 48 - (n - 1) * 4)}"/></w:rPr>
  </w:style>`).join("\n")}
</w:styles>
`;
}

function buildNumberingXml() {
  // numId=1 decimal, numId=2 bullet; supports 9 levels each.
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="1">
    ${Array.from({ length: 9 }, (_, i) => `
    <w:lvl w:ilvl="${i}">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%${i + 1}."/>
    </w:lvl>`).join("\n")}
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>

  <w:abstractNum w:abstractNumId="2">
    ${Array.from({ length: 9 }, (_, i) => `
    <w:lvl w:ilvl="${i}">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
    </w:lvl>`).join("\n")}
  </w:abstractNum>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>
`;
}

function parseDataUrl(url) {
  const s = String(url);
  if (!s.startsWith("data:")) return null;
  const comma = s.indexOf(",");
  if (comma === -1) return null;
  const meta = s.slice(5, comma);
  const data = s.slice(comma + 1);
  const isBase64 = meta.includes(";base64");
  const contentType = meta.split(";")[0] || "application/octet-stream";
  if (!isBase64) return null;
  const bytes = base64ToBytes(data);
  return { contentType, bytes };
}

function contentTypeToExtension(ct) {
  const v = String(ct).toLowerCase();
  if (v === "image/png") return "png";
  if (v === "image/jpeg") return "jpg";
  if (v === "image/gif") return "gif";
  if (v === "image/webp") return "webp";
  return null;
}
