import { parseXml, XmlDocument, XmlElement, XmlText, serializeXml } from "./internal/xml.js";
import { bytesToBase64 } from "./util/base64.js";
import { readWmlStyles, parseParagraphProperties, parseRunProperties } from "./internal/wml-styles.js";
import { readWmlPartXml, findBody as findWBody, splitBodyIntoSections, selectHeaderFooterRefs, pickRef } from "./internal/wml-sections.js";
import {
  computeTableGrid,
  tableBordersToCss,
  tableWidthToCss,
  tableCellPaddingToCss,
  tableShadingToCss,
  rowShadingToCss,
  tableBidiVisual,
  cellBordersToCss,
  cellShadingToCss,
  cellVAlignToCss,
  cellPaddingToCss,
  cellWidthToCss,
} from "./internal/wml-tables.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const WmlToHtmlConverter = {
  async convertToHtml(doc, settings = {}) {
    const s = normalizeSettings(settings);
    // Preprocess like C# converter (minimal subset): accept revisions and simplify markup.
    // Do this in-memory for `/word/document.xml` so we don't rewrite the DOCX unless user requests.
    let xmlDoc = parseXml(await doc.getPartText("/word/document.xml"));
    if (s.preprocess.acceptRevisions) xmlDoc = acceptRevisionsXml(xmlDoc);
    if (s.preprocess.simplifyMarkup) xmlDoc = simplifyForHtmlXml(xmlDoc, { includeComments: s.includeComments });

    const warnings = [];

    const ctx = await WmlConversionContext.create(doc, xmlDoc, warnings, s);
    const body = ctx.findBody();
    const htmlParts = await renderBodyWithSections(ctx, body);
    const notesParts = await ctx.renderNotesSections();
    const commentsParts = await ctx.renderCommentsSection();

    const cssText = [ctx.generatedCssText, s.generalCss, s.additionalCss].filter(Boolean).join("\n");
    const htmlElement = buildHtmlElement({
      title: s.pageTitle,
      cssText,
      bodyHtml: [...htmlParts, ...notesParts, ...commentsParts].filter(Boolean).join("\n"),
    });
    const html = `<!doctype html>\n${serializeXml(htmlElement)}`;

    return s.output.format === "xml"
      ? { html, htmlElement, cssText, warnings }
      : { html, cssText, warnings };
  },
};

export const HtmlConverter = {
  convertToHtml: WmlToHtmlConverter.convertToHtml,
};

function normalizeSettings(settings) {
  return {
    pageTitle: settings.pageTitle ?? "",
    cssClassPrefix: settings.cssClassPrefix ?? "pt-",
    fabricateCssClasses: settings.fabricateCssClasses ?? true,
    generalCss: settings.generalCss ?? "span { white-space: pre-wrap; tab-size: 4; }",
    additionalCss: settings.additionalCss ?? "",
    restrictToSupportedLanguages: settings.restrictToSupportedLanguages ?? false,
    restrictToSupportedNumberingFormats: settings.restrictToSupportedNumberingFormats ?? false,
    listItemImplementations: settings.listItemImplementations ?? null,
    imageHandler: settings.imageHandler ?? null,
    includeComments: settings.includeComments ?? false,
    output: { format: settings.output?.format ?? "string" },
    preprocess: {
      acceptRevisions: settings.preprocess?.acceptRevisions ?? true,
      simplifyMarkup: settings.preprocess?.simplifyMarkup ?? true,
    },
  };
}

function buildHtmlElement({ title, cssText, bodyHtml }) {
  const headChildren = [
    new XmlElement("meta", new Map([["charset", "utf-8"]]), []),
    new XmlElement("title", new Map(), [new XmlText(title ?? "")]),
  ];
  if (cssText) headChildren.push(new XmlElement("style", new Map(), [new XmlText(cssText)]));

  // bodyHtml is already escaped HTML strings; wrap as raw text nodes is not correct XML.
  // For now, parse the body as XML fragments by wrapping in a dummy root.
  // This keeps `output.format: "xml"` useful without needing a full HTML parser.
  const bodyNodes = parseXml(`<root>${bodyHtml}</root>`).root.children.filter((n) => n instanceof XmlElement);
  const body = new XmlElement("body", new Map(), bodyNodes);
  const head = new XmlElement("head", new Map(), headChildren);
  return new XmlElement("html", new Map(), [head, body]);
}

async function renderBodyWithSections(ctx, body) {
  const sections = splitBodyIntoSections(body);
  if (!sections.length) return [];

  const prefix = ctx.settings.cssClassPrefix;
  const out = [];

  let lastHeaderRid = null;
  let lastFooterRid = null;
  let usedHeaderFooter = false;

  for (const s of sections) {
    const refs = s.sectPr ? selectHeaderFooterRefs(s.sectPr) : { header: null, footer: null };
    let headerRid = pickRef(refs.header);
    let footerRid = pickRef(refs.footer);

    // In practice, new sections typically re-use previous header/footer unless explicitly changed.
    if (!headerRid) headerRid = lastHeaderRid;
    if (!footerRid) footerRid = lastFooterRid;
    lastHeaderRid = headerRid;
    lastFooterRid = footerRid;

    const headerHtml = headerRid ? await ctx.renderHeaderFooterByRelId(headerRid) : "";
    const footerHtml = footerRid ? await ctx.renderHeaderFooterByRelId(footerRid) : "";
    const blocks = await renderBodyChildren(ctx, { children: s.nodes });

    const sectionParts = [];
    if (headerHtml) {
      usedHeaderFooter = true;
      sectionParts.push(`<div class="${escapeHtml(prefix)}header">${headerHtml}</div>`);
    }
    sectionParts.push(...blocks);
    if (footerHtml) {
      usedHeaderFooter = true;
      sectionParts.push(`<div class="${escapeHtml(prefix)}footer">${footerHtml}</div>`);
    }

    if (sectionParts.length) out.push(`<section class="${escapeHtml(prefix)}section">${sectionParts.join("\n")}</section>`);
  }

  if (usedHeaderFooter) ensureHeaderFooterCss(ctx);
  if (sections.length > 1) ensureSectionCss(ctx);
  return out;
}

function ensureHeaderFooterCss(ctx) {
  const prefix = ctx.settings.cssClassPrefix;
  if (ctx.generatedCssText.includes(`.${prefix}header`) || ctx.generatedCssText.includes(`.${prefix}footer`)) return;
  ctx.generatedCssText += `\n.${prefix}header,.${prefix}footer{color:inherit;opacity:0.85;font-size:0.9em}`;
  ctx.generatedCssText += `\n.${prefix}header{margin-bottom:1em}`;
  ctx.generatedCssText += `\n.${prefix}footer{margin-top:1em}`;
}

function ensureSectionCss(ctx) {
  const prefix = ctx.settings.cssClassPrefix;
  if (ctx.generatedCssText.includes(`.${prefix}section`)) return;
  ctx.generatedCssText += `\n.${prefix}section{margin-bottom:1.25em}`;
}

async function renderBodyChildren(ctx, body) {
  if (!body) return [];
  const out = [];
  const listStack = [];

  for (const child of body.children ?? []) {
    if (!(child instanceof XmlElement)) continue;
    if (isW(child, "p")) {
      const listInfo = ctx.getParagraphListInfo(child);
      if (listInfo) {
        ensureListStack(ctx, out, listStack, listInfo);
        const levelState = listStack[listStack.length - 1];
        if (levelState.numId !== listInfo.numId) {
          const start = Number(listInfo.start ?? 1);
          levelState.counter = Number.isFinite(start) && start > 1 ? start - 1 : 0;
          levelState.numId = listInfo.numId;
        }
        levelState.counter = (levelState.counter ?? 0) + 1;
        const marker = ctx.renderListMarker(listInfo, levelState.counter);
        const markerHtml = marker ? `<span class="${escapeHtml(marker.className)}">${escapeHtml(marker.text)}</span>` : "";
        out.push(`<li>${markerHtml}${await renderParagraphContents(ctx, child)}</li>`);
        continue;
      }

      closeAllLists(out, listStack);
      const headingLevel = ctx.getHeadingLevel(child);
      const tag = headingLevel ? `h${headingLevel}` : "p";
      const cls = ctx.getParagraphClass(child);
      const classAttr = cls ? ` class="${escapeHtml(cls)}"` : "";
      const styleAttr = ctx.getParagraphStyleAttr(child);
      const styleHtml = styleAttr ? ` style="${escapeHtml(styleAttr)}"` : "";
      const dirAttr = ctx.getParagraphDirAttr(child);
      out.push(`<${tag}${classAttr}${dirAttr}${styleHtml}>${await renderParagraphContents(ctx, child)}</${tag}>`);
      continue;
    }

    if (isW(child, "tbl")) {
      closeAllLists(out, listStack);
      out.push(await renderTable(ctx, child));
      continue;
    }
  }

  closeAllLists(out, listStack);
  return out;
}

async function renderParagraphContents(ctx, p) {
  const inner = [];
  for (const child of p.children ?? []) {
    if (!(child instanceof XmlElement)) continue;
    if (isW(child, "r")) {
      inner.push(await renderRun(ctx, p, child));
      continue;
    }
    if (isW(child, "hyperlink")) {
      inner.push(await renderHyperlink(ctx, child));
      continue;
    }
    if (isW(child, "fldSimple")) {
      // Render visible field result (runs inside), ignore instruction if present.
      const runs = [];
      for (const c of child.children ?? []) {
        if (c instanceof XmlElement && isW(c, "r")) runs.push(await renderRun(ctx, p, c));
      }
      inner.push(runs.join(""));
      continue;
    }
  }
  return inner.join("");
}

async function renderHyperlink(ctx, hyperlink) {
  const rid = hyperlink.attributes.get("r:id") ?? hyperlink.attributes.get("id");
  const anchor = hyperlink.attributes.get("w:anchor") ?? hyperlink.attributes.get("anchor");
  const href = rid ? ctx.getHyperlinkTarget(String(rid)) : null;
  const inner = [];
  for (const c of hyperlink.children ?? []) {
    if (!(c instanceof XmlElement)) continue;
    if (isW(c, "r")) inner.push(await renderRun(ctx, null, c));
  }
  const contents = inner.join("");
  if (href) return `<a href="${escapeHtml(href)}">${contents}</a>`;
  if (anchor) return `<a href="#${escapeHtml(String(anchor))}">${contents}</a>`;
  return contents;
}

async function renderRun(ctx, paragraph, r) {
  const effective = ctx.getEffectiveRunFormatting(paragraph, r);
  const bold = !!effective.bold;
  const italic = !!effective.italic;
  const underline = !!effective.underline;
  const strike = !!effective.strike;
  const vertAlign = effective.vertAlign ? String(effective.vertAlign) : null;
  const cls = ctx.getRunClass(r);

  const pieces = [];
  for (const child of r.children ?? []) {
    if (!child?.qname) continue;
    if (isW(child, "t")) pieces.push(escapeHtml(child.textContent()));
    else if (isW(child, "tab")) pieces.push("\t");
    else if (isW(child, "br")) pieces.push("<br/>");
    else if (isW(child, "cr")) pieces.push("<br/>");
    else if (isW(child, "noBreakHyphen")) pieces.push("‑");
    else if (isW(child, "softHyphen")) pieces.push("\u00ad");
    else if (isW(child, "footnoteReference")) {
      const id = child.attributes.get("w:id") ?? child.attributes.get("id");
      pieces.push(ctx.renderFootnoteRef(id));
    }
    else if (isW(child, "endnoteReference")) {
      const id = child.attributes.get("w:id") ?? child.attributes.get("id");
      pieces.push(ctx.renderEndnoteRef(id));
    }
    else if (isW(child, "commentReference")) {
      const id = child.attributes.get("w:id") ?? child.attributes.get("id");
      pieces.push(ctx.renderCommentRef(id));
    }
    else if (isW(child, "instrText") || isW(child, "fldChar")) {
      // Field code plumbing; ignore by default (visible text typically appears as w:t between separate/end).
      continue;
    }
    else if (isW(child, "drawing") || isW(child, "pict")) pieces.push(await renderImageFromContainer(ctx, child));
    else if (isW(child, "delText")) {
      // should be removed by RevisionAccepter; ignore for now
      continue;
    } else {
      ctx.warnings.push({
        code: "OXPT_HTML_UNSUPPORTED_RUN_CHILD",
        message: `Unsupported run child: ${child.qname}`,
        part: "/word/document.xml",
      });
    }
  }

  let html = pieces.join("");
  if (underline) html = `<u>${html}</u>`;
  if (italic) html = `<em>${html}</em>`;
  if (bold) html = `<strong>${html}</strong>`;
  if (strike) html = `<s>${html}</s>`;
  if (vertAlign === "superscript") html = `<sup>${html}</sup>`;
  if (vertAlign === "subscript") html = `<sub>${html}</sub>`;
  const styleAttr = ctx.getRunStyleAttr(effective);
  const classAttr = cls ? ` class="${escapeHtml(cls)}"` : "";
  const styleHtml = styleAttr ? ` style="${escapeHtml(styleAttr)}"` : "";
  const dirAttr = effective.rtl ? ` dir="rtl"` : "";
  return html ? `<span${classAttr}${dirAttr}${styleHtml}>${html}</span>` : "";
}

function findFirstByLocal(root, localName) {
  for (const d of root.descendants()) {
    if (d.nameParts().local === localName) return d;
  }
  return null;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

async function renderTable(ctx, tbl) {
  const tblPr = (tbl.children ?? []).find((c) => c instanceof XmlElement && isW(c, "tblPr")) ?? null;
  const borders = tableBordersToCss(tblPr);
  const style = [];
  if (borders?.top || borders?.left || borders?.bottom || borders?.right || borders?.insideH || borders?.insideV) style.push("border-collapse:collapse");
  if (borders?.top) style.push(`border-top:${borders.top}`);
  if (borders?.right) style.push(`border-right:${borders.right}`);
  if (borders?.bottom) style.push(`border-bottom:${borders.bottom}`);
  if (borders?.left) style.push(`border-left:${borders.left}`);
  const tblW = tableWidthToCss(tblPr);
  if (tblW) style.push(`width:${tblW}`);
  const tableAlign = tblPr ? tableAlignToCss(tblPr) : null;
  if (tableAlign) style.push(tableAlign);
  const tblBg = tableShadingToCss(tblPr);
  if (tblBg) style.push(`background-color:${tblBg}`);
  const defaultCellPad = tableCellPaddingToCss(tblPr);
  const dirAttr = tableBidiVisual(tblPr) ? ` dir="rtl"` : "";

  const grid = computeTableGrid(tbl);
  const theadRows = [];
  const tbodyRows = [];
  let isStillInHead = true;

  for (let rowIndex = 0; rowIndex < grid.length; rowIndex++) {
    const { tr, cells } = grid[rowIndex];
    const trPr = (tr.children ?? []).find((c) => c instanceof XmlElement && isW(c, "trPr")) ?? null;
    const isHeaderRow = !!(trPr && trPr.descendantsByNameNS(W_NS, "tblHeader").next().value);
    const rowBg = rowShadingToCss(trPr);
    const rowGoesToHead = isStillInHead && isHeaderRow;
    if (!rowGoesToHead) isStillInHead = false;
    const renderedCells = [];

    for (const cell of cells) {
      const tc = cell.tc;
      const tcPr = (tc.children ?? []).find((c) => c instanceof XmlElement && isW(c, "tcPr")) ?? null;
      const cellStyles = [];
      if (borders?.insideV && cell.col > 0) cellStyles.push(`border-left:${borders.insideV}`);
      if (borders?.insideH && rowIndex > 0) cellStyles.push(`border-top:${borders.insideH}`);
      const w = cellWidthToCss(tcPr);
      if (w) cellStyles.push(`width:${w}`);
      const cb = cellBordersToCss(tcPr);
      if (cb?.top) cellStyles.push(`border-top:${cb.top}`);
      if (cb?.right) cellStyles.push(`border-right:${cb.right}`);
      if (cb?.bottom) cellStyles.push(`border-bottom:${cb.bottom}`);
      if (cb?.left) cellStyles.push(`border-left:${cb.left}`);
      const bg = cellShadingToCss(tcPr) ?? rowBg;
      if (bg) cellStyles.push(`background-color:${bg}`);
      const va = cellVAlignToCss(tcPr);
      if (va) cellStyles.push(`vertical-align:${va}`);
      const pad = cellPaddingToCss(tcPr) ?? defaultCellPad;
      if (pad) cellStyles.push(pad);

      const blocks = await renderTableCellBlocks(ctx, tc);

      const colspanAttr = cell.colspan > 1 ? ` colspan="${cell.colspan}"` : "";
      const rowspanAttr = cell.rowspan > 1 ? ` rowspan="${cell.rowspan}"` : "";
      const styleAttr = cellStyles.length ? ` style="${escapeHtml(cellStyles.join(";"))}"` : "";
      const cellTag = isHeaderRow ? "th" : "td";
      renderedCells.push(`<${cellTag}${colspanAttr}${rowspanAttr}${styleAttr}>${blocks.join("")}</${cellTag}>`);
    }

    const rowHtml = `<tr>${renderedCells.join("")}</tr>`;
    (rowGoesToHead ? theadRows : tbodyRows).push(rowHtml);
  }

  const tableStyle = style.length ? ` style="${escapeHtml(style.join(";"))}"` : "";
  const theadHtml = theadRows.length ? `<thead>${theadRows.join("")}</thead>` : "";
  const tbodyHtml = `<tbody>${tbodyRows.join("")}</tbody>`;
  return `<table${dirAttr}${tableStyle}>${theadHtml}${tbodyHtml}</table>`;
}

async function renderTableCellBlocks(ctx, tc) {
  // Render only direct blocks (paragraphs/tables) to avoid duplicating nested content.
  const out = [];
  for (const child of tc.children ?? []) {
    if (!(child instanceof XmlElement)) continue;
    if (isW(child, "p")) {
      out.push(`<p>${await renderParagraphContents(ctx, child)}</p>`);
      continue;
    }
    if (isW(child, "tbl")) {
      out.push(await renderTable(ctx, child));
      continue;
    }
  }
  return out;
}

async function renderBlocksFromContainer(ctx, container) {
  // Render with the same block rules as the main body, but scoped to this container.
  return renderBodyChildren(ctx, container);
}

function tableAlignToCss(tblPr) {
  const jc = (tblPr.children ?? []).find((c) => c instanceof XmlElement && isW(c, "jc")) ?? null;
  const val = jc?.attributes.get("w:val") ?? jc?.attributes.get("val") ?? null;
  if (!val) return null;
  const v = String(val);
  if (v === "center") return "margin-left:auto;margin-right:auto";
  if (v === "right") return "margin-left:auto";
  return null;
}

function ensureListStack(ctx, out, listStack, listInfo) {
  // listInfo.level is 0-based.
  while (listStack.length > listInfo.level + 1) {
    out.push(`</${listStack.pop().tag}>`);
  }
  while (listStack.length < listInfo.level + 1) {
    const { tag, attrs } = listInfo;
    out.push(`<${tag}${renderAttrs(attrs)}>`);
    const start = Number(listInfo.start ?? 1);
    const counter = Number.isFinite(start) && start > 1 ? start - 1 : 0;
    listStack.push({ tag, attrs, counter, numId: listInfo.numId, listInfo });
  }
  const current = listStack[listStack.length - 1];
  if (current.tag !== listInfo.tag) {
    out.push(`</${listStack.pop().tag}>`);
    out.push(`<${listInfo.tag}${renderAttrs(listInfo.attrs)}>`);
    const start = Number(listInfo.start ?? 1);
    const counter = Number.isFinite(start) && start > 1 ? start - 1 : 0;
    listStack.push({ tag: listInfo.tag, attrs: listInfo.attrs, counter, numId: listInfo.numId, listInfo });
  } else {
    current.listInfo = listInfo;
  }
}

function closeAllLists(out, listStack) {
  while (listStack.length) out.push(`</${listStack.pop().tag}>`);
}

async function renderImageFromContainer(ctx, container) {
  const blip = findFirstByLocal(container, "blip");
  if (blip instanceof XmlElement) {
    const embed = blip.attributes.get("r:embed") ?? blip.attributes.get("embed");
    if (embed) return await ctx.renderImage(String(embed), container);
  }
  const imagedata = findFirstByLocal(container, "imagedata");
  if (imagedata instanceof XmlElement) {
    const rid = imagedata.attributes.get("r:id") ?? imagedata.attributes.get("id");
    if (rid) return await ctx.renderImage(String(rid), container);
  }
  return "";
}

function acceptRevisionsXml(xmlDoc) {
  const root = acceptTransform(xmlDoc.root);
  return new XmlDocument(root);
}

function acceptTransform(node) {
  if (node instanceof XmlText) return new XmlText(node.text);
  if (!(node instanceof XmlElement)) return null;
  if (isW(node)) {
    const local = node.nameParts().local;
    if (local === "ins" || local === "moveTo") return acceptChildren(node);
    if (local === "del" || local === "moveFrom") return null;
    if (
      local === "delText" ||
      local === "delInstrText" ||
      local === "moveFromRangeStart" ||
      local === "moveFromRangeEnd" ||
      local === "moveToRangeStart" ||
      local === "moveToRangeEnd"
    ) {
      return null;
    }
  }
  const children = [];
  for (const c of node.children ?? []) {
    const t = acceptTransform(c);
    if (!t) continue;
    if (t instanceof XmlElement && t.qname === "_oxptjs_unwrap_") {
      for (const gc of t.children) children.push(gc);
      continue;
    }
    children.push(t);
  }
  return new XmlElement(node.qname, new Map(node.attributes), children);
}

function acceptChildren(el) {
  const children = [];
  for (const c of el.children ?? []) {
    const t = acceptTransform(c);
    if (t) children.push(t);
  }
  // unwrap by returning a synthetic container; caller uses children directly
  if (!children.length) return null;
  // if acceptChildren is called, caller expects nodes; but our acceptTransform returns node
  // so acceptTransform special-cases by returning first child in wrapper cases.
  // Here, return a dummy element isn't correct; instead, return children by splicing at call site.
  // We implement wrappers by calling acceptChildren and returning a marker object.
  return new XmlElement("_oxptjs_unwrap_", new Map(), children);
}

function simplifyForHtmlXml(xmlDoc, { includeComments = false } = {}) {
  // Roughly matches the C# converter's SimplifyMarkupSettings used for HTML conversion.
  const settings = {
    removeComments: !includeComments,
    removeContentControls: true,
    removeLastRenderedPageBreak: true,
    removePermissions: true,
    removeProof: true,
    removeRsidInfo: true,
    removeSmartTags: true,
    removeSoftHyphens: true,
    removeGoBackBookmark: true,
    replaceTabsWithSpaces: false,
  };
  const root = simplifyTransform(xmlDoc.root, settings);
  return new XmlDocument(root);
}

function preprocessXmlForHtml(xmlDoc, settings) {
  let x = xmlDoc;
  if (settings?.preprocess?.acceptRevisions) x = acceptRevisionsXml(x);
  if (settings?.preprocess?.simplifyMarkup) x = simplifyForHtmlXml(x, { includeComments: settings.includeComments });
  return x;
}

function simplifyTransform(node, settings) {
  if (node instanceof XmlText) return new XmlText(node.text);
  if (!(node instanceof XmlElement)) return null;

  if (isW(node)) {
    const local = node.nameParts().local;
    if (settings.removeSmartTags && local === "smartTag") return simplifyChildren(node, settings);
    if (settings.removeContentControls && local === "sdt") {
      const sdtContent = node.children.find((c) => c instanceof XmlElement && isW(c, "sdtContent"));
      return sdtContent ? simplifyChildren(sdtContent, settings) : null;
    }
    if (settings.removeLastRenderedPageBreak && local === "lastRenderedPageBreak") return null;
    if (
      settings.removeComments &&
      (local === "commentRangeStart" || local === "commentRangeEnd" || local === "commentReference")
    ) {
      return null;
    }
    if (settings.removePermissions && (local === "permStart" || local === "permEnd")) return null;
    if (settings.removeProof && local === "proofErr") return null;
    if (settings.removeSoftHyphens && local === "t") {
      const text = node.textContent().replaceAll("\u00ad", "");
      return new XmlElement(node.qname, filteredAttributes(node.attributes, settings), [new XmlText(text)]);
    }
    if (settings.removeGoBackBookmark && (local === "bookmarkStart" || local === "bookmarkEnd")) {
      const name = node.attributes.get("w:name") ?? node.attributes.get("name");
      if (local === "bookmarkStart" && name === "_GoBack") return null;
      // for bookmarkEnd, we can't easily match id here; handled below by generic removal
    }
  }

  const children = [];
  for (const c of node.children ?? []) {
    const t = simplifyTransform(c, settings);
    if (!t) continue;
    if (t instanceof XmlElement && t.qname === "_oxptjs_unwrap_") {
      for (const gc of t.children) children.push(gc);
      continue;
    }
    children.push(t);
  }
  return new XmlElement(node.qname, filteredAttributes(node.attributes, settings), children);
}

function simplifyChildren(el, settings) {
  const children = [];
  for (const c of el.children ?? []) {
    const t = simplifyTransform(c, settings);
    if (!t) continue;
    if (t instanceof XmlElement && t.qname === "_oxptjs_unwrap_") {
      for (const gc of t.children) children.push(gc);
      continue;
    }
    children.push(t);
  }
  return new XmlElement("_oxptjs_unwrap_", new Map(), children);
}

function filteredAttributes(attributes, settings) {
  if (!settings.removeRsidInfo) return new Map(attributes);
  const out = new Map();
  for (const [k, v] of attributes) {
    const local = k.includes(":") ? k.split(":")[1] : k;
    if (local.startsWith("rsid")) continue;
    out.set(k, v);
  }
  return out;
}

class WmlConversionContext {
  constructor({ doc, mainXml, warnings, settings, rels, numbering, contentTypes, styles }) {
    this.doc = doc;
    this.mainXml = mainXml;
    this.warnings = warnings;
    this.settings = settings;
    this.rels = rels;
    this.numbering = numbering;
    this.contentTypes = contentTypes;
    this.styles = styles;
    this.generatedCssText = "";
    this.footnotes = null;
    this.endnotes = null;
    this.comments = null;
    this._headerFooterByRid = new Map();
  }

  static async create(doc, mainXml, warnings, settings) {
    const [rels, numbering, contentTypes, styles, footnotes, endnotes, comments] = await Promise.all([
      readRelationships(doc, "/word/_rels/document.xml.rels"),
      readNumbering(doc),
      readContentTypes(doc),
      readWmlStyles(doc),
      readNotes(doc, "/word/footnotes.xml", settings),
      readNotes(doc, "/word/endnotes.xml", settings),
      settings.includeComments ? readComments(doc, "/word/comments.xml", settings) : Promise.resolve(null),
    ]);
    const ctx = new WmlConversionContext({ doc, mainXml, warnings, settings, rels, numbering, contentTypes, styles });
    ctx.footnotes = footnotes;
    ctx.endnotes = endnotes;
    ctx.comments = comments;
    ctx._documentLang = detectDocumentLanguage(mainXml.root);
    return ctx;
  }

  findBody() {
    return findWBody(this.mainXml.root);
  }

  getHeadingLevel(p) {
    const pPr = (p.children ?? []).find((c) => c instanceof XmlElement && isW(c, "pPr"));
    const pStyle = pPr?.children?.find((c) => c instanceof XmlElement && isW(c, "pStyle"));
    const val = pStyle?.attributes?.get("w:val") ?? pStyle?.attributes?.get("val");
    if (!val) return null;
    const m = String(val).match(/^Heading([1-6])$/i);
    return m ? Number(m[1]) : null;
  }

  getParagraphListInfo(p) {
    const eff = this.getEffectiveParagraphProperties(p);
    const numPr = eff.numPr;
    if (!numPr) return null;
    const level = Number(numPr.ilvl ?? 0);
    const numId = String(numPr.numId ?? "");
    if (!numId) return null;
    const lvl = this.numbering?.getLevel(numId, level);
    const numFmt = lvl?.numFmt ?? "decimal";
    const lvlText = lvl?.lvlText ?? "";
    const start = lvl?.start ?? 1;
    const tag = toListTag(numFmt, this.settings.restrictToSupportedNumberingFormats, this.warnings);
    const attrs = toListAttrs(numFmt);
    if (tag === "ol") {
      const n = Number(start);
      if (Number.isFinite(n) && n > 1) attrs.start = String(n);
    }
    return { numId, level, numFmt, lvlText, start, tag, attrs };
  }

  getParagraphClass(p) {
    if (!this.settings.fabricateCssClasses) return null;
    const eff = this.getEffectiveParagraphProperties(p);
    const val = eff.pStyle;
    if (!val) return null;
    this.ensureParagraphStyleCss(String(val));
    return `${this.settings.cssClassPrefix}p-${slug(String(val))}`;
  }

  getRunClass(r) {
    if (!this.settings.fabricateCssClasses) return null;
    const rPr = (r.children ?? []).find((c) => c instanceof XmlElement && isW(c, "rPr"));
    const rStyle = rPr?.children?.find((c) => c instanceof XmlElement && isW(c, "rStyle"));
    const val = rStyle?.attributes?.get("w:val") ?? rStyle?.attributes?.get("val");
    if (!val) return null;
    this.ensureRunStyleCss(String(val));
    return `${this.settings.cssClassPrefix}r-${slug(String(val))}`;
  }

  getParagraphStyleAttr(p) {
    const eff = this.getEffectiveParagraphProperties(p);
    const rules = [];
    if (eff.jc) {
      const ta = mapJustificationToCssTextAlign(eff.jc);
      if (ta) rules.push(`text-align:${ta}`);
    }
    if (eff.bidi) rules.push("direction:rtl");
    if (eff.before != null) rules.push(`margin-top:${twipsToPt(eff.before)}pt`);
    if (eff.after != null) rules.push(`margin-bottom:${twipsToPt(eff.after)}pt`);
    if (eff.left != null) rules.push(`margin-left:${twipsToPt(eff.left)}pt`);
    if (eff.right != null) rules.push(`margin-right:${twipsToPt(eff.right)}pt`);
    if (eff.firstLine != null) rules.push(`text-indent:${twipsToPt(eff.firstLine)}pt`);
    if (eff.hanging != null) rules.push(`text-indent:-${twipsToPt(eff.hanging)}pt`);
    const lh = lineHeightCss(eff.line, eff.lineRule);
    if (lh) rules.push(lh);
    return rules.length ? rules.join(";") : "";
  }

  getParagraphDirAttr(p) {
    const eff = this.getEffectiveParagraphProperties(p);
    return eff.bidi ? ` dir="rtl"` : "";
  }

  getEffectiveParagraphProperties(p) {
    const pPr = (p.children ?? []).find((c) => c instanceof XmlElement && isW(c, "pPr")) ?? null;
    const direct = pPr ? parseParagraphProperties(pPr) : {};
    const styleId = direct.pStyle ?? null;
    const style = styleId ? this.styles.resolveParagraphStyle(styleId) : null;
    const merged = { ...(style?.para ?? {}), ...direct };
    if (styleId) merged.pStyle = styleId;
    return merged;
  }

  getEffectiveRunFormatting(paragraph, r) {
    const rPr = (r.children ?? []).find((c) => c instanceof XmlElement && isW(c, "rPr")) ?? null;
    const direct = rPr ? parseRunProperties(rPr) : {};
    const runStyle = direct.rStyle ? this.styles.resolveCharacterStyle(direct.rStyle) : null;

    // paragraph style rPr applies to runs as a base
    let paragraphRun = {};
    if (paragraph) {
      const pEff = this.getEffectiveParagraphProperties(paragraph);
      if (pEff.pStyle) {
        const pStyle = this.styles.resolveParagraphStyle(pEff.pStyle);
        paragraphRun = pStyle?.run ?? {};
      }
    }

    return { ...paragraphRun, ...(runStyle?.run ?? {}), ...direct };
  }

  getRunStyleAttr(eff) {
    const rules = [];
    if (eff.color) rules.push(`color:${eff.color}`);
    if (eff.sz != null) rules.push(`font-size:${halfPointsToPt(eff.sz)}pt`);
    if (eff.fontFamily) rules.push(`font-family:${cssString(eff.fontFamily)}`);
    if (eff.rtl) rules.push("direction:rtl;unicode-bidi:isolate");
    if (eff.smallCaps) rules.push("font-variant:small-caps");
    if (eff.caps) rules.push("text-transform:uppercase");
    if (eff.highlight) {
      const bg = mapHighlightToCss(eff.highlight);
      if (bg) rules.push(`background-color:${bg}`);
    }
    return rules.length ? rules.join(";") : "";
  }

  ensureParagraphStyleCss(styleId) {
    const className = `${this.settings.cssClassPrefix}p-${slug(styleId)}`;
    if (this.generatedCssText.includes(`.${className}`)) return;
    const style = this.styles.resolveParagraphStyle(styleId);
    if (!style) return;
    const rules = [];
    if (style.para.jc) {
      const ta = mapJustificationToCssTextAlign(style.para.jc);
      if (ta) rules.push(`text-align:${ta}`);
    }
    if (style.para.bidi) rules.push("direction:rtl");
    if (style.para.before != null) rules.push(`margin-top:${twipsToPt(style.para.before)}pt`);
    if (style.para.after != null) rules.push(`margin-bottom:${twipsToPt(style.para.after)}pt`);
    if (style.para.left != null) rules.push(`margin-left:${twipsToPt(style.para.left)}pt`);
    if (style.para.right != null) rules.push(`margin-right:${twipsToPt(style.para.right)}pt`);
    if (style.para.firstLine != null) rules.push(`text-indent:${twipsToPt(style.para.firstLine)}pt`);
    if (style.para.hanging != null) rules.push(`text-indent:-${twipsToPt(style.para.hanging)}pt`);
    const lh = lineHeightCss(style.para.line, style.para.lineRule);
    if (lh) rules.push(lh);
    if (rules.length) this.generatedCssText += `\n.${className}{${rules.join(";")}}`;
  }

  ensureRunStyleCss(styleId) {
    const className = `${this.settings.cssClassPrefix}r-${slug(styleId)}`;
    if (this.generatedCssText.includes(`.${className}`)) return;
    const style = this.styles.resolveCharacterStyle(styleId);
    if (!style) return;
    const eff = style.run ?? {};
    const rules = [];
    if (eff.color) rules.push(`color:${eff.color}`);
    if (eff.sz != null) rules.push(`font-size:${halfPointsToPt(eff.sz)}pt`);
    if (eff.fontFamily) rules.push(`font-family:${cssString(eff.fontFamily)}`);
    if (eff.smallCaps) rules.push("font-variant:small-caps");
    if (eff.caps) rules.push("text-transform:uppercase");
    if (eff.highlight) {
      const bg = mapHighlightToCss(eff.highlight);
      if (bg) rules.push(`background-color:${bg}`);
    }
    if (eff.bold) rules.push("font-weight:700");
    if (eff.italic) rules.push("font-style:italic");
    if (eff.underline) rules.push("text-decoration:underline");
    if (rules.length) this.generatedCssText += `\n.${className}{${rules.join(";")}}`;
  }

  getHyperlinkTarget(rid) {
    const rel = this.rels?.byId.get(rid);
    if (!rel) return null;
    if (rel.targetMode === "External") return rel.target;
    // Internal part link; not generally meaningful in HTML.
    return rel.target;
  }

  async renderImage(rid, drawingElement) {
    const rel = this.rels?.byId.get(rid);
    if (!rel) return "";
    const partUri = resolveWordTarget(rel.target);
    let bytes;
    try {
      bytes = await this.doc.getPartBytes(partUri);
    } catch (e) {
      this.warnings.push({ code: "OXPT_IMAGE_MISSING_PART", message: `Missing image part: ${partUri}`, part: partUri });
      return "";
    }
    const contentType = this.contentTypes?.getContentType(partUri) ?? "application/octet-stream";
    const info = {
      contentType,
      bytes,
      altText: extractAltText(drawingElement),
      widthEmus: extractExtent(drawingElement)?.cx ?? undefined,
      heightEmus: extractExtent(drawingElement)?.cy ?? undefined,
      suggestedStyle: null,
    };
    info.suggestedStyle = suggestedImgStyle(info);

    if (this.settings.imageHandler) {
      const res = this.settings.imageHandler(info);
      if (!res) return "";
      if (res.element) return res.element;
      const attrs = res.attrs ?? {};
      if (info.suggestedStyle && attrs.style == null) attrs.style = info.suggestedStyle;
      return `<img src="${escapeHtml(res.src)}"${renderAttrs(attrs)}${renderAlt(info.altText)}/>`;
    }

    const b64 = bytesToBase64(bytes);
    const src = `data:${contentType};base64,${b64}`;
    const styleAttr = info.suggestedStyle ? ` style="${escapeHtml(info.suggestedStyle)}"` : "";
    return `<img src="${src}"${styleAttr}${renderAlt(info.altText)}/>`;
  }

  renderFootnoteRef(id) {
    const noteId = id == null ? "" : String(id);
    if (!noteId) return "";
    if (!this.footnotes?.byId.has(noteId)) return `<sup>[${escapeHtml(noteId)}]</sup>`;
    return `<sup><a href="#${escapeHtml(this.settings.cssClassPrefix)}footnote-${escapeHtml(noteId)}">${escapeHtml(noteId)}</a></sup>`;
  }

  renderEndnoteRef(id) {
    const noteId = id == null ? "" : String(id);
    if (!noteId) return "";
    if (!this.endnotes?.byId.has(noteId)) return `<sup>[${escapeHtml(noteId)}]</sup>`;
    return `<sup><a href="#${escapeHtml(this.settings.cssClassPrefix)}endnote-${escapeHtml(noteId)}">${escapeHtml(noteId)}</a></sup>`;
  }

  async renderNotesSections() {
    const out = [];
    const footnotesHtml = await this.renderNotesSection("footnote", this.footnotes);
    if (footnotesHtml) out.push(footnotesHtml);
    const endnotesHtml = await this.renderNotesSection("endnote", this.endnotes);
    if (endnotesHtml) out.push(endnotesHtml);
    return out;
  }

  async renderHeaderFooterByRelId(rid) {
    const cacheKey = String(rid);
    if (this._headerFooterByRid.has(cacheKey)) return this._headerFooterByRid.get(cacheKey);
    const rel = this.rels?.byId.get(String(rid));
    if (!rel) {
      this.warnings.push({ code: "OXPT_HF_MISSING_REL", message: `Missing header/footer relationship: ${rid}`, part: "/word/_rels/document.xml.rels" });
      this._headerFooterByRid.set(cacheKey, "");
      return "";
    }
    const partUri = resolveWordTarget(rel.target);
    const xml = await readWmlPartXml(this.doc, partUri);
    if (!xml) {
      this.warnings.push({ code: "OXPT_HF_MISSING_PART", message: `Missing header/footer part: ${partUri}`, part: partUri });
      this._headerFooterByRid.set(cacheKey, "");
      return "";
    }
    const processed = preprocessXmlForHtml(xml, this.settings);
    const body = findWBody(processed.root) ?? processed.root;
    const blocks = await renderBodyChildren(this, body);
    const html = blocks.join("");
    this._headerFooterByRid.set(cacheKey, html);
    return html;
  }

  renderCommentRef(id) {
    if (!this.settings.includeComments) return "";
    const cid = id == null ? "" : String(id);
    if (!cid) return "";
    if (!this.comments?.byId.has(cid)) return `<sup>[c${escapeHtml(cid)}]</sup>`;
    return `<sup><a href="#${escapeHtml(this.settings.cssClassPrefix)}comment-${escapeHtml(cid)}">c${escapeHtml(cid)}</a></sup>`;
  }

  async renderCommentsSection() {
    if (!this.settings.includeComments) return [];
    if (!this.comments?.orderedIds?.length) return [];
    const items = [];
    for (const id of this.comments.orderedIds) {
      const el = this.comments.byId.get(id) ?? null;
      const html = el ? (await renderBlocksFromContainer(this, el)).join("") : "";
      items.push(
        `<li id="${escapeHtml(this.settings.cssClassPrefix)}comment-${escapeHtml(id)}">${html}</li>`,
      );
    }
    if (!items.length) return [];
    this.generatedCssText += `\n.${this.settings.cssClassPrefix}comments{margin-top:1em;font-size:0.95em}`;
    return [`<hr/><ol class="${escapeHtml(this.settings.cssClassPrefix)}comments">${items.join("")}</ol>`];
  }

  async renderNotesSection(kind, notes) {
    if (!notes?.orderedIds?.length) return "";
    const items = [];
    for (const id of notes.orderedIds) {
      const el = notes.byId.get(id) ?? null;
      const html = el ? (await renderBlocksFromContainer(this, el)).join("") : "";
      items.push(
        `<li id="${escapeHtml(this.settings.cssClassPrefix)}${kind}-${escapeHtml(id)}">${html}</li>`,
      );
    }
    if (!items.length) return "";
    this.generatedCssText += `\n.${this.settings.cssClassPrefix}${kind}s{margin-top:1em;font-size:0.95em}`;
    return `<hr/><ol class="${escapeHtml(this.settings.cssClassPrefix)}${kind}s">${items.join("")}</ol>`;
  }

  renderListMarker(listInfo, index1Based) {
    const impl = this.getListItemTextImplementation();
    if (!impl) return null;
    this.ensureGeneratedMarkerCss();
    const levelText = listInfo.lvlText ?? "";
    const text = impl(levelText, index1Based, listInfo.numFmt);
    if (text == null) return null;
    return { className: `${this.settings.cssClassPrefix}li-marker`, text: String(text) };
  }

  getListItemTextImplementation() {
    const impls = this.settings.listItemImplementations;
    if (!impls) return null;
    const lang = this._documentLang;
    if (lang && typeof impls[lang] === "function") return impls[lang];

    if (lang && this.settings.restrictToSupportedLanguages) {
      const hasFallback =
        typeof impls.default === "function" ||
        typeof impls["en-US"] === "function" ||
        Object.values(impls).some((v) => typeof v === "function");
      if (!hasFallback) {
        this.warnings.push({
          code: "OXPT_LIST_LANG_UNSUPPORTED",
          message: `No listItemImplementations for document language '${lang}' (and no fallback)`,
          part: "/word/document.xml",
        });
        return null;
      }

      // Warn but still fall back to keep output usable.
      this.warnings.push({
        code: "OXPT_LIST_LANG_UNSUPPORTED",
        message: `No listItemImplementations for document language '${lang}', using fallback`,
        part: "/word/document.xml",
      });
    }

    if (typeof impls.default === "function") return impls.default;
    if (typeof impls["en-US"] === "function") return impls["en-US"];
    for (const v of Object.values(impls)) {
      if (typeof v === "function") return v;
    }
    return null;
  }

  ensureGeneratedMarkerCss() {
    if (this.generatedCssText.includes(`.${this.settings.cssClassPrefix}li-marker`)) return;
    this.generatedCssText += [
      `.${this.settings.cssClassPrefix}li-marker { display: inline-block; min-width: 2.2em; }`,
    ].join("\n");
  }
}

function detectDocumentLanguage(root) {
  // Heuristic: first encountered w:lang/@w:val or w:lang/@val.
  for (const el of root.descendants()) {
    if (!(el instanceof XmlElement)) continue;
    if (!isW(el, "lang")) continue;
    const val = el.attributes.get("w:val") ?? el.attributes.get("val");
    if (val) return String(val);
  }
  return null;
}

async function readNotes(doc, partUri, settings) {
  try {
    let xml = parseXml(await doc.getPartText(partUri));
    xml = preprocessXmlForHtml(xml, settings);
    const byId = new Map();
    const orderedIds = [];

    for (const el of xml.root.descendants()) {
      if (!(el instanceof XmlElement)) continue;
      const local = el.nameParts().local;
      if (local !== "footnote" && local !== "endnote") continue;
      const id = el.attributes.get("w:id") ?? el.attributes.get("id");
      if (id == null) continue;
      const type = el.attributes.get("w:type") ?? el.attributes.get("type");
      if (type) continue; // separators/continuations/etc
      const idStr = String(id);
      byId.set(idStr, el);
      orderedIds.push(idStr);
    }

    // Sort numeric where possible for stable output.
    orderedIds.sort((a, b) => Number(a) - Number(b));
    return { byId, orderedIds };
  } catch {
    return null;
  }
}

async function readComments(doc, partUri, settings) {
  try {
    let xml = parseXml(await doc.getPartText(partUri));
    xml = preprocessXmlForHtml(xml, settings);
    const byId = new Map();
    const orderedIds = [];

    for (const el of xml.root.descendants()) {
      if (!(el instanceof XmlElement)) continue;
      if (!isW(el, "comment")) continue;
      const id = el.attributes.get("w:id") ?? el.attributes.get("id");
      if (id == null) continue;
      const idStr = String(id);
      byId.set(idStr, el);
      orderedIds.push(idStr);
    }

    orderedIds.sort((a, b) => Number(a) - Number(b));
    return { byId, orderedIds };
  } catch {
    return null;
  }
}

async function readRelationships(doc, partUri) {
  try {
    const rels = parseXml(await doc.getPartText(partUri));
    const byId = new Map();
    for (const el of rels.root.descendants()) {
      if (!(el instanceof XmlElement)) continue;
      if (el.nameParts().local !== "Relationship") continue;
      const id = el.attributes.get("Id");
      const type = el.attributes.get("Type");
      const target = el.attributes.get("Target");
      const targetMode = el.attributes.get("TargetMode") ?? null;
      if (id && target) byId.set(String(id), { id: String(id), type, target: String(target), targetMode });
    }
    return { byId };
  } catch {
    return { byId: new Map() };
  }
}

async function readContentTypes(doc) {
  try {
    const ct = parseXml(await doc.getPartText("/[Content_Types].xml"));
    const defaults = new Map();
    const overrides = new Map();
    for (const el of ct.root.descendants()) {
      if (!(el instanceof XmlElement)) continue;
      const local = el.nameParts().local;
      if (local === "Default") {
        const ext = el.attributes.get("Extension");
        const type = el.attributes.get("ContentType");
        if (ext && type) defaults.set(String(ext).toLowerCase(), String(type));
      } else if (local === "Override") {
        const partName = el.attributes.get("PartName");
        const type = el.attributes.get("ContentType");
        if (partName && type) overrides.set(String(partName), String(type));
      }
    }
    return {
      getContentType(partUri) {
        const pn = partUri.startsWith("/") ? partUri : `/${partUri}`;
        if (overrides.has(pn)) return overrides.get(pn);
        const idx = pn.lastIndexOf(".");
        if (idx !== -1) {
          const ext = pn.slice(idx + 1).toLowerCase();
          if (defaults.has(ext)) return defaults.get(ext);
        }
        return null;
      },
    };
  } catch {
    return { getContentType() { return null; } };
  }
}

async function readNumbering(doc) {
  try {
    const num = parseXml(await doc.getPartText("/word/numbering.xml"));
    const abstractById = new Map();
    const numById = new Map();

    for (const el of num.root.descendants()) {
      if (!(el instanceof XmlElement) || !isW(el)) continue;
      const local = el.nameParts().local;
      if (local === "abstractNum") {
        const id = el.attributes.get("w:abstractNumId") ?? el.attributes.get("abstractNumId");
        if (id != null) abstractById.set(String(id), el);
      } else if (local === "num") {
        const id = el.attributes.get("w:numId") ?? el.attributes.get("numId");
        if (id != null) numById.set(String(id), el);
      }
    }

    function getLevel(numId, ilvl) {
      const numEl = numById.get(String(numId));
      if (!numEl) return null;
      const absIdEl = numEl.children.find((c) => c instanceof XmlElement && isW(c, "abstractNumId"));
      const absId = absIdEl?.attributes.get("w:val") ?? absIdEl?.attributes.get("val");
      const absEl = absId != null ? abstractById.get(String(absId)) : null;
      if (!absEl) return null;
      const lvl = absEl.children.find(
        (c) => c instanceof XmlElement && isW(c, "lvl") && String(c.attributes.get("w:ilvl") ?? c.attributes.get("ilvl") ?? "") === String(ilvl),
      );
      if (!lvl) return null;
      const numFmtEl = lvl.children.find((c) => c instanceof XmlElement && isW(c, "numFmt"));
      const lvlTextEl = lvl.children.find((c) => c instanceof XmlElement && isW(c, "lvlText"));
      const startEl = lvl.children.find((c) => c instanceof XmlElement && isW(c, "start"));
      const numFmt = String(numFmtEl?.attributes.get("w:val") ?? numFmtEl?.attributes.get("val") ?? "decimal");
      const lvlText = String(lvlTextEl?.attributes.get("w:val") ?? lvlTextEl?.attributes.get("val") ?? "");
      const startVal = startEl?.attributes.get("w:val") ?? startEl?.attributes.get("val") ?? null;
      const start = startVal != null ? Number(startVal) : 1;
      return { numFmt, lvlText, start: Number.isFinite(start) && start > 0 ? start : 1 };
    }

    return { getLevel };
  } catch {
    return null;
  }
}

function toListTag(numFmt, restrict, warnings) {
  const fmt = String(numFmt);
  if (fmt === "bullet") return "ul";
  if (fmt === "decimal" || fmt === "lowerLetter" || fmt === "upperLetter" || fmt === "lowerRoman" || fmt === "upperRoman") {
    return "ol";
  }
  if (restrict) warnings.push({ code: "OXPT_LIST_UNSUPPORTED_NUMFMT", message: `Unsupported numbering format: ${fmt}`, part: "/word/numbering.xml" });
  return "ol";
}

function toListAttrs(numFmt) {
  const fmt = String(numFmt);
  if (fmt === "lowerLetter") return { type: "a" };
  if (fmt === "upperLetter") return { type: "A" };
  if (fmt === "lowerRoman") return { type: "i" };
  if (fmt === "upperRoman") return { type: "I" };
  return {};
}

function resolveWordTarget(target) {
  if (target.startsWith("/")) return target;
  if (target.startsWith("../")) return `/${target.replace(/^\.\.\//, "")}`;
  return `/word/${target}`;
}

function extractAltText(drawingElement) {
  const docPr = findFirstByLocal(drawingElement, "docPr");
  if (docPr instanceof XmlElement) {
    const descr = docPr.attributes.get("descr");
    if (descr) return String(descr);
    const name = docPr.attributes.get("name");
    if (name) return String(name);
  }
  return null;
}

function extractExtent(drawingElement) {
  const extent = findFirstByLocal(drawingElement, "extent");
  if (!(extent instanceof XmlElement)) return null;
  const cx = extent.attributes.get("cx");
  const cy = extent.attributes.get("cy");
  if (!cx || !cy) return null;
  return { cx: Number(cx), cy: Number(cy) };
}

function renderAttrs(attrs) {
  let out = "";
  for (const [k, v] of Object.entries(attrs)) out += ` ${escapeHtml(k)}="${escapeHtml(v)}"`;
  return out;
}

function renderAlt(altText) {
  return altText ? ` alt="${escapeHtml(altText)}"` : "";
}

function slug(s) {
  return s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}

function isW(el, localName) {
  const { prefix, local } = el.nameParts();
  if (localName && local !== localName) return false;
  if (prefix === "w") return true;
  return el.lookupNamespaceUri(prefix) === W_NS;
}

function twipsToPt(twips) {
  return Number(twips) / 20;
}

function halfPointsToPt(halfPoints) {
  return Number(halfPoints) / 2;
}

function lineHeightCss(line, lineRule) {
  if (line == null) return null;
  const n = Number(line);
  if (!Number.isFinite(n) || n <= 0) return null;
  const rule = lineRule == null ? "auto" : String(lineRule);

  // Word stores:
  // - w:lineRule="auto": w:line in 240ths of a line (240 == single)
  // - w:lineRule="exact"/"atLeast": w:line in twips
  if (rule === "auto") {
    const mult = n / 240;
    const v = String(Math.round(mult * 1000) / 1000);
    return `line-height:${v}`;
  }

  return `line-height:${twipsToPt(n)}pt`;
}

function mapJustificationToCssTextAlign(jc) {
  const v = String(jc);
  if (v === "left") return "left";
  if (v === "right") return "right";
  if (v === "center") return "center";
  if (v === "both" || v === "distribute") return "justify";
  return null;
}

function mapHighlightToCss(val) {
  const v = String(val);
  // Minimal mapping; Word highlight values include many named colors.
  if (v === "yellow") return "yellow";
  if (v === "green") return "lime";
  if (v === "cyan") return "cyan";
  if (v === "magenta") return "magenta";
  if (v === "blue") return "blue";
  if (v === "red") return "red";
  if (v === "black") return "black";
  if (v === "white") return "white";
  if (v === "none") return null;
  return null;
}

function cssString(s) {
  const v = String(s).replaceAll('"', '\\"');
  return `"${v}"`;
}

function suggestedImgStyle(info) {
  const w = emusToPx(info.widthEmus);
  const h = emusToPx(info.heightEmus);
  if (w == null && h == null) return null;
  const rules = [];
  if (w != null) rules.push(`width:${w}px`);
  if (h != null) rules.push(`height:${h}px`);
  rules.push("max-width:100%");
  if (w != null && h == null) rules.push("height:auto");
  return rules.join(";");
}

function emusToPx(emus) {
  if (emus == null) return null;
  const n = Number(emus);
  if (!Number.isFinite(n) || n <= 0) return null;
  // 1 inch == 914400 EMUs; assume 96 CSS px per inch.
  return Math.max(1, Math.round((n * 96) / 914400));
}
