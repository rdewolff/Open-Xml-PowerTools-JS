import { parseXml, XmlDocument, XmlElement, XmlText, serializeXml } from "./internal/xml.js";
import { bytesToBase64 } from "./util/base64.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const WmlToHtmlConverter = {
  async convertToHtml(doc, settings = {}) {
    const s = normalizeSettings(settings);
    // Preprocess like C# converter (minimal subset): accept revisions and simplify markup.
    // Do this in-memory for `/word/document.xml` so we don't rewrite the DOCX unless user requests.
    let xmlDoc = parseXml(await doc.getPartText("/word/document.xml"));
    xmlDoc = acceptRevisionsXml(xmlDoc);
    xmlDoc = simplifyForHtmlXml(xmlDoc);

    const warnings = [];

    const ctx = await WmlConversionContext.create(doc, xmlDoc, warnings, s);
    const body = ctx.findBody();
    const htmlParts = await renderBodyChildren(ctx, body);

    const cssText = [ctx.generatedCssText, s.generalCss, s.additionalCss].filter(Boolean).join("\n");
    const htmlElement = buildHtmlElement({
      title: s.pageTitle,
      cssText,
      bodyHtml: htmlParts.join("\n"),
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
    generalCss: settings.generalCss ?? "span { white-space: pre-wrap; }",
    additionalCss: settings.additionalCss ?? "",
    restrictToSupportedLanguages: settings.restrictToSupportedLanguages ?? false,
    restrictToSupportedNumberingFormats: settings.restrictToSupportedNumberingFormats ?? false,
    listItemImplementations: settings.listItemImplementations ?? null,
    imageHandler: settings.imageHandler ?? null,
    output: { format: settings.output?.format ?? "string" },
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
        out.push(`<li>${await renderParagraphContents(ctx, child)}</li>`);
        continue;
      }

      closeAllLists(out, listStack);
      const headingLevel = ctx.getHeadingLevel(child);
      const tag = headingLevel ? `h${headingLevel}` : "p";
      const cls = ctx.getParagraphClass(child);
      const classAttr = cls ? ` class="${escapeHtml(cls)}"` : "";
      out.push(`<${tag}${classAttr}>${await renderParagraphContents(ctx, child)}</${tag}>`);
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
      inner.push(await renderRun(ctx, child));
      continue;
    }
    if (isW(child, "hyperlink")) {
      inner.push(await renderHyperlink(ctx, child));
      continue;
    }
  }
  return inner.join("");
}

async function renderHyperlink(ctx, hyperlink) {
  const rid = hyperlink.attributes.get("r:id") ?? hyperlink.attributes.get("id");
  const href = rid ? ctx.getHyperlinkTarget(String(rid)) : null;
  const inner = [];
  for (const c of hyperlink.children ?? []) {
    if (!(c instanceof XmlElement)) continue;
    if (isW(c, "r")) inner.push(await renderRun(ctx, c));
  }
  const contents = inner.join("");
  if (!href) return contents;
  return `<a href="${escapeHtml(href)}">${contents}</a>`;
}

async function renderRun(ctx, r) {
  const rPr = (r.children ?? []).find((c) => c?.qname && isW(c, "rPr")) ?? null;
  const bold = !!(rPr && (rPr.children ?? []).some((c) => c?.qname && isW(c, "b")));
  const italic = !!(rPr && (rPr.children ?? []).some((c) => c?.qname && isW(c, "i")));
  const underline = !!(rPr && (rPr.children ?? []).some((c) => c?.qname && isW(c, "u")));
  const cls = ctx.getRunClass(r);

  const pieces = [];
  for (const child of r.children ?? []) {
    if (!child?.qname) continue;
    if (isW(child, "t")) pieces.push(escapeHtml(child.textContent()));
    else if (isW(child, "tab")) pieces.push("    ");
    else if (isW(child, "br")) pieces.push("<br/>");
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
  const classAttr = cls ? ` class="${escapeHtml(cls)}"` : "";
  return html ? `<span${classAttr}>${html}</span>` : "";
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
  const rows = [];
  for (const tr of tbl.children ?? []) {
    if (!(tr instanceof XmlElement) || !isW(tr, "tr")) continue;
    const cells = [];
    for (const tc of tr.children ?? []) {
      if (!(tc instanceof XmlElement) || !isW(tc, "tc")) continue;
      const paras = [];
      for (const p of tc.descendantsByNameNS(W_NS, "p")) {
        paras.push(`<p>${await renderParagraphContents(ctx, p)}</p>`);
      }
      cells.push(`<td>${paras.join("")}</td>`);
    }
    rows.push(`<tr>${cells.join("")}</tr>`);
  }
  return `<table><tbody>${rows.join("")}</tbody></table>`;
}

function ensureListStack(ctx, out, listStack, listInfo) {
  // listInfo.level is 0-based.
  while (listStack.length > listInfo.level + 1) {
    out.push(`</${listStack.pop().tag}>`);
  }
  while (listStack.length < listInfo.level + 1) {
    const { tag, attrs } = listInfo;
    out.push(`<${tag}${renderAttrs(attrs)}>`);
    listStack.push({ tag, attrs });
  }
  const current = listStack[listStack.length - 1];
  if (current.tag !== listInfo.tag) {
    out.push(`</${listStack.pop().tag}>`);
    out.push(`<${listInfo.tag}${renderAttrs(listInfo.attrs)}>`);
    listStack.push({ tag: listInfo.tag, attrs: listInfo.attrs });
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

function simplifyForHtmlXml(xmlDoc) {
  // Roughly matches the C# converter's SimplifyMarkupSettings used for HTML conversion.
  const settings = {
    removeComments: true,
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
  constructor({ doc, mainXml, warnings, settings, rels, numbering, contentTypes }) {
    this.doc = doc;
    this.mainXml = mainXml;
    this.warnings = warnings;
    this.settings = settings;
    this.rels = rels;
    this.numbering = numbering;
    this.contentTypes = contentTypes;
    this.generatedCssText = "";
  }

  static async create(doc, mainXml, warnings, settings) {
    const [rels, numbering, contentTypes] = await Promise.all([
      readRelationships(doc, "/word/_rels/document.xml.rels"),
      readNumbering(doc),
      readContentTypes(doc),
    ]);
    return new WmlConversionContext({ doc, mainXml, warnings, settings, rels, numbering, contentTypes });
  }

  findBody() {
    const it = this.mainXml.root.descendantsByNameNS(W_NS, "body");
    return it.next().value ?? findFirstByLocal(this.mainXml.root, "body");
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
    const pPr = (p.children ?? []).find((c) => c instanceof XmlElement && isW(c, "pPr"));
    const numPr = pPr?.children?.find((c) => c instanceof XmlElement && isW(c, "numPr"));
    if (!numPr) return null;
    const ilvlEl = numPr.children.find((c) => c instanceof XmlElement && isW(c, "ilvl"));
    const numIdEl = numPr.children.find((c) => c instanceof XmlElement && isW(c, "numId"));
    const level = Number(ilvlEl?.attributes.get("w:val") ?? ilvlEl?.attributes.get("val") ?? 0);
    const numId = String(numIdEl?.attributes.get("w:val") ?? numIdEl?.attributes.get("val") ?? "");
    if (!numId) return null;
    const lvl = this.numbering?.getLevel(numId, level);
    const numFmt = lvl?.numFmt ?? "decimal";
    const tag = toListTag(numFmt, this.settings.restrictToSupportedNumberingFormats, this.warnings);
    const attrs = toListAttrs(numFmt);
    return { numId, level, numFmt, tag, attrs };
  }

  getParagraphClass(p) {
    if (!this.settings.fabricateCssClasses) return null;
    const pPr = (p.children ?? []).find((c) => c instanceof XmlElement && isW(c, "pPr"));
    const pStyle = pPr?.children?.find((c) => c instanceof XmlElement && isW(c, "pStyle"));
    const val = pStyle?.attributes?.get("w:val") ?? pStyle?.attributes?.get("val");
    if (!val) return null;
    return `${this.settings.cssClassPrefix}p-${slug(String(val))}`;
  }

  getRunClass(r) {
    if (!this.settings.fabricateCssClasses) return null;
    const rPr = (r.children ?? []).find((c) => c instanceof XmlElement && isW(c, "rPr"));
    const rStyle = rPr?.children?.find((c) => c instanceof XmlElement && isW(c, "rStyle"));
    const val = rStyle?.attributes?.get("w:val") ?? rStyle?.attributes?.get("val");
    if (!val) return null;
    return `${this.settings.cssClassPrefix}r-${slug(String(val))}`;
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
    const bytes = await this.doc.getPartBytes(partUri);
    const contentType = this.contentTypes?.getContentType(partUri) ?? "application/octet-stream";
    const info = {
      contentType,
      bytes,
      altText: extractAltText(drawingElement),
      widthEmus: extractExtent(drawingElement)?.cx ?? undefined,
      heightEmus: extractExtent(drawingElement)?.cy ?? undefined,
      suggestedStyle: null,
    };

    if (this.settings.imageHandler) {
      const res = this.settings.imageHandler(info);
      if (!res) return "";
      if (res.element) return res.element;
      const attrs = res.attrs ?? {};
      return `<img src="${escapeHtml(res.src)}"${renderAttrs(attrs)}${renderAlt(info.altText)}/>`;
    }

    const b64 = bytesToBase64(bytes);
    const src = `data:${contentType};base64,${b64}`;
    return `<img src="${src}"${renderAlt(info.altText)}/>`;
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
      const numFmt = String(numFmtEl?.attributes.get("w:val") ?? numFmtEl?.attributes.get("val") ?? "decimal");
      const lvlText = String(lvlTextEl?.attributes.get("w:val") ?? lvlTextEl?.attributes.get("val") ?? "");
      return { numFmt, lvlText };
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
