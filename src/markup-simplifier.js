import { parseXml, XmlDocument, XmlElement, XmlText } from "./internal/xml.js";
import { OpenXmlPowerToolsError } from "./open-xml-powertools-error.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const MarkupSimplifier = {
  async simplifyMarkup(doc, settings) {
    if (!settings || typeof settings !== "object") {
      throw new OpenXmlPowerToolsError("OXPT_INVALID_ARGUMENT", "settings is required");
    }

    const mainXmlText = await doc.getPartText("/word/document.xml");
    const mainXml = parseXml(mainXmlText);
    let simplified = simplifyMainDocument(mainXml, settings);

    if (settings.removeGoBackBookmark) {
      simplified = removeGoBackBookmarks(simplified);
    }

    return doc.replacePartXml("/word/document.xml", simplified);
  },
};

function simplifyMainDocument(xmlDoc, settings) {
  const root = transformNode(xmlDoc.root, settings)[0];
  if (!(root instanceof XmlElement)) {
    throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Simplification removed the document root");
  }
  return new XmlDocument(root);
}

function transformNode(node, settings) {
  if (node instanceof XmlText) return [new XmlText(node.text)];
  if (!(node instanceof XmlElement)) return [];

  const local = node.nameParts().local;
  const prefix = node.nameParts().prefix;
  const isWord = prefix === "w" || node.lookupNamespaceUri(prefix) === W_NS;

  if (isWord) {
    if (settings.removeSmartTags && local === "smartTag") {
      return transformChildren(node, settings);
    }

    if (settings.removeContentControls && local === "sdt") {
      const sdtContent = node.children.find(
        (c) => c instanceof XmlElement && isW(c, "sdtContent"),
      );
      if (!sdtContent) return [];
      return transformChildren(sdtContent, settings);
    }

    if (settings.removeLastRenderedPageBreak && local === "lastRenderedPageBreak") return [];

    if (
      settings.removeComments &&
      (local === "commentRangeStart" || local === "commentRangeEnd" || local === "commentReference")
    ) {
      return [];
    }

    if (settings.removeBookmarks && (local === "bookmarkStart" || local === "bookmarkEnd")) return [];

    if (settings.replaceTabsWithSpaces && local === "tab") {
      const t = new XmlElement(inferQName(prefix, "t"), new Map([["xml:space", "preserve"]]), [new XmlText(" ")]);
      return [t];
    }

    if (settings.removeSoftHyphens && local === "t") {
      const newAttrs = filteredAttributes(node.attributes, settings);
      const text = node.textContent().replaceAll("\u00ad", "");
      return [new XmlElement(node.qname, newAttrs, [new XmlText(text)])];
    }
  }

  const newAttrs = filteredAttributes(node.attributes, settings);
  const newChildren = [];
  for (const c of node.children) newChildren.push(...transformNode(c, settings));
  return [new XmlElement(node.qname, newAttrs, newChildren)];
}

function transformChildren(el, settings) {
  const out = [];
  for (const c of el.children) out.push(...transformNode(c, settings));
  return out;
}

function filteredAttributes(attributes, settings) {
  if (!settings.removeRsidInfo) return new Map(attributes);
  const out = new Map();
  for (const [k, v] of attributes) {
    const local = splitQName(k).local;
    if (local.startsWith("rsid")) continue;
    out.set(k, v);
  }
  return out;
}

function removeGoBackBookmarks(xmlDoc) {
  const ids = new Set();
  collectGoBackBookmarkIds(xmlDoc.root, ids);
  if (!ids.size) return xmlDoc;
  const root = filterBookmarks(xmlDoc.root, ids);
  return new XmlDocument(root);
}

function collectGoBackBookmarkIds(node, ids) {
  if (node instanceof XmlElement) {
    if (isW(node, "bookmarkStart")) {
      const name = getAttrByLocalName(node, "name");
      if (name === "_GoBack") {
        const id = getAttrByLocalName(node, "id");
        if (id != null) ids.add(String(id));
      }
    }
    for (const c of node.children) collectGoBackBookmarkIds(c, ids);
  }
}

function filterBookmarks(node, ids) {
  if (node instanceof XmlText) return new XmlText(node.text);
  if (!(node instanceof XmlElement)) return null;

  if (isW(node, "bookmarkStart") || isW(node, "bookmarkEnd")) {
    const id = getAttrByLocalName(node, "id");
    if (id != null && ids.has(String(id))) return null;
  }

  const children = [];
  for (const c of node.children) {
    const transformed = filterBookmarks(c, ids);
    if (transformed) children.push(transformed);
  }
  return new XmlElement(node.qname, new Map(node.attributes), children);
}

function isW(el, localName) {
  const { prefix, local } = el.nameParts();
  if (local !== localName) return false;
  if (prefix === "w") return true;
  return el.lookupNamespaceUri(prefix) === W_NS;
}

function splitQName(qname) {
  const idx = qname.indexOf(":");
  if (idx === -1) return { prefix: "", local: qname };
  return { prefix: qname.slice(0, idx), local: qname.slice(idx + 1) };
}

function getAttrByLocalName(el, localName) {
  for (const [k, v] of el.attributes) {
    if (splitQName(k).local === localName) return v;
  }
  return null;
}

function inferQName(prefix, local) {
  return prefix ? `${prefix}:${local}` : local;
}

