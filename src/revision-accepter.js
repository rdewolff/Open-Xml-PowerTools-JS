import { parseXml, XmlDocument, XmlElement, XmlText } from "./internal/xml.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const REVISION_WRAPPERS_ACCEPT = new Set(["ins", "moveTo"]);
const REVISION_WRAPPERS_DROP = new Set(["del", "moveFrom"]);
const REVISION_MARKERS_DROP = new Set([
  "delText",
  "delInstrText",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

export const RevisionAccepter = {
  async acceptRevisions(doc) {
    const mainXmlText = await doc.getPartText("/word/document.xml");
    const mainXml = parseXml(mainXmlText);
    const accepted = acceptInMainDocument(mainXml);
    return doc.replacePartXml("/word/document.xml", accepted);
  },

  async hasTrackedRevisions(doc) {
    const mainXmlText = await doc.getPartText("/word/document.xml");
    const mainXml = parseXml(mainXmlText);
    return hasTrackedRevisionsInNode(mainXml.root);
  },
};

function acceptInMainDocument(xmlDoc) {
  const rootNodes = acceptTransform(xmlDoc.root);
  const root = rootNodes[0];
  return new XmlDocument(root);
}

function acceptTransform(node) {
  if (node instanceof XmlText) return [new XmlText(node.text)];
  if (!(node instanceof XmlElement)) return [];

  if (isW(node)) {
    const local = node.nameParts().local;
    if (REVISION_WRAPPERS_ACCEPT.has(local)) {
      return acceptChildren(node);
    }
    if (REVISION_WRAPPERS_DROP.has(local) || REVISION_MARKERS_DROP.has(local)) {
      return [];
    }
  }

  const children = [];
  for (const c of node.children) children.push(...acceptTransform(c));
  return [new XmlElement(node.qname, new Map(node.attributes), children)];
}

function acceptChildren(el) {
  const out = [];
  for (const c of el.children) out.push(...acceptTransform(c));
  return out;
}

function hasTrackedRevisionsInNode(node) {
  if (!(node instanceof XmlElement)) return false;
  if (isW(node)) {
    const local = node.nameParts().local;
    if (REVISION_WRAPPERS_ACCEPT.has(local) || REVISION_WRAPPERS_DROP.has(local) || REVISION_MARKERS_DROP.has(local)) {
      return true;
    }
  }
  for (const c of node.children) {
    if (hasTrackedRevisionsInNode(c)) return true;
  }
  return false;
}

function isW(el) {
  const { prefix } = el.nameParts();
  if (prefix === "w") return true;
  return el.lookupNamespaceUri(prefix) === W_NS;
}

