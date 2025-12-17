import { parseXml, serializeXml, XmlDocument, XmlElement, XmlText } from "./internal/xml.js";
import { OpcPackage } from "./internal/opc.js";

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
    const pkg = await OpcPackage.fromBytes(doc.bytes, { adapter: doc.zipAdapter });
    const partUris = listWordprocessingRevisionParts(pkg);
    const replaceParts = {};

    for (const partUri of partUris) {
      let xmlText;
      try {
        xmlText = await doc.getPartText(partUri);
      } catch {
        continue;
      }
      let xml;
      try {
        xml = parseXml(xmlText);
      } catch {
        continue;
      }

      if (!isWordprocessingRoot(xml.root)) continue;
      if (!hasTrackedRevisionsInNode(xml.root)) continue;

      const accepted = acceptInMainDocument(xml);
      const outText = serializeXml(accepted, { xmlDeclaration: true });
      replaceParts[partUri] = new TextEncoder().encode(outText);
    }

    if (!Object.keys(replaceParts).length) return doc;
    return doc.replaceParts(replaceParts);
  },

  async hasTrackedRevisions(doc) {
    const pkg = await OpcPackage.fromBytes(doc.bytes, { adapter: doc.zipAdapter });
    const partUris = listWordprocessingRevisionParts(pkg);
    for (const partUri of partUris) {
      try {
        const xmlText = await doc.getPartText(partUri);
        const xml = parseXml(xmlText);
        if (isWordprocessingRoot(xml.root) && hasTrackedRevisionsInNode(xml.root)) return true;
      } catch {
        // ignore
      }
    }
    return false;
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

function isWordprocessingRoot(root) {
  if (!(root instanceof XmlElement)) return false;
  if (!isW(root)) return false;
  const local = root.nameParts().local;
  return (
    local === "document" ||
    local === "hdr" ||
    local === "ftr" ||
    local === "footnotes" ||
    local === "endnotes" ||
    local === "comments"
  );
}

function listWordprocessingRevisionParts(pkg) {
  const out = [];
  for (const name of pkg.listPartNames()) {
    const n = String(name);
    if (!n.startsWith("word/")) continue;
    if (!n.endsWith(".xml")) continue;
    if (n === "word/document.xml") out.push("/word/document.xml");
    else if (n === "word/footnotes.xml") out.push("/word/footnotes.xml");
    else if (n === "word/endnotes.xml") out.push("/word/endnotes.xml");
    else if (n === "word/comments.xml") out.push("/word/comments.xml");
    else if (/^word\/header\d+\.xml$/.test(n)) out.push(`/${n}`);
    else if (/^word\/footer\d+\.xml$/.test(n)) out.push(`/${n}`);
  }
  // Ensure deterministic order for stable outputs/tests.
  out.sort();
  return out;
}
