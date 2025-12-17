import { parseXml, XmlElement } from "./xml.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export async function readWmlPartXml(doc, partUri) {
  try {
    const xmlText = await doc.getPartText(partUri);
    return parseXml(xmlText);
  } catch {
    return null;
  }
}

export function findBody(root) {
  const it = root.descendantsByNameNS(W_NS, "body");
  return it.next().value ?? findFirstByLocal(root, "body");
}

export function getSectPrs(body) {
  const out = [];
  if (!body) return out;
  // sectPr can occur as direct child of w:body or within last paragraph's pPr.
  for (const el of body.descendantsByNameNS(W_NS, "sectPr")) out.push(el);
  return out;
}

export function splitBodyIntoSections(body) {
  // In WordprocessingML, section properties (w:sectPr) terminate a section and apply to the
  // content that precedes them. They can appear as:
  // - direct child of w:body (typically as the last element)
  // - within a paragraph's properties: w:p/w:pPr/w:sectPr (section break)
  if (!body) return [{ nodes: [], sectPr: null }];

  const out = [];
  let cur = [];

  for (const child of body.children ?? []) {
    if (!(child instanceof XmlElement)) continue;

    if (isW(child, "sectPr")) {
      out.push({ nodes: cur, sectPr: child });
      cur = [];
      continue;
    }

    cur.push(child);

    if (isW(child, "p")) {
      const sectPr = findParagraphSectPr(child);
      if (sectPr) {
        out.push({ nodes: cur, sectPr });
        cur = [];
      }
    }
  }

  if (cur.length || !out.length) out.push({ nodes: cur, sectPr: null });
  return out;
}

export function selectHeaderFooterRefs(sectPr) {
  // Prefer default, else first/even.
  const refs = { header: null, footer: null };
  for (const child of sectPr.children ?? []) {
    if (!(child instanceof XmlElement)) continue;
    const local = child.nameParts().local;
    if (local !== "headerReference" && local !== "footerReference") continue;
    const type = child.attributes.get("w:type") ?? child.attributes.get("type") ?? "default";
    const rid = child.attributes.get("r:id") ?? child.attributes.get("id") ?? null;
    if (!rid) continue;
    const kind = local === "headerReference" ? "header" : "footer";
    if (!refs[kind]) refs[kind] = {};
    refs[kind][String(type)] = String(rid);
  }
  return refs;
}

export function pickRef(refMap) {
  if (!refMap) return null;
  return refMap.default ?? refMap.first ?? refMap.even ?? null;
}

function findFirstByLocal(root, localName) {
  for (const d of root.descendants()) {
    if (d.nameParts().local === localName) return d;
  }
  return null;
}

function findParagraphSectPr(p) {
  const pPr = (p.children ?? []).find((c) => c instanceof XmlElement && isW(c, "pPr")) ?? null;
  if (!pPr) return null;
  return (pPr.children ?? []).find((c) => c instanceof XmlElement && isW(c, "sectPr")) ?? null;
}

function isW(el, localName) {
  const { prefix, local } = el.nameParts();
  if (local !== localName) return false;
  if (prefix === "w") return true;
  return el.lookupNamespaceUri(prefix) === W_NS;
}
