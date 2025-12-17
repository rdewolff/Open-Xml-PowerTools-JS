import { parseXml, XmlElement } from "./xml.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export async function readWmlStyles(doc) {
  try {
    const xmlText = await doc.getPartText("/word/styles.xml");
    const xml = parseXml(xmlText);
    return WmlStyles.fromXml(xml.root);
  } catch {
    return WmlStyles.empty();
  }
}

export class WmlStyles {
  constructor({ paragraphById, characterById }) {
    this.paragraphById = paragraphById;
    this.characterById = characterById;
    this._cache = new Map();
  }

  static empty() {
    return new WmlStyles({ paragraphById: new Map(), characterById: new Map() });
  }

  static fromXml(root) {
    const paragraphById = new Map();
    const characterById = new Map();

    for (const el of root.descendants()) {
      if (!(el instanceof XmlElement)) continue;
      if (!isW(el, "style")) continue;
      const type = el.attributes.get("w:type") ?? el.attributes.get("type");
      const styleId = el.attributes.get("w:styleId") ?? el.attributes.get("styleId");
      if (!type || !styleId) continue;
      const basedOnEl = el.children.find((c) => c instanceof XmlElement && isW(c, "basedOn"));
      const basedOn = basedOnEl?.attributes.get("w:val") ?? basedOnEl?.attributes.get("val") ?? null;
      const pPr = el.children.find((c) => c instanceof XmlElement && isW(c, "pPr")) ?? null;
      const rPr = el.children.find((c) => c instanceof XmlElement && isW(c, "rPr")) ?? null;
      const rec = { styleId: String(styleId), basedOn: basedOn ? String(basedOn) : null, pPr, rPr };

      if (String(type) === "paragraph") paragraphById.set(String(styleId), rec);
      else if (String(type) === "character") characterById.set(String(styleId), rec);
    }

    return new WmlStyles({ paragraphById, characterById });
  }

  resolveParagraphStyle(styleId) {
    return this._resolve("p", styleId, this.paragraphById);
  }

  resolveCharacterStyle(styleId) {
    return this._resolve("c", styleId, this.characterById);
  }

  _resolve(kind, styleId, map) {
    if (!styleId) return null;
    const key = `${kind}:${styleId}`;
    if (this._cache.has(key)) return this._cache.get(key);

    const style = map.get(styleId);
    if (!style) {
      this._cache.set(key, null);
      return null;
    }

    const chain = [];
    const seen = new Set();
    let cur = style;
    while (cur && !seen.has(cur.styleId)) {
      seen.add(cur.styleId);
      chain.push(cur);
      cur = cur.basedOn ? map.get(cur.basedOn) : null;
    }

    // base-first merge
    chain.reverse();
    const merged = {
      styleId,
      pPr: null,
      rPr: null,
      para: {},
      run: {},
    };
    for (const s of chain) {
      if (s.pPr) {
        merged.pPr = s.pPr;
        Object.assign(merged.para, parseParagraphProperties(s.pPr));
      }
      if (s.rPr) {
        merged.rPr = s.rPr;
        Object.assign(merged.run, parseRunProperties(s.rPr));
      }
    }

    this._cache.set(key, merged);
    return merged;
  }
}

export function parseParagraphProperties(pPr) {
  const out = {};
  const jc = firstChildW(pPr, "jc");
  const jcVal = jc?.attributes.get("w:val") ?? jc?.attributes.get("val");
  if (jcVal) out.jc = String(jcVal);

  const spacing = firstChildW(pPr, "spacing");
  if (spacing) {
    out.before = toInt(spacing.attributes.get("w:before") ?? spacing.attributes.get("before"));
    out.after = toInt(spacing.attributes.get("w:after") ?? spacing.attributes.get("after"));
    out.line = toInt(spacing.attributes.get("w:line") ?? spacing.attributes.get("line"));
    out.lineRule = spacing.attributes.get("w:lineRule") ?? spacing.attributes.get("lineRule") ?? null;
  }

  const ind = firstChildW(pPr, "ind");
  if (ind) {
    out.left = toInt(ind.attributes.get("w:left") ?? ind.attributes.get("left"));
    out.right = toInt(ind.attributes.get("w:right") ?? ind.attributes.get("right"));
    out.firstLine = toInt(ind.attributes.get("w:firstLine") ?? ind.attributes.get("firstLine"));
    out.hanging = toInt(ind.attributes.get("w:hanging") ?? ind.attributes.get("hanging"));
  }

  const numPr = firstChildW(pPr, "numPr");
  if (numPr) {
    const ilvlEl = firstChildW(numPr, "ilvl");
    const numIdEl = firstChildW(numPr, "numId");
    const ilvl = toInt(ilvlEl?.attributes.get("w:val") ?? ilvlEl?.attributes.get("val"));
    const numId = numIdEl?.attributes.get("w:val") ?? numIdEl?.attributes.get("val");
    if (numId != null) out.numPr = { numId: String(numId), ilvl: ilvl ?? 0 };
  }

  const pStyle = firstChildW(pPr, "pStyle");
  const pStyleId = pStyle?.attributes.get("w:val") ?? pStyle?.attributes.get("val");
  if (pStyleId) out.pStyle = String(pStyleId);

  return out;
}

export function parseRunProperties(rPr) {
  const out = {};
  if (firstChildW(rPr, "b")) out.bold = true;
  if (firstChildW(rPr, "i")) out.italic = true;

  const u = firstChildW(rPr, "u");
  if (u) {
    const val = u.attributes.get("w:val") ?? u.attributes.get("val") ?? "single";
    out.underline = String(val) !== "none";
  }

  const color = firstChildW(rPr, "color");
  if (color) {
    const val = color.attributes.get("w:val") ?? color.attributes.get("val");
    if (val && String(val) !== "auto") out.color = `#${String(val)}`;
  }

  const sz = firstChildW(rPr, "sz");
  const szVal = sz?.attributes.get("w:val") ?? sz?.attributes.get("val");
  if (szVal) out.sz = toInt(szVal);

  const rFonts = firstChildW(rPr, "rFonts");
  if (rFonts) {
    const ascii = rFonts.attributes.get("w:ascii") ?? rFonts.attributes.get("ascii");
    const hAnsi = rFonts.attributes.get("w:hAnsi") ?? rFonts.attributes.get("hAnsi");
    out.fontFamily = String(ascii ?? hAnsi ?? "");
  }

  const highlight = firstChildW(rPr, "highlight");
  const hlVal = highlight?.attributes.get("w:val") ?? highlight?.attributes.get("val");
  if (hlVal) out.highlight = String(hlVal);

  const rStyle = firstChildW(rPr, "rStyle");
  const rStyleId = rStyle?.attributes.get("w:val") ?? rStyle?.attributes.get("val");
  if (rStyleId) out.rStyle = String(rStyleId);

  return out;
}

function firstChildW(el, local) {
  return (el.children ?? []).find((c) => c instanceof XmlElement && isW(c, local)) ?? null;
}

function isW(el, localName) {
  const { prefix, local } = el.nameParts();
  if (local !== localName) return false;
  if (prefix === "w") return true;
  return el.lookupNamespaceUri(prefix) === W_NS;
}

function toInt(v) {
  if (v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

