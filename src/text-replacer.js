import { parseXml, XmlDocument, XmlElement, XmlText } from "./internal/xml.js";
import { OpenXmlPowerToolsError } from "./open-xml-powertools-error.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const TextReplacer = {
  async searchAndReplace(doc, search, replace, options = {}) {
    const matchCase = options.matchCase ?? false;
    if (typeof search !== "string" || !search.length) {
      throw new OpenXmlPowerToolsError("OXPT_INVALID_ARGUMENT", "search must be a non-empty string");
    }
    if (typeof replace !== "string") {
      throw new OpenXmlPowerToolsError("OXPT_INVALID_ARGUMENT", "replace must be a string");
    }

    const mainXmlText = await doc.getPartText("/word/document.xml");
    const mainXml = parseXml(mainXmlText);
    const newMainXml = transformDocument(mainXml, search, replace, matchCase);
    return doc.replacePartXml("/word/document.xml", newMainXml);
  },
};

function transformDocument(xmlDoc, search, replace, matchCase) {
  const root = cloneElement(xmlDoc.root);

  for (const p of root.descendantsByNameNS(W_NS, "p")) {
    const replaced = replaceInParagraph(p, search, replace, matchCase);
    if (!replaced) continue;
    p.attributes = replaced.attributes;
    p.children = replaced.children;
    for (const child of p.children) if (child instanceof XmlElement) child.parent = p;
  }

  return new XmlDocument(root);
}

function replaceInParagraph(p, search, replace, matchCase) {
  const paragraphChildren = p.children.filter((n) => !isIgnorableWhitespaceText(n));
  const expanded = [];
  for (const child of paragraphChildren) {
    if (child instanceof XmlElement && isW(child, "r")) {
      expanded.push(...splitRun(child));
    } else {
      expanded.push(cloneNode(child));
    }
  }

  const chars = expanded.map((n) => (isCharRun(n) ? getRunChar(n) : null));
  const used = new Array(expanded.length).fill(false);
  const searchChars = [...search];
  const want = matchCase ? searchChars : searchChars.map((c) => c.toUpperCase());

  const out = [];
  let i = 0;
  while (i < expanded.length) {
    if (!isMatchAt(i)) {
      out.push(expanded[i]);
      i++;
      continue;
    }

    const firstRun = expanded[i];
    const rPr = getRunProperties(firstRun);
    out.push(makeTextRun(firstRun.qname, rPr, replace));
    for (let j = 0; j < searchChars.length; j++) used[i + j] = true;
    i += searchChars.length;
  }

  const consolidated = consolidateAdjacentRuns(out);
  return new XmlElement(p.qname, new Map(p.attributes), consolidated);

  function isMatchAt(start) {
    if (start + searchChars.length > expanded.length) return false;
    for (let j = 0; j < searchChars.length; j++) {
      if (used[start + j]) return false;
      const ch = chars[start + j];
      if (ch == null) return false;
      const got = matchCase ? ch : ch.toUpperCase();
      if (got !== want[j]) return false;
    }
    return true;
  }
}

function splitRun(run) {
  const children = run.children.filter((n) => !isIgnorableWhitespaceText(n));
  const rPr = children.find((c) => c instanceof XmlElement && isW(c, "rPr")) ?? null;
  const rPrClone = rPr ? cloneElement(rPr) : null;

  const out = [];
  for (const child of children) {
    if (!(child instanceof XmlElement)) continue;
    if (isW(child, "rPr")) continue;
    if (isW(child, "t")) {
      const text = child.textContent();
      if (!text.length) continue;
      for (const ch of [...text]) out.push(makeTextRun(run.qname, rPrClone, ch));
      continue;
    }
    out.push(makeRunWithChild(run.qname, rPrClone, cloneElement(child)));
  }
  return out;
}

function isCharRun(run) {
  if (!(run instanceof XmlElement) || !isW(run, "r")) return false;
  const children = run.children.filter((n) => n instanceof XmlElement);
  const nonRpr = children.filter((c) => !isW(c, "rPr"));
  if (nonRpr.length !== 1) return false;
  const only = nonRpr[0];
  if (!isW(only, "t")) return false;
  const text = only.textContent();
  return text.length === 1;
}

function getRunChar(run) {
  const t = run.children.find((c) => c && c.qname && isW(c, "t"));
  if (!t) return null;
  const text = t.textContent();
  return text.length === 1 ? text : null;
}

function getRunProperties(run) {
  const rPr = run.children.find((c) => c instanceof XmlElement && isW(c, "rPr"));
  return rPr ? cloneElement(rPr) : null;
}

function makeTextRun(runQName, rPr, text) {
  const tAttrs = new Map();
  if (text.startsWith(" ") || text.endsWith(" ")) tAttrs.set("xml:space", "preserve");
  const tEl = new XmlElement(inferQNamePrefix(runQName, "t"), tAttrs, [new XmlText(text)]);
  return makeRunWithChild(runQName, rPr, tEl);
}

function makeRunWithChild(runQName, rPr, childEl) {
  const children = [];
  if (rPr) children.push(cloneElement(rPr));
  children.push(childEl);
  return new XmlElement(runQName, new Map(), children);
}

function consolidateAdjacentRuns(nodes) {
  const out = [];
  let i = 0;
  while (i < nodes.length) {
    const node = nodes[i];
    const key = getConsolidationKey(node);
    if (!key) {
      out.push(node);
      i++;
      continue;
    }

    let j = i + 1;
    let text = getRunText(node);
    while (j < nodes.length && getConsolidationKey(nodes[j]) === key) {
      text += getRunText(nodes[j]);
      j++;
    }

    const merged = cloneElement(node);
    const t = merged.children.find((c) => c instanceof XmlElement && isW(c, "t"));
    if (t) {
      t.children = [new XmlText(text)];
      const needsPreserve = text.startsWith(" ") || text.endsWith(" ");
      if (needsPreserve) t.attributes.set("xml:space", "preserve");
      else t.attributes.delete("xml:space");
    }
    out.push(merged);
    i = j;
  }
  return out;
}

function getConsolidationKey(run) {
  if (!(run instanceof XmlElement) || !isW(run, "r")) return null;
  const childEls = run.children.filter((c) => c instanceof XmlElement);
  const nonRpr = childEls.filter((c) => !isW(c, "rPr"));
  if (nonRpr.length !== 1 || !isW(nonRpr[0], "t")) return null;
  const rPr = childEls.find((c) => isW(c, "rPr"));
  return rPr ? stableKey(rPr) : "";
}

function getRunText(run) {
  const t = run.children.find((c) => c && c.qname && isW(c, "t"));
  return t ? t.textContent() : "";
}

function stableKey(el) {
  // Minimal stable key: qname + attributes + textContent for the subtree.
  // This is good enough for grouping runs in our initial implementation.
  let out = el.qname;
  for (const [k, v] of el.attributes) out += `|${k}=${v}`;
  for (const child of el.children) {
    if (child instanceof XmlElement) out += `{${stableKey(child)}}`;
    else if (child instanceof XmlText) out += `#${child.text}`;
  }
  return out;
}

function isW(el, local) {
  const { prefix, local: l } = el.nameParts();
  if (l !== local) return false;
  // Fast path for the common DOCX prefix.
  if (prefix === "w") return true;
  const ns = el.lookupNamespaceUri(prefix);
  return ns === W_NS;
}

function isIgnorableWhitespaceText(n) {
  return n instanceof XmlText && n.text.trim() === "";
}

function cloneNode(node) {
  if (node instanceof XmlText) return new XmlText(node.text);
  if (node instanceof XmlElement) return cloneElement(node);
  return node;
}

function cloneElement(el) {
  return new XmlElement(el.qname, new Map(el.attributes), el.children.map(cloneNode));
}

function inferQNamePrefix(runQName, local) {
  const idx = String(runQName).indexOf(":");
  if (idx === -1) return local;
  const prefix = String(runQName).slice(0, idx);
  return `${prefix}:${local}`;
}
