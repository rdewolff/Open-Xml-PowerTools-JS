import { OpenXmlPowerToolsError } from "../open-xml-powertools-error.js";

export function parseXml(xmlText) {
  const parser = new XmlParser(xmlText);
  return parser.parse();
}

class XmlParser {
  constructor(text) {
    this.text = text;
    this.i = 0;
    this.len = text.length;
  }

  parse() {
    this.skipBomAndWhitespace();
    this.consumeProlog();

    const root = this.parseElement();
    this.skipWhitespace();
    return new XmlDocument(root);
  }

  skipBomAndWhitespace() {
    if (this.text.charCodeAt(0) === 0xfeff) this.i++;
    this.skipWhitespace();
  }

  consumeProlog() {
    if (this.peek("<?xml")) {
      const end = this.text.indexOf("?>", this.i);
      if (end === -1) throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Unterminated XML declaration");
      this.i = end + 2;
      this.skipWhitespace();
    }
    while (this.peek("<!--")) this.consumeComment();
    if (this.peek("<!DOCTYPE")) this.consumeDoctype();
    while (this.peek("<!--")) this.consumeComment();
  }

  consumeComment() {
    const end = this.text.indexOf("-->", this.i);
    if (end === -1) throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Unterminated comment");
    this.i = end + 3;
    this.skipWhitespace();
  }

  consumeDoctype() {
    // Minimal doctype skipper, handles internal subset brackets by balancing.
    if (!this.peek("<!DOCTYPE")) return;
    let depth = 0;
    while (this.i < this.len) {
      const ch = this.text[this.i++];
      if (ch === "[") depth++;
      else if (ch === "]") depth = Math.max(0, depth - 1);
      else if (ch === ">" && depth === 0) break;
    }
    this.skipWhitespace();
  }

  parseElement() {
    this.expectChar("<");
    if (this.peek("!--")) throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Unexpected comment");
    if (this.peek("/")) throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Unexpected closing tag");

    const { qname } = this.parseQName();
    const attrs = new Map();

    while (true) {
      this.skipWhitespace();
      if (this.peek("/>")) {
        this.i += 2;
        return new XmlElement(qname, attrs, []);
      }
      if (this.peek(">")) {
        this.i += 1;
        break;
      }
      const { qname: attrName } = this.parseQName();
      this.skipWhitespace();
      this.expectChar("=");
      this.skipWhitespace();
      const value = this.parseAttributeValue();
      attrs.set(attrName, value);
    }

    const children = [];
    while (true) {
      if (this.peek("</")) break;
      if (this.peek("<!--")) {
        this.consumeComment();
        continue;
      }
      if (this.peek("<")) {
        if (this.peek("<![CDATA[")) {
          children.push(new XmlText(this.parseCdata()));
          continue;
        }
        children.push(this.parseElement());
        continue;
      }
      const text = this.parseText();
      if (text.length) children.push(new XmlText(text));
    }

    this.expectString("</");
    const { qname: closeName } = this.parseQName();
    if (closeName !== qname) {
      throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", `Mismatched closing tag: expected </${qname}> got </${closeName}>`);
    }
    this.skipWhitespace();
    this.expectChar(">");
    return new XmlElement(qname, attrs, children);
  }

  parseQName() {
    this.skipWhitespace();
    const start = this.i;
    while (this.i < this.len) {
      const c = this.text[this.i];
      if (
        (c >= "a" && c <= "z") ||
        (c >= "A" && c <= "Z") ||
        (c >= "0" && c <= "9") ||
        c === ":" ||
        c === "_" ||
        c === "-" ||
        c === "."
      ) {
        this.i++;
        continue;
      }
      break;
    }
    if (this.i === start) throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Expected name");
    return { qname: this.text.slice(start, this.i) };
  }

  parseAttributeValue() {
    const quote = this.text[this.i];
    if (quote !== `"` && quote !== `'`) {
      throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Expected quoted attribute value");
    }
    this.i++;
    const start = this.i;
    while (this.i < this.len && this.text[this.i] !== quote) this.i++;
    if (this.i >= this.len) throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Unterminated attribute value");
    const raw = this.text.slice(start, this.i);
    this.i++;
    return decodeEntities(raw);
  }

  parseText() {
    const start = this.i;
    while (this.i < this.len && this.text[this.i] !== "<") this.i++;
    return decodeEntities(this.text.slice(start, this.i));
  }

  parseCdata() {
    this.expectString("<![CDATA[");
    const end = this.text.indexOf("]]>", this.i);
    if (end === -1) throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", "Unterminated CDATA");
    const data = this.text.slice(this.i, end);
    this.i = end + 3;
    return data;
  }

  skipWhitespace() {
    while (this.i < this.len) {
      const c = this.text.charCodeAt(this.i);
      if (c === 0x20 || c === 0x0a || c === 0x0d || c === 0x09) {
        this.i++;
        continue;
      }
      break;
    }
  }

  peek(s) {
    return this.text.startsWith(s, this.i);
  }

  expectString(s) {
    if (!this.text.startsWith(s, this.i)) {
      throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", `Expected '${s}'`);
    }
    this.i += s.length;
  }

  expectChar(ch) {
    if (this.text[this.i] !== ch) {
      throw new OpenXmlPowerToolsError("OXPT_XML_INVALID", `Expected '${ch}'`);
    }
    this.i++;
  }
}

function decodeEntities(text) {
  if (!text.includes("&")) return text;
  return text
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'");
}

export class XmlDocument {
  constructor(root) {
    this.root = root;
  }
}

export class XmlElement {
  constructor(qname, attributes, children) {
    this.qname = qname;
    this.attributes = attributes;
    this.children = children;
    this.parent = null;
    for (const child of children) {
      if (child instanceof XmlElement) child.parent = this;
    }
  }

  nameParts() {
    const idx = this.qname.indexOf(":");
    if (idx === -1) return { prefix: "", local: this.qname };
    return { prefix: this.qname.slice(0, idx), local: this.qname.slice(idx + 1) };
  }

  getAttribute(qname) {
    return this.attributes.get(qname) ?? null;
  }

  textContent() {
    let out = "";
    for (const child of this.children) {
      if (child instanceof XmlText) out += child.text;
      else out += child.textContent();
    }
    return out;
  }

  *descendants() {
    for (const child of this.children) {
      if (!(child instanceof XmlElement)) continue;
      yield child;
      yield* child.descendants();
    }
  }

  *descendantsByQName(qname) {
    for (const d of this.descendants()) if (d.qname === qname) yield d;
  }

  *descendantsByNameNS(nsUri, localName) {
    for (const d of this.descendants()) {
      const { prefix, local } = d.nameParts();
      if (local !== localName) continue;
      const xmlns = prefix ? d.lookupNamespaceUri(prefix) : d.lookupNamespaceUri("");
      if (xmlns === nsUri) yield d;
    }
  }

  lookupNamespaceUri(prefix) {
    const key = prefix ? `xmlns:${prefix}` : "xmlns";
    for (let el = this; el; el = el.parent) {
      const direct = el.attributes.get(key);
      if (direct != null) return direct;
    }
    return null;
  }
}

export class XmlText {
  constructor(text) {
    this.text = text;
  }
}
