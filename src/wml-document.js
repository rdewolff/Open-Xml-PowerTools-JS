import { OpenXmlPowerToolsDocument, coerceToBytes } from "./open-xml-powertools-document.js";
import { OpcPackage } from "./internal/opc.js";
import { parseXml } from "./internal/xml.js";

const WORD_MAIN_DOCUMENT_URI = "/word/document.xml";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export class WmlDocument extends OpenXmlPowerToolsDocument {
  static fromBytes(bytes, options) {
    return new WmlDocument(bytes, options);
  }

  constructor(bytes, options = {}) {
    super(bytes, options);
    this.bytes = coerceToBytes(bytes);
  }

  async getPartBytes(uri) {
    const pkg = await OpcPackage.fromBytes(this.bytes);
    return pkg.getPartBytes(uri);
  }

  async getMainDocumentXml() {
    const xmlBytes = await this.getPartBytes(WORD_MAIN_DOCUMENT_URI);
    const xmlText = new TextDecoder("utf-8").decode(xmlBytes);
    return parseXml(xmlText);
  }

  async getMainDocumentText() {
    const doc = await this.getMainDocumentXml();
    const paragraphs = [];
    for (const p of doc.root.descendantsByNameNS(W_NS, "p")) {
      const runs = [];
      for (const t of p.descendantsByNameNS(W_NS, "t")) runs.push(t.textContent());
      paragraphs.push(runs.join(""));
    }
    const text = paragraphs.join("\n");
    return { paragraphs, text };
  }
}

