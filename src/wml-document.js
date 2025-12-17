import { OpenXmlPowerToolsDocument, coerceToBytes } from "./open-xml-powertools-document.js";
import { OpcPackage } from "./internal/opc.js";
import { parseXml, serializeXml } from "./internal/xml.js";
import { getDefaultZipAdapter } from "./internal/zip-adapter-auto.js";

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

  async getPartText(uri) {
    const bytes = await this.getPartBytes(uri);
    return new TextDecoder("utf-8").decode(bytes);
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

  async searchAndReplace(search, replace, matchCase = false) {
    const { TextReplacer } = await import("./text-replacer.js");
    return TextReplacer.searchAndReplace(this, search, replace, { matchCase });
  }

  async simplifyMarkup(settings) {
    const { MarkupSimplifier } = await import("./markup-simplifier.js");
    return MarkupSimplifier.simplifyMarkup(this, settings);
  }

  async acceptRevisions() {
    const { RevisionAccepter } = await import("./revision-accepter.js");
    return RevisionAccepter.acceptRevisions(this);
  }

  async replacePartXml(partUri, xmlDocumentOrElement) {
    const xmlText = serializeXml(xmlDocumentOrElement, { xmlDeclaration: true });
    const bytes = new TextEncoder().encode(xmlText);
    return this.replaceParts({ [partUri]: bytes });
  }

  async replaceParts(replaceParts, options = {}) {
    const adapter = options.adapter ?? (await getDefaultZipAdapter());
    const pkg = await OpcPackage.fromBytes(this.bytes, { adapter });
    const newBytes = await pkg.toBytes({ replaceParts, adapter, deflateLevel: options.deflateLevel });
    return new WmlDocument(newBytes, { fileName: this.fileName });
  }
}
