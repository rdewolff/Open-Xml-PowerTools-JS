import { OpenXmlPowerToolsError } from "../open-xml-powertools-error.js";
import { ZipArchive } from "./zip.js";
import { getDefaultZipAdapter } from "./zip-adapter-auto.js";
import { parseXml } from "./xml.js";

const CONTENT_TYPES = "[Content_Types].xml";
const ROOT_RELS = "_rels/.rels";
const OFFICE_DOCUMENT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

export class OpcPackage {
  constructor(zip) {
    this.zip = zip;
  }

  static async fromBytes(bytes, { adapter } = {}) {
    const resolvedAdapter = adapter ?? (await getDefaultZipAdapter());
    const zip = await ZipArchive.fromBytes(bytes, { adapter: resolvedAdapter });
    return new OpcPackage(zip);
  }

  listPartNames() {
    return this.zip.entries.map((e) => e.name);
  }

  async getPartBytes(partUri) {
    const name = partUri.startsWith("/") ? partUri.slice(1) : partUri;
    return this.zip.getEntryBytes(name);
  }

  async getContentTypesXml() {
    const bytes = await this.zip.getEntryBytes(CONTENT_TYPES);
    const text = new TextDecoder("utf-8").decode(bytes);
    return parseXml(text);
  }

  async getRootRelationshipsXml() {
    const bytes = await this.zip.getEntryBytes(ROOT_RELS);
    const text = new TextDecoder("utf-8").decode(bytes);
    return parseXml(text);
  }

  async getOfficeDocumentPartUri() {
    const rels = await this.getRootRelationshipsXml();
    // Relationships xmlns is usually default:
    // http://schemas.openxmlformats.org/package/2006/relationships
    const relsRoot = rels.root;
    const relationships = [...relsRoot.descendantsByQName("Relationship")];
    // If default namespace is set, qname will include prefix in source; our parser keeps qnames.
    // Fallback: scan all descendants and match localName.
    const all = relationships.length ? relationships : [...relsRoot.descendants()].filter((e) => e.nameParts().local === "Relationship");
    for (const rel of all) {
      const type = rel.getAttribute("Type");
      const target = rel.getAttribute("Target");
      if (type === OFFICE_DOCUMENT_REL_TYPE) return normalizePartUri(target);
    }
    return null;
  }

  async isWordprocessingDocument() {
    try {
      const office = await this.getOfficeDocumentPartUri();
      return office === "/word/document.xml";
    } catch {
      return false;
    }
  }

  async assertIsValidOpc() {
    if (!this.zip.getEntry(CONTENT_TYPES)) {
      throw new OpenXmlPowerToolsError("OXPT_INVALID_DOCX", "Missing [Content_Types].xml");
    }
    if (!this.zip.getEntry(ROOT_RELS)) {
      throw new OpenXmlPowerToolsError("OXPT_INVALID_DOCX", "Missing _rels/.rels");
    }
  }

  async toBytes({ replaceParts, adapter, deflateLevel } = {}) {
    const resolvedAdapter = adapter ?? (await getDefaultZipAdapter());
    const byName = normalizeReplaceParts(replaceParts);

    const entries = await Promise.all(
      this.zip.entries.map(async (e) => {
        const name = e.name;
        const replacement = byName.get(name);
        const isDir = name.endsWith("/");
        const bytes = replacement ?? (isDir ? new Uint8Array() : await e.getBytes());
        return { name, bytes, compressionMethod: isDir ? 0 : 8 };
      }),
    );

    // Include any new entries supplied in replaceParts but not present in the original.
    for (const [name, bytes] of byName) {
      if (this.zip.getEntry(name)) continue;
      const isDir = name.endsWith("/");
      entries.push({ name, bytes, compressionMethod: isDir ? 0 : 8 });
    }

    return ZipArchive.build(entries, { adapter: resolvedAdapter, level: deflateLevel ?? 6 });
  }
}

function normalizePartUri(target) {
  if (!target) return null;
  if (target.startsWith("/")) return target;
  return `/${target}`;
}

function normalizeReplaceParts(replaceParts) {
  const byName = new Map();
  if (!replaceParts) return byName;
  for (const [key, value] of replaceParts instanceof Map ? replaceParts : Object.entries(replaceParts)) {
    const name = key.startsWith("/") ? key.slice(1) : key;
    byName.set(name, value instanceof Uint8Array ? value : new Uint8Array(value));
  }
  return byName;
}
