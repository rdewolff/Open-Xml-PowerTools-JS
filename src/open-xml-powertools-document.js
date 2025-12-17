import { bytesToBase64, base64ToBytes } from "./util/base64.js";
import { OpcPackage } from "./internal/opc.js";
import { OpenXmlPowerToolsError } from "./open-xml-powertools-error.js";

export { OpenXmlPowerToolsError };

export class OpenXmlPowerToolsDocument {
  constructor(bytes, options = {}) {
    if (!bytes) throw new OpenXmlPowerToolsError("OXPT_INVALID_ARGUMENT", "bytes is required");
    this.bytes = coerceToBytes(bytes);
    this.fileName = options.fileName;
    this.zipAdapter = options.zipAdapter;
  }

  static fromBytes(bytes, options) {
    return new OpenXmlPowerToolsDocument(bytes, options);
  }

  static fromBase64(base64, options) {
    return new OpenXmlPowerToolsDocument(base64ToBytes(base64), options);
  }

  toBytes() {
    return new Uint8Array(this.bytes);
  }

  toBase64() {
    return bytesToBase64(this.bytes);
  }

  async detectType() {
    const pkg = await OpcPackage.fromBytes(this.bytes, { adapter: this.zipAdapter });
    const officeUri = await pkg.getOfficeDocumentPartUri().catch(() => null);
    if (officeUri === "/word/document.xml") return "docx";
    if (officeUri === "/xl/workbook.xml") return "xlsx";
    if (officeUri === "/ppt/presentation.xml") return "pptx";
    if (officeUri) return "opc";
    // Could still be an OPC package without officeDocument relationship; treat as unknown.
    return "unknown";
  }
}

export function coerceToBytes(bytes) {
  if (bytes instanceof Uint8Array) return bytes;
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
  if (ArrayBuffer.isView(bytes) && bytes.buffer instanceof ArrayBuffer) {
    return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }
  throw new OpenXmlPowerToolsError("OXPT_INVALID_ARGUMENT", "Expected Uint8Array/ArrayBuffer");
}
