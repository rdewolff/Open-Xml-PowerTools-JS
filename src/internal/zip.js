import { OpenXmlPowerToolsError } from "../open-xml-powertools-error.js";
import { u16le, u32le, writeU16le, writeU32le } from "./binary.js";
import { crc32 } from "./crc32.js";

const SIG_EOCD = 0x06054b50;
const SIG_CEN = 0x02014b50;
const SIG_LOC = 0x04034b50;

export class ZipArchive {
  constructor(entries) {
    this.entries = entries;
    this.byName = new Map(entries.map((e) => [e.name, e]));
  }

  static async fromBytes(bytes, { adapter }) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const eocdOffset = findEndOfCentralDirectory(data);
    const cdSize = u32le(data, eocdOffset + 12);
    const cdOffset = u32le(data, eocdOffset + 16);
    const entryCount = u16le(data, eocdOffset + 10);

    if (cdOffset + cdSize > data.length) {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Central directory out of bounds");
    }

    const entries = [];
    let cursor = cdOffset;
    for (let i = 0; i < entryCount; i++) {
      if (u32le(data, cursor) !== SIG_CEN) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Invalid central directory signature");
      }

      const compressionMethod = u16le(data, cursor + 10);
      const crc = u32le(data, cursor + 16);
      const compressedSize = u32le(data, cursor + 20);
      const uncompressedSize = u32le(data, cursor + 24);
      const fileNameLength = u16le(data, cursor + 28);
      const extraFieldLength = u16le(data, cursor + 30);
      const fileCommentLength = u16le(data, cursor + 32);
      const localHeaderOffset = u32le(data, cursor + 42);

      const nameStart = cursor + 46;
      const nameBytes = data.subarray(nameStart, nameStart + fileNameLength);
      const name = new TextDecoder("utf-8").decode(nameBytes);

      cursor = nameStart + fileNameLength + extraFieldLength + fileCommentLength;

      entries.push(
        new ZipEntry({
          name,
          compressionMethod,
          crc,
          compressedSize,
          uncompressedSize,
          localHeaderOffset,
          _zipBytes: data,
          _adapter: adapter,
        }),
      );
    }

    return new ZipArchive(entries);
  }

  getEntry(name) {
    return this.byName.get(name) ?? null;
  }

  async getEntryBytes(name) {
    const entry = this.getEntry(name);
    if (!entry) throw new OpenXmlPowerToolsError("OXPT_ZIP_NOT_FOUND", `Entry not found: ${name}`);
    return entry.getBytes();
  }

  static async build(entries, { adapter, level } = {}) {
    const encoder = new TextEncoder();
    const normalized = entries.map((e) => ({
      name: e.name,
      bytes: e.bytes instanceof Uint8Array ? e.bytes : new Uint8Array(e.bytes),
      compressionMethod: e.compressionMethod ?? 8,
    }));

    const fileRecords = [];
    let offset = 0;

    for (const entry of normalized) {
      const nameBytes = encoder.encode(entry.name);
      const uncompressed = entry.bytes;
      const method = entry.compressionMethod;
      let compressed = uncompressed;
      if (method === 8) {
        if (!adapter?.deflateRaw) {
          throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "Deflate adapter not provided");
        }
        compressed = adapter.deflateRaw(uncompressed, { level });
      } else if (method !== 0) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", `Unsupported compression method: ${method}`);
      }

      const crc = crc32(uncompressed);

      const localHeader = new Uint8Array(30 + nameBytes.length);
      writeU32le(localHeader, 0, SIG_LOC);
      writeU16le(localHeader, 4, 20); // version needed
      writeU16le(localHeader, 6, 0); // flags
      writeU16le(localHeader, 8, method);
      writeU16le(localHeader, 10, 0); // time
      writeU16le(localHeader, 12, 0); // date
      writeU32le(localHeader, 14, crc);
      writeU32le(localHeader, 18, compressed.length);
      writeU32le(localHeader, 22, uncompressed.length);
      writeU16le(localHeader, 26, nameBytes.length);
      writeU16le(localHeader, 28, 0); // extra
      localHeader.set(nameBytes, 30);

      const localHeaderOffset = offset;
      offset += localHeader.length + compressed.length;

      fileRecords.push({
        name: entry.name,
        nameBytes,
        method,
        crc,
        compressed,
        uncompressedSize: uncompressed.length,
        compressedSize: compressed.length,
        localHeader,
        localHeaderOffset,
      });
    }

    const centralDirectoryStart = offset;
    const centralRecords = [];
    for (const fr of fileRecords) {
      const cen = new Uint8Array(46 + fr.nameBytes.length);
      writeU32le(cen, 0, SIG_CEN);
      writeU16le(cen, 4, 20); // version made by
      writeU16le(cen, 6, 20); // version needed
      writeU16le(cen, 8, 0); // flags
      writeU16le(cen, 10, fr.method);
      writeU16le(cen, 12, 0); // time
      writeU16le(cen, 14, 0); // date
      writeU32le(cen, 16, fr.crc);
      writeU32le(cen, 20, fr.compressedSize);
      writeU32le(cen, 24, fr.uncompressedSize);
      writeU16le(cen, 28, fr.nameBytes.length);
      writeU16le(cen, 30, 0); // extra
      writeU16le(cen, 32, 0); // comment
      writeU16le(cen, 34, 0); // disk
      writeU16le(cen, 36, 0); // internal attrs
      writeU32le(cen, 38, 0); // external attrs
      writeU32le(cen, 42, fr.localHeaderOffset);
      cen.set(fr.nameBytes, 46);
      centralRecords.push(cen);
      offset += cen.length;
    }
    const centralDirectorySize = offset - centralDirectoryStart;

    const eocd = new Uint8Array(22);
    writeU32le(eocd, 0, SIG_EOCD);
    writeU16le(eocd, 4, 0); // disk
    writeU16le(eocd, 6, 0); // cd start disk
    writeU16le(eocd, 8, fileRecords.length);
    writeU16le(eocd, 10, fileRecords.length);
    writeU32le(eocd, 12, centralDirectorySize);
    writeU32le(eocd, 16, centralDirectoryStart);
    writeU16le(eocd, 20, 0); // comment length
    offset += eocd.length;

    const out = new Uint8Array(offset);
    let outCursor = 0;
    for (const fr of fileRecords) {
      out.set(fr.localHeader, outCursor);
      outCursor += fr.localHeader.length;
      out.set(fr.compressed, outCursor);
      outCursor += fr.compressed.length;
    }
    for (const cen of centralRecords) {
      out.set(cen, outCursor);
      outCursor += cen.length;
    }
    out.set(eocd, outCursor);
    return out;
  }
}

export class ZipEntry {
  constructor({
    name,
    compressionMethod,
    crc,
    compressedSize,
    uncompressedSize,
    localHeaderOffset,
    _zipBytes,
    _adapter,
  }) {
    this.name = name;
    this.compressionMethod = compressionMethod;
    this.crc = crc;
    this.compressedSize = compressedSize;
    this.uncompressedSize = uncompressedSize;
    this._localHeaderOffset = localHeaderOffset;
    this._zipBytes = _zipBytes;
    this._adapter = _adapter;
  }

  async getBytes() {
    const data = this._zipBytes;
    const loc = this._localHeaderOffset;
    if (u32le(data, loc) !== SIG_LOC) {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Invalid local header signature");
    }
    const fileNameLength = u16le(data, loc + 26);
    const extraFieldLength = u16le(data, loc + 28);
    const dataStart = loc + 30 + fileNameLength + extraFieldLength;
    const compressed = data.subarray(dataStart, dataStart + this.compressedSize);

    if (this.compressionMethod === 0) return new Uint8Array(compressed);
    if (this.compressionMethod === 8) {
      if (!this._adapter?.inflateRaw) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_UNSUPPORTED", "Inflate adapter not provided");
      }
      return this._adapter.inflateRaw(compressed);
    }
    throw new OpenXmlPowerToolsError(
      "OXPT_ZIP_UNSUPPORTED",
      `Unsupported compression method: ${this.compressionMethod}`,
    );
  }
}

function findEndOfCentralDirectory(bytes) {
  // EOCD record is at least 22 bytes and may have a variable-length comment at the end.
  const min = Math.max(0, bytes.length - 0xffff - 22);
  for (let i = bytes.length - 22; i >= min; i--) {
    if (u32le(bytes, i) !== SIG_EOCD) continue;
    const commentLength = u16le(bytes, i + 20);
    if (i + 22 + commentLength === bytes.length) return i;
  }
  throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "End of central directory not found");
}

