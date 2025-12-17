import { OpenXmlPowerToolsError } from "../open-xml-powertools-error.js";

const MAX_BITS = 15;

const LENGTH_BASE = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195,
  227, 258,
];
const LENGTH_EXTRA = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
];

const DIST_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097,
  6145, 8193, 12289, 16385, 24577,
];
const DIST_EXTRA = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
];

export function inflateRaw(data, options = {}) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const expectedSize = options.expectedSize;
  if (expectedSize != null && (typeof expectedSize !== "number" || !Number.isFinite(expectedSize) || expectedSize < 0)) {
    throw new OpenXmlPowerToolsError("OXPT_INVALID_ARGUMENT", "expectedSize must be a non-negative finite number");
  }

  const reader = new BitReader(bytes);
  let output = new Uint8Array(initialCapacity(expectedSize));
  let outPos = 0;

  let isFinal = false;
  while (!isFinal) {
    isFinal = reader.readBits(1) === 1;
    const type = reader.readBits(2);

    if (type === 0) {
      reader.alignToByte();
      const len = reader.readU16le();
      const nlen = reader.readU16le();
      if (((len ^ 0xffff) & 0xffff) !== nlen) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Invalid stored block length");
      }
      output = ensureCapacity(output, outPos + len);
      reader.readBytesInto(output, outPos, len);
      outPos += len;
      continue;
    }

    const { litLen, dist } = type === 1 ? getFixedTrees() : type === 2 ? readDynamicTrees(reader) : { litLen: null, dist: null };
    if (!litLen || !dist) {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", `Unsupported DEFLATE block type: ${type}`);
    }

    while (true) {
      const symbol = litLen.decode(reader);
      if (symbol < 256) {
        output = ensureCapacity(output, outPos + 1);
        output[outPos++] = symbol;
        continue;
      }
      if (symbol === 256) break; // end-of-block
      if (symbol < 257 || symbol > 285) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", `Invalid length symbol: ${symbol}`);
      }

      const lenIndex = symbol - 257;
      const extraLenBits = LENGTH_EXTRA[lenIndex];
      const length = LENGTH_BASE[lenIndex] + (extraLenBits ? reader.readBits(extraLenBits) : 0);

      const distSymbol = dist.decode(reader);
      if (distSymbol < 0 || distSymbol >= DIST_BASE.length) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", `Invalid distance symbol: ${distSymbol}`);
      }
      const extraDistBits = DIST_EXTRA[distSymbol];
      const distance = DIST_BASE[distSymbol] + (extraDistBits ? reader.readBits(extraDistBits) : 0);
      if (distance > outPos) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Distance exceeds output size");
      }

      output = ensureCapacity(output, outPos + length);
      for (let i = 0; i < length; i++) {
        output[outPos] = output[outPos - distance];
        outPos++;
      }
    }
  }

  const result = output.subarray(0, outPos);
  if (expectedSize != null && result.length !== expectedSize) {
    throw new OpenXmlPowerToolsError(
      "OXPT_ZIP_INVALID",
      `Inflated size mismatch (expected ${expectedSize}, got ${result.length})`,
    );
  }
  return new Uint8Array(result);
}

class BitReader {
  constructor(bytes) {
    this.bytes = bytes;
    this.pos = 0;
    this.bitBuf = 0;
    this.bitLen = 0;
  }

  alignToByte() {
    const drop = this.bitLen & 7;
    if (drop) this.readBits(drop);
  }

  readU16le() {
    this.alignToByte();
    if (this.pos + 2 > this.bytes.length) {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Unexpected end of DEFLATE stream");
    }
    const v = this.bytes[this.pos] | (this.bytes[this.pos + 1] << 8);
    this.pos += 2;
    return v;
  }

  readBytesInto(target, offset, length) {
    this.alignToByte();
    if (this.pos + length > this.bytes.length) {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Unexpected end of DEFLATE stream");
    }
    target.set(this.bytes.subarray(this.pos, this.pos + length), offset);
    this.pos += length;
  }

  readBits(n) {
    if (n === 0) return 0;
    while (this.bitLen < n) {
      if (this.pos >= this.bytes.length) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Unexpected end of DEFLATE stream");
      }
      this.bitBuf |= this.bytes[this.pos++] << this.bitLen;
      this.bitLen += 8;
    }
    const mask = (1 << n) - 1;
    const out = this.bitBuf & mask;
    this.bitBuf >>>= n;
    this.bitLen -= n;
    return out;
  }
}

class HuffmanTree {
  constructor(codeLengths) {
    const maxLen = Math.max(0, ...codeLengths);
    if (maxLen > MAX_BITS) {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", `Huffman code length exceeds ${MAX_BITS}`);
    }

    const count = new Array(maxLen + 1).fill(0);
    for (const len of codeLengths) if (len > 0) count[len]++;

    const nextCode = new Array(maxLen + 1).fill(0);
    let code = 0;
    for (let bits = 1; bits <= maxLen; bits++) {
      code = (code + count[bits - 1]) << 1;
      nextCode[bits] = code;
    }

    this.left = [-1];
    this.right = [-1];
    this.symbol = [-1];

    for (let sym = 0; sym < codeLengths.length; sym++) {
      const len = codeLengths[sym];
      if (!len) continue;
      const c = nextCode[len]++;
      const wireCode = reverseBits(c, len);
      this.insert(sym, wireCode, len);
    }
  }

  insert(symbol, code, length) {
    let node = 0;
    for (let i = 0; i < length; i++) {
      if (this.symbol[node] !== -1) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Huffman code extends past a leaf");
      }
      const bit = (code >> i) & 1;
      const next = bit ? this.right[node] : this.left[node];
      if (next === -1) {
        const idx = this.left.length;
        this.left.push(-1);
        this.right.push(-1);
        this.symbol.push(-1);
        if (bit) this.right[node] = idx;
        else this.left[node] = idx;
        node = idx;
      } else {
        node = next;
      }
    }
    if (this.symbol[node] !== -1) {
      throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Overlapping Huffman code");
    }
    this.symbol[node] = symbol;
  }

  decode(reader) {
    let node = 0;
    while (true) {
      const sym = this.symbol[node];
      if (sym !== -1) return sym;
      const bit = reader.readBits(1);
      node = bit ? this.right[node] : this.left[node];
      if (node === -1) {
        throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Invalid Huffman code");
      }
    }
  }
}

let cachedFixedTrees = null;
function getFixedTrees() {
  if (cachedFixedTrees) return cachedFixedTrees;

  const litLenLengths = new Array(288).fill(0);
  for (let i = 0; i <= 143; i++) litLenLengths[i] = 8;
  for (let i = 144; i <= 255; i++) litLenLengths[i] = 9;
  for (let i = 256; i <= 279; i++) litLenLengths[i] = 7;
  for (let i = 280; i <= 287; i++) litLenLengths[i] = 8;

  const distLengths = new Array(32).fill(5);

  cachedFixedTrees = {
    litLen: new HuffmanTree(litLenLengths),
    dist: new HuffmanTree(distLengths),
  };
  return cachedFixedTrees;
}

function readDynamicTrees(reader) {
  const hlit = reader.readBits(5) + 257;
  const hdist = reader.readBits(5) + 1;
  const hclen = reader.readBits(4) + 4;

  const order = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
  const clen = new Array(19).fill(0);
  for (let i = 0; i < hclen; i++) clen[order[i]] = reader.readBits(3);
  const codeLenTree = new HuffmanTree(clen);

  const lengths = [];
  while (lengths.length < hlit + hdist) {
    const sym = codeLenTree.decode(reader);
    if (sym <= 15) {
      lengths.push(sym);
      continue;
    }
    if (sym === 16) {
      if (lengths.length === 0) throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "Repeat code with no previous");
      const repeat = reader.readBits(2) + 3;
      const prev = lengths[lengths.length - 1];
      for (let i = 0; i < repeat; i++) lengths.push(prev);
      continue;
    }
    if (sym === 17) {
      const repeat = reader.readBits(3) + 3;
      for (let i = 0; i < repeat; i++) lengths.push(0);
      continue;
    }
    if (sym === 18) {
      const repeat = reader.readBits(7) + 11;
      for (let i = 0; i < repeat; i++) lengths.push(0);
      continue;
    }
    throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", `Invalid code length symbol: ${sym}`);
  }

  const litLenLengths = lengths.slice(0, hlit);
  const distLengths = lengths.slice(hlit, hlit + hdist);

  // If all distance code lengths are 0, the stream is invalid.
  if (!distLengths.some((x) => x !== 0)) {
    throw new OpenXmlPowerToolsError("OXPT_ZIP_INVALID", "No distance codes");
  }

  return {
    litLen: new HuffmanTree(litLenLengths),
    dist: new HuffmanTree(distLengths),
  };
}

function initialCapacity(expectedSize) {
  if (expectedSize != null) return Math.max(256, expectedSize);
  return 1024;
}

function ensureCapacity(buf, required) {
  if (required <= buf.length) return buf;
  let size = buf.length ? buf.length : 1;
  while (size < required) size *= 2;
  const next = new Uint8Array(size);
  next.set(buf);
  return next;
}

function reverseBits(code, length) {
  let out = 0;
  for (let i = 0; i < length; i++) {
    out = (out << 1) | (code & 1);
    code >>>= 1;
  }
  return out;
}
