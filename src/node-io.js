import { readFile, writeFile } from "node:fs/promises";
import { WmlDocument } from "./wml-document.js";

export async function readWmlDocument(path) {
  const bytes = await readFile(path);
  return new WmlDocument(new Uint8Array(bytes), { fileName: path });
}

export async function writeWmlDocument(path, doc) {
  await writeFile(path, Buffer.from(doc.toBytes()));
}

