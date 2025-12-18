import type { WmlDocument } from "./index.js";

export function readWmlDocument(path: string): Promise<WmlDocument>;
export function writeWmlDocument(path: string, doc: WmlDocument): Promise<void>;

