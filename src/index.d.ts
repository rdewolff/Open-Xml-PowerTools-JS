export class OpenXmlPowerToolsError extends Error {
  code: string;
  cause?: unknown;
  constructor(code: string, message: string, cause?: unknown);
}

export type BytesLike = Uint8Array | ArrayBuffer | ArrayBufferView;

export type OpenXmlType = "unknown" | "docx" | "xlsx" | "pptx" | "opc";

export interface ZipAdapter {
  inflateRaw(data: Uint8Array): Uint8Array | Promise<Uint8Array>;
  deflateRaw(data: Uint8Array, options?: { level?: number }): Uint8Array | Promise<Uint8Array>;
}

export interface DocumentOptions {
  fileName?: string;
  zipAdapter?: ZipAdapter;
}

export class OpenXmlPowerToolsDocument {
  constructor(bytes: BytesLike, options?: DocumentOptions);
  static fromBytes(bytes: BytesLike, options?: DocumentOptions): OpenXmlPowerToolsDocument;
  static fromBase64(base64: string, options?: DocumentOptions): OpenXmlPowerToolsDocument;

  bytes: Uint8Array;
  fileName?: string;
  zipAdapter?: ZipAdapter;

  toBytes(): Uint8Array;
  toBase64(): string;
  detectType(): Promise<OpenXmlType>;
}

export class WmlDocument extends OpenXmlPowerToolsDocument {
  static fromBytes(bytes: BytesLike, options?: DocumentOptions): WmlDocument;

  getPartBytes(uri: string): Promise<Uint8Array>;
  getPartText(uri: string): Promise<string>;
  getMainDocumentXml(): Promise<unknown>;
  getMainDocumentText(): Promise<{ paragraphs: string[]; text: string }>;

  searchAndReplace(search: string, replace: string, matchCase?: boolean): Promise<WmlDocument>;
  simplifyMarkup(settings?: unknown): Promise<WmlDocument>;
  acceptRevisions(): Promise<WmlDocument>;
  convertToHtml(settings?: unknown): Promise<unknown>;

  replacePartXml(partUri: string, xmlDocumentOrElement: unknown): Promise<WmlDocument>;
  replaceParts(replaceParts: Record<string, Uint8Array>, options?: { adapter?: ZipAdapter; deflateLevel?: number }): Promise<WmlDocument>;
}

export const TextReplacer: {
  searchAndReplace(
    doc: WmlDocument,
    search: string,
    replace: string,
    options?: { matchCase?: boolean },
  ): Promise<WmlDocument>;
};

export const MarkupSimplifier: {
  simplifyMarkup(doc: WmlDocument, settings?: unknown): Promise<WmlDocument>;
};

export const RevisionAccepter: {
  acceptRevisions(doc: WmlDocument): Promise<WmlDocument>;
  hasTrackedRevisions(doc: WmlDocument): Promise<boolean>;
};

export const WmlToHtmlConverter: {
  convertToHtml(doc: WmlDocument, settings?: unknown): Promise<{
    html: string;
    cssText: string;
    warnings: unknown[];
    htmlElement?: unknown;
  }>;
};

export const HtmlConverter: {
  convertToHtml: typeof WmlToHtmlConverter.convertToHtml;
};

export const HtmlToWmlConverter: {
  convertHtmlToWml(
    defaultCss: string,
    authorCss: string,
    userCss: string,
    xhtml: string | unknown,
    settings?: { fileName?: string } & Record<string, unknown>,
    templateDoc?: WmlDocument | null,
    annotatedHtmlDumpFileName?: string | null,
  ): Promise<WmlDocument>;
};

