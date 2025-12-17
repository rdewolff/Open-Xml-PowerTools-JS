# Open-Xml-PowerTools-JS (DOCX-focused) — API Spec + Plan

## Goals

- Provide a **zero-dependency** JavaScript port of the **DOCX (WordprocessingML)** parts of `Open-Xml-PowerTools`, with a **similar public API** to the C# library.
- Run in **browsers and Node.js** (and other JS runtimes) with no npm deps.
- First-class scenario: **high-fidelity DOCX ⇄ HTML/CSS** conversion (starting with DOCX → HTML).
- Keep the core API **byte-oriented** (no filesystem assumptions): inputs/outputs are `Uint8Array` (or `ArrayBuffer`).

## Non-goals (initially)

- XLSX/PPTX support.
- A full Open XML SDK analog (strongly-typed parts, schema-driven classes, etc.).
- Perfect round-trip HTML ⇄ DOCX in v0.x (it’s a roadmap item).

## Compatibility model

- **No external dependencies.** “No deps” means no npm packages; using **runtime built-ins** is allowed.
- DOCX is a ZIP/OPC package; deflate/inflate is required.
  - Browser: use `CompressionStream` / `DecompressionStream` when available.
  - Node: use `node:zlib` when available.
  - Other runtimes: user supplies a small adapter (see `ZipAdapter`).

## Core data types

- `Bytes`: `Uint8Array` (accept `ArrayBuffer` where convenient).
- `XmlNode` / `XmlElement`: a small internal XML tree (LINQ-to-XML-like) that works in all runtimes (Node has no built-in DOMParser).
- `OpenXmlPartUri`: string like `"/word/document.xml"`.

## Package entry points (proposed)

- ESM-first:
  - `src/index.js` exports the public API.
  - Secondary entry: `src/node.js` exposes optional Node helpers (fs I/O), still “no deps”.

Public namespace mirrors C# names where it helps discoverability:

```js
import {
  OpenXmlPowerToolsDocument,
  WmlDocument,
  WmlToHtmlConverter,
  HtmlConverter, // compatibility alias
  MarkupSimplifier,
  RevisionAccepter,
  TextReplacer,
} from "open-xml-powertools-js";
```

## Documents

### `class OpenXmlPowerToolsDocument`

Minimal base for “a blob that is an Open XML package”.

```ts
class OpenXmlPowerToolsDocument {
  fileName?: string;
  bytes: Uint8Array;

  constructor(bytes: Bytes, options?: { fileName?: string });

  static fromBytes(bytes: Bytes, options?: { fileName?: string }): OpenXmlPowerToolsDocument;
  static fromBase64(base64: string, options?: { fileName?: string }): OpenXmlPowerToolsDocument;

  toBytes(): Uint8Array;        // returns a copy for immutability
  toBase64(): string;

  detectType(): "docx" | "pptx" | "xlsx" | "opc" | "unknown"; // v0 supports docx/opc
}
```

### `class WmlDocument extends OpenXmlPowerToolsDocument`

DOCX-specialized conveniences, roughly analogous to C# `WmlDocument`.

```ts
class WmlDocument extends OpenXmlPowerToolsDocument {
  static fromBytes(bytes: Bytes, options?: { fileName?: string }): WmlDocument;

  // Convenience wrappers mirroring C# instance methods:
  async convertToHtml(settings?: WmlToHtmlConverterSettings): Promise<HtmlResult>;
  async simplifyMarkup(settings: SimplifyMarkupSettings): Promise<WmlDocument>;
  async searchAndReplace(search: string, replace: string, matchCase?: boolean): Promise<WmlDocument>;
  async acceptRevisions(): Promise<WmlDocument>;
}
```

Node-only helper (optional, separate entry):

```js
import { WmlDocument } from "open-xml-powertools-js";
import { readWmlDocument, writeWmlDocument } from "open-xml-powertools-js/node";
```

## DOCX → HTML conversion

### Settings (mirrors C# `WmlToHtmlConverterSettings` / `HtmlConverterSettings`)

```ts
type WmlToHtmlConverterSettings = {
  pageTitle?: string;
  cssClassPrefix?: string;                 // default "pt-"
  fabricateCssClasses?: boolean;           // default true
  generalCss?: string;                     // default `span { white-space: pre-wrap; }`
  additionalCss?: string;                  // appended after generated CSS
  restrictToSupportedLanguages?: boolean;  // default false
  restrictToSupportedNumberingFormats?: boolean; // default false

  // List numbering text generation, keyed by locale (e.g., "en-US", "fr-FR").
  // Signature matches the C# intent: (listLevelText, levelNumber, numFmt) => renderedText
  listItemImplementations?: Record<
    string,
    (levelText: string, levelNumber: number, numFmt: string) => string
  >;

  // Image emission hook; return null to skip emitting an <img>.
  imageHandler?: (info: ImageInfo) => HtmlImageResult | null;

  // Advanced: control how the result HTML is produced (string vs XmlElement)
  output?: { format?: "string" | "xml" };
};

type ImageInfo = {
  contentType: string;          // e.g. "image/png"
  bytes: Uint8Array;            // decoded image bytes (original)
  widthEmus?: number;
  heightEmus?: number;
  altText?: string | null;
  suggestedStyle?: string | null; // e.g. `style="..."` equivalent
};

type HtmlImageResult =
  | { src: string; attrs?: Record<string, string> } // library builds <img>
  | { element: string };                            // raw <img ...> string, power-user escape hatch
```

### Converter surface

```ts
type HtmlResult = {
  html: string;                  // complete HTML5 document string by default
  cssText: string;               // generated + general + additional (also embedded in html)
  warnings: Array<{ code: string; message: string; part?: string }>;
};

const WmlToHtmlConverter: {
  convertToHtml(doc: WmlDocument, settings?: WmlToHtmlConverterSettings): Promise<HtmlResult>;
};

// Compatibility names (as in C#):
type HtmlConverterSettings = WmlToHtmlConverterSettings;
const HtmlConverter: {
  convertToHtml(doc: WmlDocument, settings?: HtmlConverterSettings): Promise<HtmlResult>;
};
```

### Default image behavior

- If `imageHandler` is not set, images are emitted as **data URLs** (`src="data:..."`) so the result is self-contained and browser-friendly.
- If `imageHandler` is set, the callback decides `src` (e.g., `"/assets/image1.png"`), enabling the classic C# pattern of writing image files to an adjacent folder.

## DOCX mutation utilities (DOCX-focused subset)

These are the most-used Open-Xml-PowerTools operations for preprocessing prior to HTML conversion and for template-style tasks.

### `MarkupSimplifier`

```ts
type SimplifyMarkupSettings = {
  acceptRevisions?: boolean;
  normalizeXml?: boolean;
  removeBookmarks?: boolean;
  removeComments?: boolean;
  removeContentControls?: boolean;
  removeEndAndFootNotes?: boolean;
  removeFieldCodes?: boolean;
  removeGoBackBookmark?: boolean;
  removeHyperlinks?: boolean;
  removeLastRenderedPageBreak?: boolean;
  removeMarkupForDocumentComparison?: boolean;
  removePermissions?: boolean;
  removeProof?: boolean;
  removeRsidInfo?: boolean;
  removeSmartTags?: boolean;
  removeSoftHyphens?: boolean;
  removeWebHidden?: boolean;
  replaceTabsWithSpaces?: boolean;
};

const MarkupSimplifier: {
  simplifyMarkup(doc: WmlDocument, settings: SimplifyMarkupSettings): Promise<WmlDocument>;
};
```

### `RevisionAccepter`

```ts
const RevisionAccepter: {
  acceptRevisions(doc: WmlDocument): Promise<WmlDocument>;
  hasTrackedRevisions(doc: WmlDocument): Promise<boolean>;
};
```

### `TextReplacer`

Mirrors C# `TextReplacer.SearchAndReplace` (run-splitting across `<w:t>`).

```ts
const TextReplacer: {
  searchAndReplace(
    doc: WmlDocument,
    search: string,
    replace: string,
    options?: { matchCase?: boolean }
  ): Promise<WmlDocument>;
};
```

## Errors

```ts
class OpenXmlPowerToolsError extends Error {
  code: string;          // e.g. "OXPT_ZIP_UNSUPPORTED", "OXPT_INVALID_DOCX"
  details?: unknown;
}
```

## Minimal example (browser or Node)

```js
import { WmlDocument, WmlToHtmlConverter } from "open-xml-powertools-js";

const doc = WmlDocument.fromBytes(docxBytes, { fileName: "input.docx" });
const { html } = await WmlToHtmlConverter.convertToHtml(doc, {
  additionalCss: "body { margin: 1cm auto; max-width: 20cm; }",
});
```

## API mapping to C# (initial subset)

- `WmlDocument` → `WmlDocument`
- `WmlDocument.ConvertToHtml(settings)` → `await doc.convertToHtml(settings)`
- `WmlToHtmlConverter.ConvertToHtml(doc, settings)` → `await WmlToHtmlConverter.convertToHtml(doc, settings)`
- `HtmlConverter` / `HtmlConverterSettings` → alias to `WmlToHtmlConverter` API
- `MarkupSimplifier.SimplifyMarkup` → `await MarkupSimplifier.simplifyMarkup(...)`
- `RevisionAccepter.AcceptRevisions` → `await RevisionAccepter.acceptRevisions(...)`
- `TextReplacer.SearchAndReplace` → `await TextReplacer.searchAndReplace(...)`

## Implementation plan (phased)

### Phase 0 — Golden-path DOCX parse + tests

- Add a **single minimal valid DOCX fixture** (one paragraph, one run, optional bold) and a **tiny image fixture** (optional) for deterministic tests.
- Implement the smallest end-to-end “open → parse main document XML → extract plain text” path that proves the architecture:
  - Load ZIP/OPC, locate `"[Content_Types].xml"`, `"_rels/.rels"`, and `"/word/document.xml"`.
  - Parse `"/word/document.xml"` and return expected results for the fixture (e.g., paragraph text array).
- Establish the **test suite mechanism** (no deps):
  - Use Node’s built-in runner: `node --test`.
  - Tests live in `test/**/*.test.js` and run in CI/local.

Definition of done:
- `node --test` passes and validates that the library can open a valid DOCX fixture and produce expected extracted text.

### Phase 1 — Core plumbing (OPC + XML)

- Implement a small ZIP reader/writer sufficient for DOCX:
  - Read central directory, list entries, extract bytes.
  - Write updated entries, rebuild central directory.
  - Deflate/inflate via `ZipAdapter` (built-in adapters for Browser + Node).
- Implement a minimal XML parser/serializer with namespaces:
  - Preserve attribute order where feasible (not required for validity).
  - Provide helpers for querying by `{namespace}localName` and by prefix mappings.
- Build an OPC package layer:
  - Parts (`/word/document.xml`, `/word/styles.xml`, etc.)
  - Relationships (`_rels/.rels`, `word/_rels/document.xml.rels`)
  - Content types (`[Content_Types].xml`)

Definition of done:
- Can open a DOCX, read `word/document.xml`, round-trip without changes.

### Phase 2 — `WmlDocument` + basic transforms

- Implement `WmlDocument` on top of the package layer.
- Implement `MarkupSimplifier` subset required by HTML conversion pipeline:
  - remove rsid, comments, smart tags, content controls, etc.
- Implement `RevisionAccepter` (initially: detect + basic accept for common revision markup).
- Implement `TextReplacer.searchAndReplace` for main document part.

Definition of done:
- Can apply `searchAndReplace` and re-save a valid DOCX.

### Phase 3 — DOCX → HTML v0 (structure-first)

- Port the high-level pipeline used in C# `WmlToHtmlConverter.ConvertToHtml`:
  - accept revisions (optional), simplify markup, assemble formatting (subset), then transform to XHTML.
- Start with core WordprocessingML constructs:
  - paragraphs/runs, basic character formatting (b/i/u), hyperlinks, headings, lists, tables, breaks.
- Implement image extraction and `imageHandler` hook; default to data URLs.

Definition of done:
- Common DOCX files render readable HTML with correct paragraph/list/table structure.

### Phase 4 — Fidelity + CSS parity

- Expand formatting assembler to match C# behavior more closely:
  - styles inheritance, numbering formats, bidi/rtl considerations, borders, spacing, tabs.
- Expand supported elements: footnotes/endnotes, fields (as text), content controls (optional), comments (optional).
- Add “warnings” stream with actionable codes for unsupported features.

Definition of done:
- Side-by-side comparison with C# converter is “close” on representative `TestFiles` DOCX set.

### Phase 5 — HTML → DOCX (roadmap)

- Port `HtmlToWmlConverter` surface:
  - `convertHtmlToWml(defaultCss, authorCss, userCss, htmlOrXhtml, settings, templateDoc?)`
- Implement a small HTML parser (or accept a restricted XHTML-in-XML input first).

Definition of done:
- Can generate a simple DOCX from HTML with paragraphs, runs, lists, tables, and images.

## Testing strategy

- Primary runner: Node built-in `node:test` (`node --test`), keeping the library dependency-free.
- Tests cover:
  - ZIP/OPC: open fixture, list parts, read bytes, write/repack and re-open.
  - XML: namespace parsing/serialization invariants needed by transforms.
  - `WmlDocument`: extract main document text for fixture(s).
  - Conversion: `WmlToHtmlConverter.convertToHtml` produces expected structure for fixture(s).
- For conversion assertions:
  - Prefer **structure checks** (e.g., contains `<p>` with expected text) over brittle full-string equality.
  - If snapshots are used, normalize HTML (whitespace + attribute ordering) via an internal canonicalizer.

## Open questions (to resolve early)

- Should `convertToHtml` default to returning a full HTML document string, or just the `<html>` element XHTML string?
- How strict should XML round-tripping be (whitespace preservation vs normalization)?
- What is the minimum acceptable revision-accepting support for v0 (enough for converter preprocessing vs full fidelity)?
