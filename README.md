# Open-Xml-PowerTools-JS

JavaScript (ESM) port of the **DOCX-focused** parts of Open XML PowerTools, built to run with **zero npm dependencies** in both **browsers** and **Node.js**.

This project is **not affiliated with, endorsed by, or sponsored by** Microsoft or the original Open-Xml-PowerTools authors.

This repo currently focuses on:
- **DOCX → HTML** conversion (structure-first, expanding fidelity over time)
- A small set of **DOCX transforms** used by the converter (revisions, markup simplification, text replace)
- A minimal **HTML (XHTML) → DOCX** path for simple content

## Status

- The API is usable, but still evolving.
- The implementation is dependency-free by design: ZIP/OPC + XML parsing/serialization are implemented in `./src/`.

## Install / Use

This project is currently designed to be used directly from source (or via git dependency) as an ESM module.

### Node.js

```js
import { readFile } from "node:fs/promises";
import { WmlDocument, WmlToHtmlConverter } from "./src/index.js";

const bytes = new Uint8Array(await readFile("input.docx"));
const doc = WmlDocument.fromBytes(bytes, { fileName: "input.docx" });

const { html, warnings } = await WmlToHtmlConverter.convertToHtml(doc, {
  additionalCss: "body { margin: 1cm auto; max-width: 20cm; }",
});

console.log(warnings);
await Bun?.write?.("output.html", html); // optional (Bun)
```

Node helper entry (optional):

```js
import { readWmlDocument } from "./src/node.js";
import { WmlToHtmlConverter } from "./src/index.js";

const doc = await readWmlDocument("input.docx");
const { html } = await WmlToHtmlConverter.convertToHtml(doc);
```

Run tests:

```sh
npm test
```

### Browser (ES Modules)

Browsers usually don’t allow ESM imports from `file://`. Use a local HTTP server:

```sh
python3 -m http.server
```

Then open:

- `http://localhost:8000/playground.html`

Minimal browser example:

```html
<input id="file" type="file" accept=".docx" />
<script type="module">
  import { WmlDocument, WmlToHtmlConverter } from "./src/index.js";

  document.getElementById("file").addEventListener("change", async (e) => {
    const f = e.target.files[0];
    const bytes = new Uint8Array(await f.arrayBuffer());
    const doc = WmlDocument.fromBytes(bytes, { fileName: f.name });
    const { html } = await WmlToHtmlConverter.convertToHtml(doc);
    document.open(); document.write(html); document.close();
  });
</script>
```

## API Overview

All public exports come from `./src/index.js`:

### Documents

- `OpenXmlPowerToolsDocument`: byte container + type detection
- `WmlDocument`: DOCX wrapper with convenience methods

```js
import { WmlDocument } from "./src/index.js";
const doc = WmlDocument.fromBytes(uint8Array, { fileName: "input.docx" });
```

### DOCX → HTML

```js
import { WmlToHtmlConverter } from "./src/index.js";

const result = await WmlToHtmlConverter.convertToHtml(doc, {
  pageTitle: "My Document",
  additionalCss: "body { max-width: 20cm; margin: 1cm auto; }",

  // Optional: customize list markers
  listItemImplementations: {
    default: (_lvlText, levelNumber, _numFmt) => `#${levelNumber}`,
  },

  // Optional: control image output
  // If not set, images are embedded as data URLs.
  imageHandler: null,

  // Optional: include a lightweight XML object for the produced HTML
  output: { format: "xml" }, // also returns `htmlElement`
});

console.log(result.html);
console.log(result.warnings);
```

Current converter coverage (high-level):
- paragraphs/runs, basic formatting (`b/i/u`)
- headings via `Heading1..Heading6` styles
- hyperlinks (external)
- lists via `numbering.xml` (basic)
- tables (including `gridSpan`/`vMerge` and basic borders)
- images (data URLs by default)
- footnotes/endnotes (references + appended section)

### Transforms (DOCX mutation)

```js
import { MarkupSimplifier, RevisionAccepter, TextReplacer } from "./src/index.js";

const noRevs = await RevisionAccepter.acceptRevisions(doc);
const simplified = await MarkupSimplifier.simplifyMarkup(noRevs, {
  removeComments: true,
  removeContentControls: true,
  removeRsidInfo: true,
  removeGoBackBookmark: true,
});

const replaced = await TextReplacer.searchAndReplace(simplified, "Hello", "Hi", { matchCase: false });
```

### HTML (XHTML) → DOCX

This is intentionally minimal and currently expects well-formed XML (XHTML-like).

```js
import { HtmlToWmlConverter } from "./src/index.js";

const xhtml = `<?xml version="1.0"?>
<html><body>
  <h1>Title</h1>
  <p>Hello <strong>World</strong><br/>Line2</p>
</body></html>`;

const newDoc = await HtmlToWmlConverter.convertHtmlToWml("", "", "", xhtml, {});
```

## Runtime ZIP support (no deps)

DOCX is a ZIP/OPC package and requires deflate/inflate:

- **Node.js**: uses `node:zlib` automatically (lazy import)
- **Browsers**: uses built-in `CompressionStream` / `DecompressionStream` when available
- Other runtimes: pass a `zipAdapter` when constructing documents:

```js
const doc = WmlDocument.fromBytes(bytes, {
  zipAdapter: { inflateRaw: async (u8) => ..., deflateRaw: async (u8) => ... }
});
```

## Development

- Tests: `node --test` (run via `npm test`)
- CI: GitHub Actions workflow runs tests on push/PR (`.github/workflows/test.yml`)

Implementation layout:

- `src/internal/zip*.js`: ZIP reader/writer + adapters
- `src/internal/xml.js`: minimal XML parse/serialize
- `src/internal/opc.js`: OPC package access
- `src/wml-to-html-converter.js`: DOCX → HTML
- `src/html-to-wml-converter.js`: XHTML → DOCX (minimal)

## Attribution

This project is a JavaScript **port/derivative** of **Open-Xml-PowerTools** (C#), including familiar type and module naming (e.g., `WmlDocument`, `WmlToHtmlConverter`, `MarkupSimplifier`).

- Upstream project: https://github.com/OpenXmlDev/Open-Xml-PowerTools
- This repo contains a **JavaScript translation** with additional modifications and re-architecture (ZIP/OPC + XML implemented in JS for browser compatibility).
- Required upstream notices and the full upstream MIT license text are included in `NOTICE.md`.

This repo does **not** use the Open XML SDK; instead it re-implements the needed ZIP/OPC and XML manipulation in JavaScript to remain dependency-free and browser-compatible.

### Trademarks / endorsement

“Microsoft” and related marks are trademarks of their respective owners. Use of names is for identification only and does not imply endorsement.

### Dependencies & sample assets

This repository does not intentionally ship upstream sample DOCX documents or other external assets. The test fixtures in `test/fixtures/` are minimal generated files intended to be MIT-compatible.

## License

- This repository is licensed under MIT: see `LICENSE`.
- Portions derived from Open-Xml-PowerTools remain under the upstream MIT license and required notices: see `NOTICE.md`.
