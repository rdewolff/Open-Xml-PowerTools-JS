Fixtures are stored in `test/fixtures/`.

## Bundled (small) fixtures

Bundled fixtures are stored as text (`.base64`) so the repository remains mostly source-only.

- `minimal.docx.base64` is a minimal valid DOCX ZIP/OPC package with one paragraph containing `Hello OpenXmlPowerTools-JS`.

To regenerate (requires `zip` + `base64` tools):

1. Create the fixture folder structure with XML parts.
2. Run `zip -X -r -9 minimal.docx .`
3. Run `base64 minimal.docx > minimal.docx.base64`

## Upstream fixtures (Open-Xml-PowerTools)

`test/fixtures/open-xml-powertools/` contains DOCX fixtures copied from the upstream **Open-Xml-PowerTools** repository, used for broad compatibility smoke testing.

- Source: `/Users/rdewolff/Projects/Open-Xml-PowerTools`
- License/attribution: covered by the upstream MIT license and notices (see `NOTICE.md` at repo root).

## Local (untracked) fixtures

Put proprietary/large local docs under `test/fixtures/local/` (this folder is gitignored), e.g.:

- `test/fixtures/local/Whisperit v3 changes.docx`

The smoke test `test/wml-to-html-fixtures.test.js` will automatically pick them up if present.
