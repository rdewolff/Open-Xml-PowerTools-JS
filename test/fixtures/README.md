Fixtures are stored as text (`.base64`) so the repository remains source-only.

- `minimal.docx.base64` is a minimal valid DOCX ZIP/OPC package with one paragraph containing `Hello OpenXmlPowerTools-JS`.

To regenerate (requires `zip` + `base64` tools):

1. Create the fixture folder structure with XML parts.
2. Run `zip -X -r -9 minimal.docx .`
3. Run `base64 minimal.docx > minimal.docx.base64`

