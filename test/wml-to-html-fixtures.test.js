import test from "node:test";
import assert from "node:assert/strict";
import { WmlDocument, WmlToHtmlConverter } from "../src/index.js";
import { fixtureId, listDocxFixturePaths, readFixtureBytes } from "./helpers/fixtures.js";

test("DOCX fixtures: convert to HTML (smoke)", async () => {
  const fixtures = await listDocxFixturePaths();
  assert.ok(fixtures.length > 0, "No DOCX fixtures found");

  const failures = [];
  for (const path of fixtures) {
    try {
      const bytes = await readFixtureBytes(path);
      const doc = WmlDocument.fromBytes(bytes, { fileName: fixtureId(path) });
      const res = await WmlToHtmlConverter.convertToHtml(doc, { additionalCss: "" });

      assert.equal(typeof res.html, "string");
      assert.ok(res.html.includes("<html"), "Expected HTML output to contain <html>");
      assert.ok(Array.isArray(res.warnings));
    } catch (e) {
      failures.push({
        fixture: fixtureId(path),
        code: e?.code ?? null,
        message: e?.message ?? String(e),
      });
      if (failures.length >= 20) break;
    }
  }

  assert.deepEqual(failures, []);
});
