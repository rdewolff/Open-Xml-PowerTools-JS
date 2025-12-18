import { readdir, readFile, stat } from "node:fs/promises";
import { join, resolve } from "node:path";
import process from "node:process";
import { OpenXmlPowerToolsDocument, WmlDocument, WmlToHtmlConverter } from "../src/index.js";

function parseArgs(argv) {
  const args = { dir: resolve(process.env.HOME ?? ".", "Downloads"), maxDepth: 2, includeTemp: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--dir") args.dir = resolve(argv[++i]);
    else if (a === "--maxDepth") args.maxDepth = Number.parseInt(argv[++i], 10);
    else if (a === "--includeTemp") args.includeTemp = true;
    else if (a === "--help" || a === "-h") args.help = true;
  }
  return args;
}

function isTempOfficeFile(name) {
  return name.startsWith("~$");
}

async function* walk(dir, depth, maxDepth) {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const entry of entries) {
    const p = join(dir, entry.name);
    if (entry.isDirectory()) {
      if (depth < maxDepth) yield* walk(p, depth + 1, maxDepth);
      continue;
    }
    if (!entry.isFile()) continue;
    const lower = entry.name.toLowerCase();
    if (lower.endsWith(".docx") || lower.endsWith(".doc")) yield p;
  }
}

function formatMs(ms) {
  return `${ms.toFixed(1)}ms`;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log("Usage: node tools/probe-word-files.mjs [--dir <path>] [--maxDepth <n>] [--includeTemp]");
    process.exit(0);
  }

  const dirExists = await stat(args.dir)
    .then((s) => s.isDirectory())
    .catch(() => false);
  if (!dirExists) {
    console.error(`Directory not found: ${args.dir}`);
    process.exit(2);
  }

  const files = [];
  for await (const p of walk(args.dir, 0, args.maxDepth)) files.push(p);
  files.sort((a, b) => a.localeCompare(b));

  console.log(`Node: ${process.version}`);
  console.log(`Scanning: ${args.dir} (maxDepth=${args.maxDepth})`);
  console.log(`Found: ${files.length} (*.docx|*.doc)`);

  const results = [];
  for (const path of files) {
    const name = path.split("/").pop() ?? path;
    const lower = name.toLowerCase();
    const ext = lower.endsWith(".docx") ? "docx" : lower.endsWith(".doc") ? "doc" : "other";

    if (!args.includeTemp && isTempOfficeFile(name)) {
      results.push({ path, ext, status: "skipped", reason: "temp-office-file" });
      continue;
    }

    const start = performance.now();
    try {
      const bytes = new Uint8Array(await readFile(path));
      if (ext === "doc") {
        results.push({ path, ext, status: "skipped", reason: "legacy-doc-not-supported", bytes: bytes.length });
        continue;
      }

      const doc = OpenXmlPowerToolsDocument.fromBytes(bytes, { fileName: name });
      const type = await doc.detectType().catch(() => "unknown");
      if (type !== "docx") {
        results.push({ path, ext, status: "skipped", reason: `not-docx (${type})`, bytes: bytes.length });
        continue;
      }

      const wml = WmlDocument.fromBytes(bytes, { fileName: name });
      const res = await WmlToHtmlConverter.convertToHtml(wml, { additionalCss: "" });
      results.push({
        path,
        ext,
        status: "ok",
        bytes: bytes.length,
        htmlBytes: res.html.length,
        warnings: res.warnings.length,
        timeMs: performance.now() - start,
      });
    } catch (e) {
      results.push({
        path,
        ext,
        status: "error",
        timeMs: performance.now() - start,
        code: e?.code ?? null,
        message: e?.message ?? String(e),
      });
    }
  }

  const ok = results.filter((r) => r.status === "ok");
  const skipped = results.filter((r) => r.status === "skipped");
  const errors = results.filter((r) => r.status === "error");

  console.log("");
  console.log(`OK: ${ok.length}  Skipped: ${skipped.length}  Errors: ${errors.length}`);

  if (errors.length) {
    console.log("");
    console.log("Errors:");
    for (const e of errors) {
      const ms = e.timeMs != null ? formatMs(e.timeMs) : "";
      console.log(`- ${e.path} ${ms} ${e.code ? `[${e.code}]` : ""} ${e.message}`);
    }
  }

  // Machine-readable output for copy/paste into issues.
  console.log("");
  console.log("JSON:");
  console.log(JSON.stringify({ dir: args.dir, maxDepth: args.maxDepth, results }, null, 2));
}

await main();

