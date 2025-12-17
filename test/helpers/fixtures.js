import { readdir, readFile, stat } from "node:fs/promises";
import { join, relative, resolve } from "node:path";

export const FIXTURE_ROOT = resolve(new URL("../fixtures", import.meta.url).pathname);

export async function listDocxFixturePaths(options = {}) {
  const roots = [];
  const includeBundled = options.includeBundled ?? true;
  const includeUpstream = options.includeUpstream ?? true;
  const includeLocal = options.includeLocal ?? true;

  if (includeBundled) roots.push(join(FIXTURE_ROOT));
  if (includeUpstream) roots.push(join(FIXTURE_ROOT, "open-xml-powertools"));
  if (includeLocal) roots.push(join(FIXTURE_ROOT, "local"));

  const found = [];
  for (const root of roots) {
    const exists = await pathExists(root);
    if (!exists) continue;
    // eslint-disable-next-line no-await-in-loop
    for await (const file of walkFiles(root)) {
      if (file.toLowerCase().endsWith(".docx")) found.push(file);
    }
  }

  const unique = [...new Set(found.map((p) => resolve(p)))].sort();
  return applyFilters(unique, options);
}

export async function readFixtureBytes(path) {
  const bytes = new Uint8Array(await readFile(path));
  return bytes;
}

export function fixtureId(path) {
  const rel = relative(FIXTURE_ROOT, path).replaceAll("\\", "/");
  return rel;
}

function applyFilters(paths, options) {
  let out = paths;

  const filter = options.filter ?? process.env.OXPT_FIXTURE_FILTER ?? "";
  if (filter) {
    const needle = filter.toLowerCase();
    out = out.filter((p) => p.toLowerCase().includes(needle));
  }

  const limitRaw = options.limit ?? process.env.OXPT_FIXTURE_LIMIT ?? "";
  if (limitRaw) {
    const n = Number.parseInt(String(limitRaw), 10);
    if (Number.isFinite(n) && n > 0) out = out.slice(0, n);
  }

  return out;
}

async function pathExists(path) {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

async function* walkFiles(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const p = join(dir, entry.name);
    if (entry.isDirectory()) {
      yield* walkFiles(p);
      continue;
    }
    if (entry.isFile()) yield p;
  }
}

