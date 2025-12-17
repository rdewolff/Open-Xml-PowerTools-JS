import { parseXml } from "./internal/xml.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const WmlToHtmlConverter = {
  async convertToHtml(doc, settings = {}) {
    const s = normalizeSettings(settings);
    const mainXmlText = await doc.getPartText("/word/document.xml");
    const xmlDoc = parseXml(mainXmlText);

    const warnings = [];
    const it = xmlDoc.root.descendantsByNameNS(W_NS, "body");
    const body = it.next().value ?? findFirstByLocal(xmlDoc.root, "body");
    const paragraphs = body ? body.children.filter((c) => c?.qname && isW(c, "p")) : [];

    const htmlParts = [];
    for (const p of paragraphs) htmlParts.push(renderParagraph(p, warnings));

    const cssText = [s.generalCss, s.additionalCss].filter(Boolean).join("\n");
    const html = renderHtmlDocument({
      title: s.pageTitle,
      cssText,
      bodyHtml: htmlParts.join("\n"),
    });

    return { html, cssText, warnings };
  },
};

export const HtmlConverter = {
  convertToHtml: WmlToHtmlConverter.convertToHtml,
};

function normalizeSettings(settings) {
  return {
    pageTitle: settings.pageTitle ?? "",
    generalCss: settings.generalCss ?? "span { white-space: pre-wrap; }",
    additionalCss: settings.additionalCss ?? "",
  };
}

function renderHtmlDocument({ title, cssText, bodyHtml }) {
  const safeTitle = escapeHtml(title);
  const styleBlock = cssText ? `<style>${escapeStyle(cssText)}</style>` : "";
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${safeTitle}</title>
${styleBlock}
</head>
<body>
${bodyHtml}
</body>
</html>`;
}

function renderParagraph(p, warnings) {
  const inner = [];
  for (const child of p.children ?? []) {
    if (!child?.qname) continue;
    if (isW(child, "r")) {
      inner.push(renderRun(child, warnings));
      continue;
    }
    if (isW(child, "tbl")) {
      warnings.push({ code: "OXPT_HTML_UNSUPPORTED_TABLE", message: "Tables not yet supported", part: "/word/document.xml" });
      continue;
    }
  }
  return `<p>${inner.join("")}</p>`;
}

function renderRun(r, warnings) {
  const rPr = (r.children ?? []).find((c) => c?.qname && isW(c, "rPr")) ?? null;
  const bold = !!(rPr && (rPr.children ?? []).some((c) => c?.qname && isW(c, "b")));
  const italic = !!(rPr && (rPr.children ?? []).some((c) => c?.qname && isW(c, "i")));
  const underline = !!(rPr && (rPr.children ?? []).some((c) => c?.qname && isW(c, "u")));

  const pieces = [];
  for (const child of r.children ?? []) {
    if (!child?.qname) continue;
    if (isW(child, "t")) pieces.push(escapeHtml(child.textContent()));
    else if (isW(child, "tab")) pieces.push("    ");
    else if (isW(child, "br")) pieces.push("<br/>");
    else if (isW(child, "delText")) {
      // should be removed by RevisionAccepter; ignore for now
      continue;
    } else {
      warnings.push({
        code: "OXPT_HTML_UNSUPPORTED_RUN_CHILD",
        message: `Unsupported run child: ${child.qname}`,
        part: "/word/document.xml",
      });
    }
  }

  let html = pieces.join("");
  if (underline) html = `<u>${html}</u>`;
  if (italic) html = `<em>${html}</em>`;
  if (bold) html = `<strong>${html}</strong>`;
  return html ? `<span>${html}</span>` : "";
}

function isW(el, localName) {
  const { prefix, local } = el.nameParts();
  if (local !== localName) return false;
  if (prefix === "w") return true;
  return el.lookupNamespaceUri(prefix) === W_NS;
}

function findFirstByLocal(root, localName) {
  for (const d of root.descendants()) {
    if (d.nameParts().local === localName) return d;
  }
  return null;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeStyle(text) {
  // Style tag content doesn't need HTML entity escaping beyond closing tag safety.
  return String(text).replaceAll("</style", "<\\/style");
}
