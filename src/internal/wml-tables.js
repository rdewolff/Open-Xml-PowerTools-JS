import { XmlElement } from "./xml.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export function computeTableGrid(tbl) {
  // Returns a grid of cells with rowspan/colspan, omitting vMerge continuations.
  const rows = [];
  // Map column index -> origin cell for an active vMerge.
  const originByCol = new Map();

  const trEls = (tbl.children ?? []).filter((c) => c instanceof XmlElement && isW(c, "tr"));
  for (const tr of trEls) {
    const row = [];
    let col = 0;

    const tcEls = (tr.children ?? []).filter((c) => c instanceof XmlElement && isW(c, "tc"));
    for (const tc of tcEls) {
      const tcPr = (tc.children ?? []).find((c) => c instanceof XmlElement && isW(c, "tcPr")) ?? null;
      const gridSpanEl = tcPr ? firstChildW(tcPr, "gridSpan") : null;
      const gridSpan = toInt(gridSpanEl?.attributes.get("w:val") ?? gridSpanEl?.attributes.get("val")) ?? 1;

      const vMergeEl = tcPr ? firstChildW(tcPr, "vMerge") : null;
      const vMergeVal = vMergeEl ? (vMergeEl.attributes.get("w:val") ?? vMergeEl.attributes.get("val") ?? null) : null;
      const isVMergeContinue = vMergeEl && (vMergeVal == null || String(vMergeVal) === "continue");
      const isVMergeRestart = vMergeEl && String(vMergeVal) === "restart";

      if (isVMergeContinue) {
        // Extend the origin cell's rowspan, and do not emit a cell for this row.
        const origin = originByCol.get(col);
        if (origin) origin.rowspan = (origin.rowspan ?? 1) + 1;
        col += gridSpan;
        continue;
      }

      const cell = { tc, col, colspan: gridSpan, rowspan: 1 };
      row.push(cell);

      // This cell replaces any previous merge origin for covered columns.
      for (let k = 0; k < gridSpan; k++) originByCol.delete(col + k);

      if (isVMergeRestart) {
        for (let k = 0; k < gridSpan; k++) originByCol.set(col + k, cell);
      }

      col += gridSpan;
    }

    rows.push({ tr, cells: row });
  }

  return rows;
}

export function tableBordersToCss(tblPr) {
  const bordersEl = tblPr ? firstChildW(tblPr, "tblBorders") : null;
  if (!bordersEl) return null;
  const sides = ["top", "left", "bottom", "right", "insideH", "insideV"];
  const css = {};
  for (const side of sides) {
    const el = firstChildW(bordersEl, side);
    if (!el) continue;
    const val = el.attributes.get("w:val") ?? el.attributes.get("val") ?? "single";
    if (String(val) === "nil" || String(val) === "none") continue;
    const sz = toInt(el.attributes.get("w:sz") ?? el.attributes.get("sz")) ?? 4; // eighths of a point
    const color = el.attributes.get("w:color") ?? el.attributes.get("color") ?? "000000";
    const widthPt = sz / 8;
    const border = `${widthPt}pt solid #${color}`;
    css[side] = border;
  }
  return css;
}

export function cellBordersToCss(tcPr) {
  const bordersEl = tcPr ? firstChildW(tcPr, "tcBorders") : null;
  if (!bordersEl) return null;
  const sides = ["top", "left", "bottom", "right"];
  const css = {};
  for (const side of sides) {
    const el = firstChildW(bordersEl, side);
    if (!el) continue;
    const val = el.attributes.get("w:val") ?? el.attributes.get("val") ?? "single";
    if (String(val) === "nil" || String(val) === "none") continue;
    const sz = toInt(el.attributes.get("w:sz") ?? el.attributes.get("sz")) ?? 4;
    const color = el.attributes.get("w:color") ?? el.attributes.get("color") ?? "000000";
    const widthPt = sz / 8;
    css[side] = `${widthPt}pt solid #${color}`;
  }
  return css;
}

export function cellShadingToCss(tcPr) {
  const shd = tcPr ? firstChildW(tcPr, "shd") : null;
  if (!shd) return null;
  const fill = shd.attributes.get("w:fill") ?? shd.attributes.get("fill") ?? null;
  if (!fill || String(fill) === "auto") return null;
  return `#${String(fill)}`;
}

export function cellVAlignToCss(tcPr) {
  const vAlign = tcPr ? firstChildW(tcPr, "vAlign") : null;
  if (!vAlign) return null;
  const val = vAlign.attributes.get("w:val") ?? vAlign.attributes.get("val") ?? null;
  if (!val) return null;
  const v = String(val);
  if (v === "top") return "top";
  if (v === "center") return "middle";
  if (v === "bottom") return "bottom";
  return null;
}

export function cellPaddingToCss(tcPr) {
  const tcMar = tcPr ? firstChildW(tcPr, "tcMar") : null;
  if (!tcMar) return null;
  const map = { top: "padding-top", left: "padding-left", bottom: "padding-bottom", right: "padding-right" };
  const rules = [];
  for (const [side, cssName] of Object.entries(map)) {
    const el = firstChildW(tcMar, side);
    if (!el) continue;
    const w = el.attributes.get("w:w") ?? el.attributes.get("w") ?? null;
    const type = el.attributes.get("w:type") ?? el.attributes.get("type") ?? "dxa";
    if (!w) continue;
    const n = Number(w);
    if (!Number.isFinite(n)) continue;
    if (String(type) === "dxa") rules.push(`${cssName}:${n / 20}pt`);
  }
  return rules.length ? rules.join(";") : null;
}

export function cellWidthToCss(tcPr) {
  const tcW = tcPr ? firstChildW(tcPr, "tcW") : null;
  if (!tcW) return null;
  const type = tcW.attributes.get("w:type") ?? tcW.attributes.get("type") ?? "dxa";
  const w = tcW.attributes.get("w:w") ?? tcW.attributes.get("w") ?? null;
  if (!w) return null;
  const n = Number(w);
  if (!Number.isFinite(n)) return null;
  if (String(type) === "pct") {
    // stored as 50ths of a percent
    return `${n / 50}%`;
  }
  // dxa (twips)
  return `${n / 20}pt`;
}

export function isW(el, localName) {
  const { prefix, local } = el.nameParts();
  if (local !== localName) return false;
  if (prefix === "w") return true;
  return el.lookupNamespaceUri(prefix) === W_NS;
}

function firstChildW(el, local) {
  return (el.children ?? []).find((c) => c instanceof XmlElement && isW(c, local)) ?? null;
}

function toInt(v) {
  if (v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
