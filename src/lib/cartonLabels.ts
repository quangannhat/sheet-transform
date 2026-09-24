import PDFDocument from "pdfkit";
import bwip from "bwip-js";
import {
  KNOWN_SIZES,
  loadWorkbook,
  normHyphen,
  sizeRank,
} from "@/lib/solidSku";
import { ARIAL_BOLD_TTF_BASE64 } from "@/lib/fonts/arialBold";
import { ARIAL_REGULAR_TTF_BASE64 } from "@/lib/fonts/arialRegular";

export type LabelRow = {
  lpn: string;
  po: string;
  art: string;
  col: string;
  size: string;
  qty: string;
  ean: string;
  boxLabel: string;
};

export const DEFAULT_SERIAL_BASE = "12471600000049";

export function resolveSerialBase(raw: string): {
  base: number;
  valid: boolean;
} {
  const trimmed = raw.trim();
  if (!trimmed) return { base: Number(DEFAULT_SERIAL_BASE), valid: true };
  if (!/^\d+$/.test(trimmed)) return { base: 0, valid: false };
  return { base: Number(trimmed), valid: true };
}

export function cartonLabelFileName(sourceName: string): string {
  const base = sourceName.replace(/\.xl[sm]?$/i, "") || "SOLID_SKU";
  return `${base} labels.pdf`;
}

/** Serial numbers in the sample: 12471600000050, ...51, ...62 = base + solid-carton index (mixed cartons take no LPN sticker serial) */
export function cartonSerial(serialBase: number, rowIndex: number): string {
  return String(serialBase + rowIndex + 1);
}

function cellText(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number" || typeof v === "string") return String(v);
  if (typeof v === "boolean") return v ? "TRUE" : "";
  const o = v as Record<string, unknown>;
  if ("result" in o) return cellText(o.result);
  if ("text" in o) return String(o.text);
  if ("richText" in o) {
    return ((o.richText ?? []) as Array<{ text?: string }>)
      .map((p) => p.text ?? "")
      .join("");
  }
  return "";
}

/** Read rows back from a generated SOLID SKU workbook. */
export async function parseSolidSkuRows(fileBuffer: Buffer): Promise<LabelRow[]> {
  const wb = await loadWorkbook(fileBuffer);
  const ws =
    wb.worksheets.find((w) => w.name === "SOLID SKU") ?? wb.worksheets[0];
  if (!ws) throw new Error("Workbook has no worksheets");

  const header = ws.getRow(1);
  if (cellText(header.getCell(1).value).trim().toUpperCase() !== "LPN") {
    throw new Error(
      "This file is not a generated SOLID SKU workbook (missing LPN header)",
    );
  }

  const rows: LabelRow[] = [];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber < 3) return;
    const art = cellText(row.getCell(4).value).trim();
    if (!art) return;
    const size = cellText(row.getCell(6).value).trim();
    const ean = cellText(row.getCell(8).value).trim();
    rows.push({
      lpn: cellText(row.getCell(1).value).trim(),
      po: cellText(row.getCell(2).value).trim(),
      art,
      col: normHyphen(cellText(row.getCell(5).value).trim()),
      size,
      qty: cellText(row.getCell(7).value).trim(),
      ean,
      boxLabel: cellText(row.getCell(9).value).trim(),
    });
  });

  if (rows.length === 0) throw new Error("No label rows found in the workbook");
  return rows;
}

const PT_PER_MM = 72 / 25.4;
const PAGE_W = 595.35;
const PAGE_H = 420.9;

type RenderedBarcode = {
  buffer: Buffer;
  widthPt: number;
  heightPt: number;
};

/**
 * Bars only (includetext off): this bwip-js build ignores textmargin and
 * paints EAN guard bars through the digits, so captions are typeset with
 * pdfkit directly under the image instead.
 */
async function renderBarcode(
  bcid: string,
  text: string,
  targetWidthPt: number,
  heightMm: number,
): Promise<RenderedBarcode> {
  const widthMm = targetWidthPt / PT_PER_MM;
  const buffer = await bwip.toBuffer({
    bcid,
    text,
    padding: 0,
    width: widthMm,
    height: heightMm,
  });
  const pngW = buffer.readUInt32BE(16);
  const pngH = buffer.readUInt32BE(20);
  return {
    buffer,
    widthPt: targetWidthPt,
    heightPt: (targetWidthPt * pngH) / pngW,
  };
}

function cachedRenderer(
  bcid: string,
  widthPt: number,
  heightMm: number,
): (text: string) => Promise<RenderedBarcode | null> {
  const cache = new Map<string, Promise<RenderedBarcode | null>>();
  return (text: string) => {
    let hit = cache.get(text);
    if (!hit) {
      hit = renderBarcode(bcid, text, widthPt, heightMm).catch(() => null);
      cache.set(text, hit);
    }
    return hit;
  };
}

function barcodeCaption(
  doc: PDFKit.PDFDocument,
  rb: RenderedBarcode,
  x: number,
  y: number,
  digits: string,
  size: number,
  characterSpacing = 0,
) {
  doc
    .font(ARIAL_REGULAR_FONT)
    .fontSize(size)
    .fillColor("#000000")
    .text(digits, x, y + rb.heightPt + 1.5, {
      width: rb.widthPt,
      align: "center",
      characterSpacing,
      lineBreak: false,
    });
}

type CenteredLine = { text: string; size: number };

/** Arial (TTF embedded, see fonts/arialBold.ts) for all label text. */
export const LABEL_FONT = "Arial-Bold";
const ARIAL_REGULAR_FONT = "Arial-Regular";
export const LABEL_FONT_SIZE = 26;

/** Split lines that are too wide for the box into word-wrapped rows. */
function wrapLines(
  doc: PDFKit.PDFDocument,
  lines: CenteredLine[],
  w: number,
): CenteredLine[] {
  const out: CenteredLine[] = [];
  for (const l of lines) {
    if (!l.text) continue;
    doc.font(LABEL_FONT).fontSize(l.size);
    if (doc.widthOfString(l.text) <= w - 6) {
      out.push(l);
      continue;
    }
    let cur = "";
    for (const word of l.text.split(/\s+/)) {
      const cand = cur ? `${cur} ${word}` : word;
      if (!cur || doc.widthOfString(cand) <= w - 6) cur = cand;
      else {
        out.push({ text: cur, size: l.size });
        cur = word;
      }
    }
    if (cur) out.push({ text: cur, size: l.size });
  }
  return out;
}

function boxLabel(
  doc: PDFKit.PDFDocument,
  x: number,
  y: number,
  w: number,
  h: number,
  lines: CenteredLine[],
) {
  doc.lineWidth(0.8).rect(x, y, w, h).stroke();
  const rows = wrapLines(doc, lines, w);
  if (rows.length === 0) return;
  const maxLh = (h - 4) / rows.length;
  const lineHeight = (s: number) => Math.min(s * 1.35, maxLh);
  const total = rows.reduce((a, l) => a + lineHeight(l.size), 0);
  let cy = y + (h - total) / 2;
  doc.fillColor("#000000");
  for (const l of rows) {
    doc
      .font(LABEL_FONT)
      .fontSize(l.size)
      .text(l.text, x, cy, { width: w, align: "center", lineBreak: false });
    cy += lineHeight(l.size);
  }
}

/** A carton is one page; >1 rows (polybags) makes it a mixed carton. */
type CartonGroup = {
  boxLabel: string;
  rows: LabelRow[];
  mixed: boolean;
  /** index of the carton's first row in the sheet (keeps LPN/serial aligned) */
  firstRowIndex: number;
};

function groupCartons(rows: LabelRow[]): CartonGroup[] {
  const out: CartonGroup[] = [];
  rows.forEach((r, idx) => {
    const last = out[out.length - 1];
    if (last && last.boxLabel === r.boxLabel) last.rows.push(r);
    else
      out.push({
        boxLabel: r.boxLabel,
        rows: [r],
        mixed: false,
        firstRowIndex: idx,
      });
  });
  for (const g of out) g.mixed = g.rows.length > 1;
  return out;
}

/**
 * Shrink text from LABEL_FONT_SIZE until it word-wraps into the cell.
 * widthOfString only honours the size set via doc.fontSize().
 */
function fitCellText(
  doc: PDFKit.PDFDocument,
  text: string,
  w: number,
  h: number,
  startSize: number,
): { size: number; lines: string[] } | null {
  if (!text) return null;
  const usableW = w - 4;
  const usableH = h - 4;
  let size = startSize;
  for (;;) {
    doc.fontSize(size);
    const lines: string[] = [];
    let cur = "";
    let overflow = false;
    for (const word of text.split(/\s+/)) {
      if (doc.widthOfString(word) > usableW) {
        overflow = true;
        break;
      }
      const cand = cur ? `${cur} ${word}` : word;
      if (doc.widthOfString(cand) <= usableW) cur = cand;
      else {
        lines.push(cur);
        cur = word;
      }
    }
    if (cur) lines.push(cur);
    if (!overflow && lines.length * size * 1.25 <= usableH) {
      return { size, lines };
    }
    if (size <= 6) return overflow ? null : { size, lines };
    size = Math.max(6, size - 2);
  }
}

function tableCell(
  doc: PDFKit.PDFDocument,
  x: number,
  y: number,
  w: number,
  h: number,
  text: string,
  startSize = LABEL_FONT_SIZE,
) {
  doc.lineWidth(0.8).rect(x, y, w, h).stroke();
  doc.font(LABEL_FONT);
  const lay = fitCellText(doc, text, w, h, startSize);
  if (!lay) return;
  const lh = lay.size * 1.25;
  let cy = y + (h - lay.lines.length * lh) / 2;
  doc.fillColor("#000000");
  for (const ln of lay.lines) {
    doc
      .font(LABEL_FONT)
      .fontSize(lay.size)
      .text(ln, x, cy, { width: w, align: "center", lineBreak: false });
    cy += lh;
  }
}

/** Base size for table cells that are not one of the Polybags columns. */
const TABLE_FONT_SIZE = 14;

/**
 * "Mixed SKU carton version 2" sticker: Order no + PO barcode + BOX CARTON
 * NUMBER header, then one row per (article, color) and a fixed column per
 * size (XXS..XXXL, empty when absent) with polybag counts. Intentionally
 * carries no LPN (LPN goes on the polybags, i.e. the per-size rows of the
 * SOLID SKU sheet).
 */
function drawMixedSticker(
  doc: PDFKit.PDFDocument,
  g: CartonGroup,
  poBarcode: RenderedBarcode | null,
) {
  const r0 = g.rows[0];
  doc.lineWidth(1).rect(2, 3, PAGE_W - 4, PAGE_H - 9).stroke();

  boxLabel(doc, 24, 28, 236, 76, [
    { text: `Order no:${r0.po}`, size: LABEL_FONT_SIZE },
  ]);
  if (poBarcode) {
    doc.image(poBarcode.buffer, 270, 34, { width: poBarcode.widthPt });
    barcodeCaption(doc, poBarcode, 270, 34, r0.po, 12, 0.8);
  }
  boxLabel(doc, 464, 28, 107, 76, [
    { text: "BOX CARTON", size: 14 },
    { text: "NUMBER:", size: 14 },
    { text: r0.boxLabel, size: 22 },
  ]);

  type ArticleRow = {
    art: string;
    col: string;
    qtys: Map<string, string>;
    polybags: number;
  };
  const byArt = new Map<string, ArticleRow>();
  const present = new Set<string>();
  for (const r of g.rows) {
    present.add(r.size);
    const key = `${r.art}\u0000${r.col}`;
    let a = byArt.get(key);
    if (!a) {
      a = { art: r.art, col: r.col, qtys: new Map(), polybags: 0 };
      byArt.set(key, a);
    }
    if (!a.qtys.has(r.size)) {
      a.qtys.set(r.size, r.qty);
      a.polybags += 1;
    }
  }
  // fixed size columns in PL order, plus any unexpected sizes after them
  const sizes = [
    ...KNOWN_SIZES,
    ...[...present]
      .filter((s) => !KNOWN_SIZES.includes(s.toUpperCase()))
      .sort((a, b) => sizeRank(a) - sizeRank(b) || a.localeCompare(b)),
  ];
  const articles = [...byArt.values()];

  const x0 = 24;
  const wArt = 70;
  const wCol = 76;
  const wPoly = 66;
  const wTot = 78;
  const wSize =
    (PAGE_W - 2 * x0 - 4 - wArt - wCol - wPoly - wTot) / sizes.length;
  // taller size header, shorter Article/Color header, data rows, then two
  // blank rows for hand-written articles (per the manual's sticker photo)
  const EMPTY_ROWS = 2;
  const yTop = 116;
  const yBot = PAGE_H - 16;
  const unit =
    (yBot - yTop) / (1.35 + 0.65 + articles.length + EMPTY_ROWS);
  const headH = 1.35 * unit;
  const subH = 0.65 * unit;

  let x = x0;
  tableCell(doc, x, yTop, wArt, headH, "", TABLE_FONT_SIZE);
  x += wArt;
  tableCell(doc, x, yTop, wCol, headH, "Size:", TABLE_FONT_SIZE);
  x += wCol;
  for (const s of sizes) {
    tableCell(doc, x, yTop, wSize, headH, s, TABLE_FONT_SIZE);
    x += wSize;
  }
  tableCell(doc, x, yTop, wPoly, headH, "Polybags");
  x += wPoly;
  const totalX = x;
  tableCell(doc, totalX, yTop, wTot, headH, "Total Polybags");

  const y1 = yTop + headH;
  x = x0 + wArt + wCol;
  tableCell(doc, x0, y1, wArt, subH, "Article no:", TABLE_FONT_SIZE);
  tableCell(doc, x0 + wArt, y1, wCol, subH, "Color no:", TABLE_FONT_SIZE);
  for (let i = 0; i < sizes.length; i += 1) {
    tableCell(doc, x, y1, wSize, subH, "", TABLE_FONT_SIZE);
    x += wSize;
  }
  tableCell(doc, x, y1, wPoly, subH, "", TABLE_FONT_SIZE);
  const totalPolybags = articles.reduce((a, b) => a + b.polybags, 0);
  tableCell(doc, totalX, y1, wTot, yBot - y1, String(totalPolybags));

  const dataRows = articles.length + EMPTY_ROWS;
  const y2 = y1 + subH;
  for (let i = 0; i < dataRows; i += 1) {
    const a = i < articles.length ? articles[i] : null;
    const y = y2 + i * unit;
    x = x0 + wArt + wCol;
    tableCell(doc, x0, y, wArt, unit, a ? a.art : "", TABLE_FONT_SIZE);
    tableCell(doc, x0 + wArt, y, wCol, unit, a ? a.col : "", TABLE_FONT_SIZE);
    for (const s of sizes) {
      tableCell(
        doc,
        x,
        y,
        wSize,
        unit,
        a ? (a.qtys.get(s) ?? "") : "",
        TABLE_FONT_SIZE,
      );
      x += wSize;
    }
    tableCell(
      doc,
      x,
      y,
      wPoly,
      unit,
      a ? String(a.polybags) : "",
      TABLE_FONT_SIZE,
    );
  }
}

export async function buildLabelsPdf(
  rows: LabelRow[],
  serialBase: number,
): Promise<Buffer> {
  const cartons = groupCartons(rows);
  const renderSerial = cachedRenderer("code128", 200, 18);
  const renderPo = cachedRenderer("code128", 190, 16);
  const renderEan = cachedRenderer("ean13", 255, 15);

  // serials follow the carton's first sheet row, so a solid carton's sticker
  // serial always carries the same index as its LPN in the workbook; mixed
  // cartons take no LPN sticker at all
  const serialTexts = cartons.map((g) =>
    g.mixed ? null : cartonSerial(serialBase, g.firstRowIndex),
  );

  // pre-render everything so the doc build stays synchronous
  const serials = await Promise.all(
    cartons.map((_, i) =>
      serialTexts[i] === null
        ? Promise.resolve(null)
        : renderSerial(serialTexts[i] as string),
    ),
  );
  const pos = await Promise.all(cartons.map((g) => renderPo(g.rows[0].po)));
  const eans = await Promise.all(
    cartons.map((g) =>
      g.mixed || !g.rows[0].ean
        ? Promise.resolve(null)
        : renderEan(g.rows[0].ean),
    ),
  );

  const doc = new PDFDocument({
    size: [PAGE_W, PAGE_H],
    margin: 0,
    autoFirstPage: false,
  });
  doc.registerFont(
    LABEL_FONT,
    Buffer.from(ARIAL_BOLD_TTF_BASE64, "base64"),
  );
  doc.registerFont(
    ARIAL_REGULAR_FONT,
    Buffer.from(ARIAL_REGULAR_TTF_BASE64, "base64"),
  );
  const chunks: Buffer[] = [];
  doc.on("data", (c: Buffer) => chunks.push(c));
  const done = new Promise<Buffer>((resolve, reject) => {
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  cartons.forEach((g, i) => {
    doc.addPage({ size: [PAGE_W, PAGE_H], margin: 0 });
    if (g.mixed) {
      drawMixedSticker(doc, g, pos[i]);
      return;
    }
    const r = g.rows[0];
    doc.lineWidth(1).rect(2, 3, PAGE_W - 4, PAGE_H - 9).stroke();

    const serial = serialTexts[i];
    const sb = serials[i];
    if (sb && serial !== null) {
      const sx = (PAGE_W - sb.widthPt) / 2;
      doc.image(sb.buffer, sx, 18, { width: sb.widthPt });
      barcodeCaption(doc, sb, sx, 18, serial, 13, 1.8);
    }

    boxLabel(doc, 30, 119, 300, 53, [
      { text: `Order no:${r.po}`, size: LABEL_FONT_SIZE },
    ]);
    const pb = pos[i];
    if (pb) {
      doc.image(pb.buffer, 353, 112, { width: pb.widthPt });
      barcodeCaption(doc, pb, 353, 112, r.po, 12, 0.8);
    }

    boxLabel(doc, 30, 211, 186, 66, [
      { text: "Article no:", size: LABEL_FONT_SIZE },
      { text: r.art, size: LABEL_FONT_SIZE },
    ]);
    boxLabel(doc, 224, 211, 129, 66, [
      { text: "Color no:", size: LABEL_FONT_SIZE },
      { text: r.col, size: LABEL_FONT_SIZE },
    ]);
    boxLabel(doc, 364, 211, 76, 66, [
      { text: "Size:", size: LABEL_FONT_SIZE },
      { text: r.size, size: LABEL_FONT_SIZE },
    ]);
    boxLabel(doc, 450, 211, 129, 66, [
      { text: "QTY/pcs:", size: LABEL_FONT_SIZE },
      { text: r.qty, size: LABEL_FONT_SIZE },
    ]);

    const eb = eans[i];
    if (eb) {
      doc.image(eb.buffer, 62, 300, { width: eb.widthPt });
      const eanDigits =
        r.ean.length === 14 && r.ean.startsWith("0")
          ? r.ean.slice(1)
          : r.ean;
      barcodeCaption(doc, eb, 62, 300, eanDigits, 11, 1.8);
    }

    // start 10pt below the QTY box's bottom border (277)
    boxLabel(doc, 450, 287, 130, 116, [
      { text: "CARTON BOX", size: 17 },
      { text: "NUMBER:", size: 17 },
      { text: r.boxLabel, size: 22 },
    ]);
  });

  doc.end();
  return done;
}
