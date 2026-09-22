import PDFDocument from "pdfkit";
import bwip from "bwip-js";
import { loadWorkbook, normHyphen } from "@/lib/solidSku";
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
  mixed: boolean;
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

/** Serial numbers in the sample: 12471600000050, ...51, ...62 = base + row index */
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
      mixed: !size && !ean,
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

export async function buildLabelsPdf(
  rows: LabelRow[],
  serialBase: number,
): Promise<Buffer> {
  const renderSerial = cachedRenderer("code128", 200, 18);
  const renderPo = cachedRenderer("code128", 190, 16);
  const renderEan = cachedRenderer("ean13", 255, 15);

  // pre-render everything so the doc build stays synchronous
  const serials = await Promise.all(
    rows.map((_, i) => renderSerial(cartonSerial(serialBase, i))),
  );
  const pos = await Promise.all(rows.map((r) => renderPo(r.po)));
  const eans = await Promise.all(
    rows.map((r) => (r.mixed || !r.ean ? Promise.resolve(null) : renderEan(r.ean))),
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

  rows.forEach((r, i) => {
    doc.addPage({ size: [PAGE_W, PAGE_H], margin: 0 });
    doc.lineWidth(1).rect(2, 3, PAGE_W - 4, PAGE_H - 9).stroke();

    const serial = cartonSerial(serialBase, i);
    const sb = serials[i];
    if (sb) {
      const sx = (PAGE_W - sb.widthPt) / 2;
      doc.image(sb.buffer, sx, 18, { width: sb.widthPt });
      barcodeCaption(doc, sb, sx, 18, serial, 11, 1.6);
    }

    boxLabel(doc, 30, 119, 300, 53, [
      { text: `Order no:${r.po}`, size: LABEL_FONT_SIZE },
    ]);
    const pb = pos[i];
    if (pb) {
      doc.image(pb.buffer, 353, 112, { width: pb.widthPt });
      barcodeCaption(doc, pb, 353, 112, r.po, 10, 0.6);
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
      barcodeCaption(doc, eb, 62, 300, eanDigits, 9, 1.4);
    }

    // start 10pt below the QTY box's bottom border (277)
    boxLabel(doc, 450, 287, 130, 116, [
      { text: "CARTON BOX", size: LABEL_FONT_SIZE },
      { text: "NUMBER:", size: LABEL_FONT_SIZE },
      { text: r.boxLabel, size: LABEL_FONT_SIZE },
    ]);
  });

  doc.end();
  return done;
}
