import PDFDocument from "pdfkit";
import bwip from "bwip-js";
import { loadWorkbook, normHyphen } from "@/lib/solidSku";

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
const BWIP_SCALE = 4; // px per mm

type RenderedBarcode = {
  buffer: Buffer;
  widthPt: number;
};

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
    includetext: true,
    textsize: 6.5,
    textmargin: 0.5,
    padding: 0,
    scale: BWIP_SCALE,
    width: widthMm,
    height: heightMm,
  });
  return { buffer, widthPt: widthMm * PT_PER_MM };
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

function boxLabel(
  doc: PDFKit.PDFDocument,
  x: number,
  y: number,
  w: number,
  h: number,
  label: string,
  value: string,
  opts: { labelSize?: number; valueSize?: number; labelLines?: number } = {},
) {
  const { labelSize = 8, valueSize = 14, labelLines = 1 } = opts;
  doc.lineWidth(0.8).rect(x, y, w, h).stroke();
  const labelH = label ? labelSize * 1.25 * labelLines + 6 : 0;
  if (label) {
    doc
      .font("Helvetica-Bold")
      .fontSize(labelSize)
      .fillColor("#000000")
      .text(label, x + 6, y + 5, { width: w - 12, lineBreak: labelLines > 1 });
  }
  if (value) {
    const valueY = y + labelH + 2;
    doc
      .font("Helvetica-Bold")
      .fontSize(valueSize)
      .fillColor("#000000")
      .text(value, x, valueY, {
        width: w,
        align: "center",
        ellipsis: false,
        lineBreak: false,
      });
  }
}

export async function buildLabelsPdf(
  rows: LabelRow[],
  serialBase: number,
): Promise<Buffer> {
  const renderSerial = cachedRenderer("code128", 200, 18);
  const renderPo = cachedRenderer("code128", 200, 17);
  const renderEan = cachedRenderer("ean13", 266, 15);

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
  const chunks: Buffer[] = [];
  doc.on("data", (c: Buffer) => chunks.push(c));
  const done = new Promise<Buffer>((resolve, reject) => {
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  rows.forEach((r, i) => {
    doc.addPage({ size: [PAGE_W, PAGE_H], margin: 0 });
    doc.lineWidth(1).rect(2, 3, PAGE_W - 4, PAGE_H - 9).stroke();

    const sb = serials[i];
    if (sb) {
      doc.image(sb.buffer, (PAGE_W - sb.widthPt) / 2, 18, { width: sb.widthPt });
    }

    boxLabel(doc, 30, 119, 300, 53, "Order no:", r.po, { valueSize: 16 });
    const pb = pos[i];
    if (pb) doc.image(pb.buffer, 364, 114, { width: pb.widthPt });

    boxLabel(doc, 30, 211, 186, 66, "Article no:", r.art, { valueSize: 15 });
    boxLabel(doc, 224, 211, 129, 66, "Color no:", r.col, { valueSize: 15 });
    boxLabel(doc, 364, 211, 76, 66, "Size:", r.size, { valueSize: 15 });
    boxLabel(doc, 450, 211, 129, 66, "QTY/pcs:", r.qty, { valueSize: 15 });

    const eb = eans[i];
    if (eb) doc.image(eb.buffer, 57, 300, { width: eb.widthPt });

    boxLabel(doc, 450, 277, 130, 126, "CARTON BOX NUMBER:", r.boxLabel, {
      labelSize: 9,
      valueSize: 15,
      labelLines: 2,
    });
  });

  doc.end();
  return done;
}
