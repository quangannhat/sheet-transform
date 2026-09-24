import ExcelJS from "exceljs";

export type PlRecord = {
  art: string;
  artBare: string;
  col: string;
  colRaw: string;
  size: string;
  qty: number;
  box: number;
};

export type GenerateResult = {
  buffer: Buffer;
  po: string;
  totalCartons: number | null;
  rowCount: number;
  mixedCount: number;
};

type Cell = number | string | boolean | null;

export function normHyphen(s: string): string {
  return s.replace(/[\u2010-\u2015\u2212\uff0d]/g, "-");
}

/** '543‐111 Dawn Blue‐Eggshell' → '543-111' */
export function colorCode(raw: Cell): string {
  if (raw === null || raw === "") return "";
  return normHyphen(String(raw).trim()).split(/\s+/)[0] ?? "";
}

/** 81373 or 'F81373' or 'f81373' → 'F81373' */
export function artNo(raw: Cell): string {
  const s = String(raw ?? "").trim().replace(/^[Ff]+/, "");
  return `F${s}`;
}

function cellValue(v: ExcelJS.CellValue | undefined): Cell {
  if (v === null || v === undefined) return null;
  if (typeof v === "number" || typeof v === "string" || typeof v === "boolean") {
    return v;
  }
  if (v instanceof Date) return v.toISOString();
  const o = v as unknown as Record<string, unknown>;
  if ("result" in o) return cellValue(o.result as ExcelJS.CellValue);
  if ("richText" in o) {
    const parts = (o.richText ?? []) as Array<{ text?: string }>;
    return parts.map((p) => p.text ?? "").join("");
  }
  if ("text" in o) return String(o.text);
  return null;
}

function readRows(ws: ExcelJS.Worksheet): Cell[][] {
  const rows: Cell[][] = [];
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    while (rows.length < rowNumber - 1) rows.push([]);
    const cells: Cell[] = [];
    for (let c = 1; c <= ws.columnCount; c += 1) {
      cells.push(cellValue(row.getCell(c).value));
    }
    rows.push(cells);
  });
  return rows;
}

function asInt(v: Cell): number | null {
  if (typeof v === "number") return Number.isFinite(v) ? Math.trunc(v) : null;
  if (typeof v === "string") {
    const t = v.trim();
    if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(t)) {
      const n = Number(t);
      return Number.isFinite(n) ? Math.trunc(n) : null;
    }
  }
  return null;
}

/** Scan PL sheet for P/O NR row */
function poNumber(rows: Cell[][]): string {
  for (const row of rows.slice(0, 20)) {
    const flat = row.filter((c) => c !== null);
    if (flat.length >= 2) {
      const first = String(flat[0]).trim().toUpperCase().replace(/[\s:]+$/, "");
      if (first === "P/O NR" || first === "P/O NO" || first === "PO NR") {
        return String(flat[1]).trim();
      }
    }
  }
  return "";
}

/** 0-based index of the first data row (one after the From/To header row) */
function findDataStart(rows: Cell[][]): number {
  for (let i = 0; i < rows.length; i += 1) {
    const flat = rows[i].filter((c) => c !== null).map((c) => String(c).trim());
    if (flat.includes("From") && flat.includes("To")) {
      return i + 1;
    }
  }
  throw new Error("Could not find 'From'/'To' header in PL sheet");
}

export const KNOWN_SIZES = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL"];

/** PL header order for sizing columns on the mixed-carton sticker table. */
export function sizeRank(size: string): number {
  const i = KNOWN_SIZES.indexOf(size.toUpperCase());
  return i === -1 ? KNOWN_SIZES.length : i;
}

/** {size: 0-based column index}, ordered as they appear in the header row */
function detectSizeColumns(headerRow: Cell[] | undefined): Array<[string, number]> {
  const sizes: Array<[string, number]> = [];
  (headerRow ?? []).forEach((cell, i) => {
    const v = cell === null ? "" : String(cell).trim().toUpperCase();
    if (KNOWN_SIZES.includes(v) && !sizes.some(([s]) => s === v)) {
      sizes.push([v, i]);
    }
  });
  return sizes;
}

const COL_FROM = 1;
const COL_TO = 2;
const COL_STYLE = 4;
const COL_COLOR = 5;
const COL_TOTAL_CARTON = 15;

export async function loadWorkbook(
  fileBuffer: Buffer,
): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(
    fileBuffer as unknown as Parameters<typeof wb.xlsx.load>[0],
  );
  return wb;
}

/** Locate art/color/size/barcode columns by header name (export layouts vary). */
function detectBarcodeColumns(
  rows: Cell[][],
): { art: number; color: number; size: number; barcode: number } | null {
  for (const row of rows.slice(0, 15)) {
    const norm = row.map((c) =>
      c === null ? "" : String(c).trim().toLowerCase(),
    );
    const pick = (...names: string[]) => norm.findIndex((v) => names.includes(v));
    const barcode = pick("barcode");
    if (barcode === -1) continue;
    const art = pick("art no:", "art no", "item number");
    const color = pick("color id", "color", "colour");
    const size = pick("size", "sizes");
    if (art !== -1 && color !== -1 && size !== -1) return { art, color, size, barcode };
  }
  return null;
}

/** barcode file: (art_bare, color_code_normalized, size_upper) → barcode int */
export async function buildBarcodeMap(
  fileBuffer: Buffer,
): Promise<Map<string, number>> {
  const wb = await loadWorkbook(fileBuffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("Barcode file has no worksheets");
  const allRows = readRows(ws);
  const cols = detectBarcodeColumns(allRows) ?? {
    art: 1,
    color: 2,
    size: 3,
    barcode: 4,
  };
  const barcodes = new Map<string, number>();
  for (const row of allRows) {
    const artRaw = row[cols.art] ?? null;
    const colorRaw = row[cols.color] ?? null;
    const sizeRaw = row[cols.size] ?? null;
    const bcRaw = row[cols.barcode] ?? null;
    if (!artRaw || !bcRaw) continue;
    const bc = asInt(bcRaw);
    if (bc === null) continue;
    const art = String(artRaw).trim().replace(/^[Ff]+/, "");
    const col = colorRaw
      ? normHyphen(String(colorRaw).split("/")[0].trim()).toUpperCase()
      : "";
    const size = sizeRaw ? String(sizeRaw).trim().toUpperCase() : "";
    if (art && col && size) {
      barcodes.set(lookupKey(art, col, size), bc);
    }
  }
  return barcodes;
}

function lookupKey(artBare: string, colNorm: string, size: string): string {
  return `${artBare}\u0000${colNorm}\u0000${size}`;
}

/** Lookup barcode, normalizing color for matching. */
function getBarcode(
  barcodes: Map<string, number>,
  artBare: string,
  colRaw: string,
  size: string,
): number | undefined {
  const colNorm = normHyphen(colorCode(colRaw || null)).toUpperCase();
  return barcodes.get(lookupKey(artBare, colNorm, size.toUpperCase()));
}

export async function parsePl(
  fileBuffer: Buffer,
): Promise<{ records: PlRecord[]; totalCartons: number | null; po: string }> {
  const wb = await loadWorkbook(fileBuffer);
  const ws = wb.worksheets.find((w) => w.name === "PL");
  if (!ws) throw new Error("Worksheet PL does not exist.");

  const rows = readRows(ws);
  const po = poNumber(rows);
  const dataStart = findDataStart(rows);
  const sizeCols = detectSizeColumns(rows[dataStart - 1]);

  const records: PlRecord[] = [];
  let totalCartons: number | null = null;

  for (let i = dataStart; i < rows.length; i += 1) {
    const row = rows[i];
    const boxFrom = row[COL_FROM] ?? null;

    // detect TOTAL row
    if (boxFrom !== null && String(boxFrom).trim().toUpperCase() === "TOTAL") {
      const t = asInt(row[COL_TOTAL_CARTON] ?? null);
      if (t !== null) totalCartons = t;
      break;
    }

    if (typeof boxFrom !== "number") continue;

    const boxFromI = Math.trunc(boxFrom);
    const boxToRaw = row[COL_TO];
    const boxTo =
      typeof boxToRaw === "number" ? Math.trunc(boxToRaw) : boxFromI;
    const style = row[COL_STYLE] ?? null;
    const color = row[COL_COLOR] ?? null;

    // gather filled sizes
    const filled: Array<{ size: string; qty: number }> = [];
    for (const [sz, ci] of sizeCols) {
      const v = row[ci];
      if (v === null || v === undefined) continue;
      const q = asInt(v);
      if (q !== null && q > 0) filled.push({ size: sz, qty: q });
    }

    const art = artNo(style);
    const artBare = art.replace(/^F/, "");
    const colStr = color ? String(color).trim().split(/\s+/)[0] : "";

    for (let box = boxFromI; box <= boxTo; box += 1) {
      // A mixed carton becomes one row per polybag (article/size cell);
      // a solid carton has exactly one cell, so it stays one row.
      for (const { size, qty } of filled) {
        records.push({
          art,
          artBare,
          col: colStr,
          colRaw: color ? String(color) : "",
          size,
          qty,
          box,
        });
      }
    }
  }

  return { records, totalCartons, po };
}

const thinSide = { style: "thin" as const };
const cellBorder = {
  top: thinSide,
  left: thinSide,
  bottom: thinSide,
  right: thinSide,
};
const center = { horizontal: "center" as const, vertical: "middle" as const };

export async function writeOutput(
  records: PlRecord[],
  barcodes: Map<string, number>,
  totalCartons: number | null,
  po: string,
  lpnStart: string,
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("SOLID SKU");

  const headers = [
    "LPN",
    "Oder no",
    "Purchase order number",
    "Article no",
    "Color no",
    "Size",
    "QTY/pcs",
    "EAN 13 SKU barcode",
    "CARTON BOX NUMBER",
  ];
  const subhdr = ["a", "b", "c", "d", "e", "f", "g", "h", "i"];

  const headerRow = ws.addRow(headers);
  ws.addRow(subhdr);
  for (let c = 1; c <= headers.length; c += 1) {
    const cell = headerRow.getCell(c);
    cell.font = { name: "Arial", bold: true, size: 10 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFBDD7EE" } };
    cell.alignment = center;
    cell.border = cellBorder;
  }
  const subRow = ws.getRow(2);
  for (let c = 1; c <= subhdr.length; c += 1) {
    const cell = subRow.getCell(c);
    cell.font = { name: "Arial", bold: true, italic: true, size: 9 };
    cell.alignment = center;
  }

  const tc = totalCartons ?? records.length;
  let lpn = BigInt(lpnStart);

  for (const r of records) {
    const bc = getBarcode(barcodes, r.artBare, r.colRaw, r.size);
    const row = ws.addRow([
      String(lpn),
      po,
      po,
      r.art,
      r.col,
      r.size,
      String(r.qty),
      bc === undefined ? null : String(bc),
      `${r.box}/${tc}`,
    ]);
    const font = { name: "Arial", size: 10 };
    for (let c = 1; c <= 9; c += 1) {
      const cell = row.getCell(c);
      cell.font = font;
      cell.alignment = center;
      cell.border = cellBorder;
    }
    lpn += BigInt(1);
  }

  const colWidths = [18, 12, 22, 14, 12, 8, 10, 22, 18, 14];
  colWidths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  const out = await wb.xlsx.writeBuffer();
  return Buffer.from(out);
}

export async function generate(
  plBuffer: Buffer,
  barcodeBuffer: Buffer,
  lpnStart: string,
): Promise<GenerateResult> {
  const barcodes = await buildBarcodeMap(barcodeBuffer);
  const { records, totalCartons, po } = await parsePl(plBuffer);
  const buffer = await writeOutput(records, barcodes, totalCartons, po, lpnStart);
  const boxCounts = new Map<number, number>();
  for (const r of records) {
    boxCounts.set(r.box, (boxCounts.get(r.box) ?? 0) + 1);
  }
  let mixedCartons = 0;
  for (const n of boxCounts.values()) {
    if (n > 1) mixedCartons += 1;
  }
  return {
    buffer,
    po,
    totalCartons,
    rowCount: records.length,
    mixedCount: mixedCartons,
  };
}

export const DEFAULT_OUTPUT_NAME = "SOLID_SKU_output.xlsx";

export function normalizeOutputName(raw: string): string {
  // strip any directory parts, mirror the tkinter "save into PL folder" behavior
  const base = raw.trim().split(/[/\\]/).pop() ?? "";
  if (!base) return DEFAULT_OUTPUT_NAME;
  const lower = base.toLowerCase();
  return lower.endsWith(".xlsx") || lower.endsWith(".xlsm") ? base : `${base}.xlsx`;
}
