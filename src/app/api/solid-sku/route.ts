import {
  generate,
  normalizeOutputName,
} from "@/lib/solidSku";

export const runtime = "nodejs";

const ACCEPTED_EXTENSIONS = [".xlsx", ".xlsm"];

function badRequest(errors: string[]): Response {
  return Response.json({ errors }, { status: 422 });
}

function isExcel(file: unknown): file is File {
  if (!(file instanceof File) || file.size === 0) return false;
  const name = file.name.toLowerCase();
  return ACCEPTED_EXTENSIONS.some((ext) => name.endsWith(ext));
}

export async function POST(request: Request): Promise<Response> {
  let formData: FormData;
  try {
    formData = await request.formData();
  } catch {
    return badRequest(["Could not read the submitted form."]);
  }

  const plFile = formData.get("plFile");
  const barcodeFile = formData.get("barcodeFile");
  const lpnRaw = String(formData.get("lpn") ?? "").trim();
  const outputRaw = String(formData.get("output") ?? "").trim();

  const errors: string[] = [];
  if (!(plFile instanceof File) || plFile.size === 0) {
    errors.push("• INV-PL file not selected");
  } else if (!isExcel(plFile)) {
    errors.push("• INV-PL file must be .xlsx or .xlsm");
  }
  if (!(barcodeFile instanceof File) || barcodeFile.size === 0) {
    errors.push("• Barcode file not selected");
  } else if (!isExcel(barcodeFile)) {
    errors.push("• Barcode file must be .xlsx or .xlsm");
  }
  let lpnStart = 0;
  if (!lpnRaw) {
    errors.push("• LPN start number is required");
  } else if (!/^-?\d+$/.test(lpnRaw)) {
    errors.push("• LPN start must be a number");
  } else {
    lpnStart = Number.parseInt(lpnRaw, 10);
  }
  if (errors.length > 0) return badRequest(errors);

  try {
    const [plArrayBuffer, barcodeArrayBuffer] = await Promise.all([
      (plFile as File).arrayBuffer(),
      (barcodeFile as File).arrayBuffer(),
    ]);
    const result = await generate(
      Buffer.from(plArrayBuffer),
      Buffer.from(barcodeArrayBuffer),
      lpnStart,
    );

    const fileName = normalizeOutputName(outputRaw);
    return new Response(new Uint8Array(result.buffer), {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
        "X-Po": encodeURIComponent(result.po),
        "X-Row-Count": String(result.rowCount),
        "X-Mixed-Count": String(result.mixedCount),
        "X-Total-Cartons": String(result.totalCartons ?? ""),
      },
    });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return badRequest([`Error: ${message}`]);
  }
}
