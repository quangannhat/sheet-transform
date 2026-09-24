import {
  buildPolybagLabelsPdf,
  filterMixedRows,
  parseSolidSkuRows,
  polybagLabelFileName,
} from "@/lib/cartonLabels";

export const runtime = "nodejs";

function badRequest(errors: string[]): Response {
  return Response.json({ errors }, { status: 422 });
}

export async function POST(request: Request): Promise<Response> {
  let formData: FormData;
  try {
    formData = await request.formData();
  } catch {
    return badRequest(["Could not read the submitted form."]);
  }

  const file = formData.get("file");
  if (!(file instanceof File) || file.size === 0) {
    return badRequest(["• SOLID SKU file not selected"]);
  }
  if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    return badRequest(["• File must be an .xlsx / .xlsm workbook"]);
  }

  try {
    const rows = await parseSolidSkuRows(
      Buffer.from(await file.arrayBuffer()),
    );
    const pdf = await buildPolybagLabelsPdf(rows);
    const fileName = polybagLabelFileName(file.name);
    return new Response(new Uint8Array(pdf), {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
        "X-Row-Count": String(filterMixedRows(rows).length),
      },
    });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return badRequest([`Error: ${message}`]);
  }
}
