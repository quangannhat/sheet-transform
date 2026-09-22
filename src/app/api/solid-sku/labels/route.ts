import {
  buildLabelsPdf,
  cartonLabelFileName,
  parseSolidSkuRows,
  resolveSerialBase,
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
  const errors: string[] = [];
  if (!(file instanceof File) || file.size === 0) {
    errors.push("• SOLID SKU file not selected");
  } else if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    errors.push("• File must be an .xlsx / .xlsm workbook");
  }

  const { base, valid } = resolveSerialBase(
    String(formData.get("serialBase") ?? ""),
  );
  if (!valid) errors.push("• Carton serial base must be a number");

  if (errors.length > 0) return badRequest(errors);

  try {
    const rows = await parseSolidSkuRows(
      Buffer.from(await (file as File).arrayBuffer()),
    );
    const pdf = await buildLabelsPdf(rows, base);
    const fileName = cartonLabelFileName((file as File).name);
    return new Response(new Uint8Array(pdf), {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
        "X-Row-Count": String(rows.length),
      },
    });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return badRequest([`Error: ${message}`]);
  }
}
