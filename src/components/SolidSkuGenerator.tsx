"use client";

import { useRef, useState } from "react";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

type GeneratedFile = { url: string; name: string };

type Status =
  | { kind: "idle" }
  | { kind: "error"; messages: string[] }
  | {
      kind: "done";
      xlsx: GeneratedFile;
      pdf?: GeneratedFile;
      pdfError?: string;
      polybag?: GeneratedFile;
      polybagError?: string;
      polybagSkipped?: boolean;
      rowCount: number;
      mixedCount: number;
    };

function download(blobUrl: string, name: string) {
  const a = document.createElement("a");
  a.href = blobUrl;
  a.download = name;
  a.click();
}

function Errors({ messages }: { messages: string[] }) {
  return (
    <div className="rounded-lg bg-red-50 px-4 py-3 text-sm text-red-700 dark:bg-red-950/40 dark:text-red-300">
      <ul className="list-inside space-y-1 whitespace-pre-wrap">
        {messages.map((m) => (
          <li key={m}>{m}</li>
        ))}
      </ul>
    </div>
  );
}

const fileInputClass =
  "block w-full cursor-pointer rounded-lg border border-black/[.12] p-2 text-sm text-zinc-600 file:mr-3 file:rounded-md file:border-0 file:bg-zinc-900 file:px-3 file:py-1.5 file:text-sm file:font-medium file:text-white hover:file:bg-zinc-700 dark:border-white/[.145] dark:text-zinc-400 dark:file:bg-white dark:file:text-black";
const labelClass =
  "flex flex-col gap-1.5 text-sm font-medium text-zinc-700 dark:text-zinc-300";
const cardClass =
  "flex flex-col gap-4 rounded-xl border border-black/[.08] bg-white p-5 dark:border-white/[.145] dark:bg-black";
const inputClass =
  "rounded-lg border border-black/[.12] p-2 text-sm text-zinc-800 dark:border-white/[.145] dark:text-zinc-200";
const buttonClass =
  "h-11 shrink-0 self-start rounded-full bg-blue-600 px-8 text-sm font-semibold text-white transition-colors hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-50";

async function readErrors(res: Response): Promise<string[]> {
  const body = (await res.json().catch(() => null)) as {
    errors?: string[];
  } | null;
  return body?.errors ?? [`Request failed (${res.status})`];
}

function fileNameFrom(res: Response, fallback: string): string {
  return (
    decodeURIComponent(
      /filename\*=UTF-8''([^;]+)/i
        .exec(res.headers.get("content-disposition") ?? "")?.[1] ?? "",
    ) || fallback
  );
}

export function SolidSkuGenerator() {
  const [status, setStatus] = useState<Status>({ kind: "idle" });
  const [pending, setPending] = useState(false);
  const urlsRef = useRef<string[]>([]);

  const revokeAll = () => {
    urlsRef.current.forEach((u) => URL.revokeObjectURL(u));
    urlsRef.current = [];
  };

  async function onSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    const data = new FormData(form);
    const serialBase = String(data.get("serialBase") ?? "").trim();

    setPending(true);
    revokeAll();
    try {
      // step 1: build the SOLID SKU workbook
      const res = await fetch("/api/solid-sku", { method: "POST", body: data });
      if (!(res.headers.get("content-type") ?? "").includes("spreadsheetml")) {
        setStatus({ kind: "error", messages: await readErrors(res) });
        return;
      }

      const xlsxName = fileNameFrom(res, "SOLID_SKU_output.xlsx");
      const xlsxUrl = URL.createObjectURL(await res.blob());
      urlsRef.current.push(xlsxUrl);
      download(xlsxUrl, xlsxName);
      // reuse the generated workbook as input for the two PDF steps
      const xlsxBlob = await blobFromXlsx(xlsxUrl);
      const rowCount = Number(res.headers.get("x-row-count") ?? 0);
      const mixedCount = Number(res.headers.get("x-mixed-count") ?? 0);

      const xlsxFile = () =>
        new File([xlsxBlob], xlsxName, { type: XLSX_MIME });

      // step 2: carton-labels PDF from the generated workbook
      let pdf: GeneratedFile | undefined;
      let pdfError: string | undefined;
      const pdfData = new FormData();
      pdfData.set("file", xlsxFile());
      pdfData.set("serialBase", serialBase);
      const pdfRes = await fetch("/api/solid-sku/labels", {
        method: "POST",
        body: pdfData,
      });
      if (
        (pdfRes.headers.get("content-type") ?? "").includes("application/pdf")
      ) {
        const pdfName = fileNameFrom(pdfRes, "SOLID_SKU labels.pdf");
        const pdfUrl = URL.createObjectURL(await pdfRes.blob());
        urlsRef.current.push(pdfUrl);
        download(pdfUrl, pdfName);
        pdf = { url: pdfUrl, name: pdfName };
      } else {
        pdfError = (await readErrors(pdfRes)).join("\n");
      }

      // step 3: polybag LPN stickers from the same workbook (mixed cartons)
      let polybag: GeneratedFile | undefined;
      let polybagError: string | undefined;
      let polybagSkipped = false;
      if (mixedCount > 0) {
        const polyData = new FormData();
        polyData.set("file", xlsxFile());
        const polyRes = await fetch("/api/solid-sku/polybags", {
          method: "POST",
          body: polyData,
        });
        if (
          (polyRes.headers.get("content-type") ?? "").includes("application/pdf")
        ) {
          const polyName = fileNameFrom(polyRes, "SOLID_SKU polybag labels.pdf");
          const polyUrl = URL.createObjectURL(await polyRes.blob());
          urlsRef.current.push(polyUrl);
          download(polyUrl, polyName);
          polybag = { url: polyUrl, name: polyName };
        } else {
          polybagError = (await readErrors(polyRes)).join("\n");
        }
      } else {
        polybagSkipped = true;
      }

      setStatus({
        kind: "done",
        xlsx: { url: xlsxUrl, name: xlsxName },
        pdf,
        pdfError,
        polybag,
        polybagError,
        polybagSkipped,
        rowCount,
        mixedCount,
      });
    } catch {
      setStatus({ kind: "error", messages: ["Network request failed."] });
    } finally {
      setPending(false);
    }
  }

  return (
    <div className="flex w-full flex-col gap-6">
      <form
        onSubmit={onSubmit}
        className={cardClass}
      >
        <label className={labelClass}>
          INV-PL file (must contain a &ldquo;PL&rdquo; sheet)
          <input
            type="file"
            name="plFile"
            accept=".xlsx,.xlsm"
            required
            className={fileInputClass}
          />
        </label>

        <label className={labelClass}>
          Barcode file (article / color / size &rarr; EAN)
          <input
            type="file"
            name="barcodeFile"
            accept=".xlsx,.xlsm"
            required
            className={fileInputClass}
          />
        </label>

        <div className="flex flex-col gap-4 sm:flex-row">
          <label className={`flex-1 ${labelClass}`}>
            LPN start number
            <input
              type="text"
              inputMode="numeric"
              name="lpn"
              required
              placeholder="1"
              className={inputClass}
            />
          </label>
          <label className={`flex-1 ${labelClass}`}>
            Carton serial base (PDF barcodes)
            <input
              type="text"
              inputMode="numeric"
              name="serialBase"
              placeholder="12471600000049"
              className={inputClass}
            />
          </label>
          <label className={`flex-[1.5] ${labelClass}`}>
            Output filename
            <input
              type="text"
              name="output"
              placeholder="SOLID_SKU_output.xlsx"
              className={inputClass}
            />
          </label>
        </div>

        <button type="submit" disabled={pending} className={buttonClass}>
          {pending ? "Working…" : "Generate →"}
        </button>
      </form>

      {status.kind === "error" && <Errors messages={status.messages} />}

      {status.kind === "done" && (
        <div className="flex flex-col gap-2 rounded-lg bg-green-50 px-4 py-3 text-sm text-green-800 dark:bg-green-950/40 dark:text-green-300">
          <p>
            &#10003; {status.rowCount} rows &middot; {status.mixedCount} mixed
            cartons &middot; saved to{" "}
            <span className="font-medium">{status.xlsx.name}</span>
          </p>
          {status.pdf && (
            <p>
              &#10003; carton labels PDF ({status.pdf.name})
            </p>
          )}
          {status.pdfError && (
            <p className="text-red-600 dark:text-red-400">
              &#10007; PDF step failed: {status.pdfError}
            </p>
          )}
          {status.polybag && (
            <p>
              &#10003; polybag labels PDF ({status.polybag.name})
            </p>
          )}
          {status.polybagError && (
            <p className="text-red-600 dark:text-red-400">
              &#10007; polybag step failed: {status.polybagError}
            </p>
          )}
          {status.polybagSkipped && (
            <p>No mixed cartons &mdash; polybag labels skipped.</p>
          )}
          <p className="text-green-700 dark:text-green-400">
            Downloads didn&rsquo;t start?{" "}
            <a
              href={status.xlsx.url}
              download={status.xlsx.name}
              className="underline underline-offset-2"
            >
              {status.xlsx.name}
            </a>
            {status.pdf && (
              <>
                {" · "}
                <a
                  href={status.pdf.url}
                  download={status.pdf.name}
                  className="underline underline-offset-2"
                >
                  {status.pdf.name}
                </a>
              </>
            )}
            {status.polybag && (
              <>
                {" · "}
                <a
                  href={status.polybag.url}
                  download={status.polybag.name}
                  className="underline underline-offset-2"
                >
                  {status.polybag.name}
                </a>
              </>
            )}
          </p>
        </div>
      )}

      <LabelsPdfCard />
      <PolybagLabelsCard />
    </div>
  );
}

async function blobFromXlsx(url: string): Promise<Blob> {
  const res = await fetch(url);
  return res.blob();
}

function SolidSkuPdfCard({
  title,
  endpoint,
  fallbackName,
  buttonLabel,
  withSerialBase,
}: {
  title: string;
  endpoint: string;
  fallbackName: string;
  buttonLabel: string;
  withSerialBase: boolean;
}) {
  const [pending, setPending] = useState(false);
  const [error, setError] = useState<string[] | null>(null);
  const [done, setDone] = useState<GeneratedFile | null>(null);

  async function onSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setPending(true);
    setError(null);
    setDone(null);
    try {
      const res = await fetch(endpoint, {
        method: "POST",
        body: new FormData(event.currentTarget),
      });
      if (!(res.headers.get("content-type") ?? "").includes("application/pdf")) {
        setError(await readErrors(res));
        return;
      }
      const name =
        decodeURIComponent(
          /filename\*=UTF-8''([^;]+)/i
            .exec(res.headers.get("content-disposition") ?? "")?.[1] ?? "",
        ) || fallbackName;
      const url = URL.createObjectURL(await res.blob());
      download(url, name);
      setDone({ url, name });
    } catch {
      setError(["Network request failed."]);
    } finally {
      setPending(false);
    }
  }

  return (
    <form onSubmit={onSubmit} className={cardClass}>
      <h2 className="text-base font-semibold text-zinc-900 dark:text-zinc-100">
        {title}
      </h2>
      <label className={labelClass}>
        SOLID SKU workbook (generated output)
        <input
          type="file"
          name="file"
          accept=".xlsx,.xlsm"
          required
          className={fileInputClass}
        />
      </label>
      {withSerialBase && (
        <label className={`sm:max-w-xs ${labelClass}`}>
          Carton serial base (optional)
          <input
            type="text"
            inputMode="numeric"
            name="serialBase"
            placeholder="12471600000049"
            className={inputClass}
          />
        </label>
      )}
      <button type="submit" disabled={pending} className={buttonClass}>
        {pending ? "Working…" : buttonLabel}
      </button>
      {error && <Errors messages={error} />}
      {done && (
        <p className="text-sm text-green-700 dark:text-green-400">
          &#10003; {done.name} downloaded — didn&rsquo;t start?{" "}
          <a
            href={done.url}
            download={done.name}
            className="underline underline-offset-2"
          >
            {done.name}
          </a>
        </p>
      )}
    </form>
  );
}

function LabelsPdfCard() {
  return (
    <SolidSkuPdfCard
      title="Carton labels PDF from an existing SOLID SKU file"
      endpoint="/api/solid-sku/labels"
      fallbackName="carton labels.pdf"
      buttonLabel="Download labels PDF →"
      withSerialBase
    />
  );
}

function PolybagLabelsCard() {
  return (
    <SolidSkuPdfCard
      title="Polybag labels PDF (LPN on each polybag of mixed cartons)"
      endpoint="/api/solid-sku/polybags"
      fallbackName="polybag labels.pdf"
      buttonLabel="Download polybag labels PDF →"
      withSerialBase={false}
    />
  );
}
