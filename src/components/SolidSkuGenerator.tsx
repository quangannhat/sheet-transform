"use client";

import { useRef, useState } from "react";

type Status =
  | { kind: "idle" }
  | { kind: "error"; messages: string[] }
  | {
      kind: "done";
      fileName: string;
      rowCount: number;
      mixedCount: number;
      url: string;
    };

export function SolidSkuGenerator() {
  const [status, setStatus] = useState<Status>({ kind: "idle" });
  const [pending, setPending] = useState(false);
  const urlRef = useRef<string | null>(null);

  const revokeUrl = () => {
    if (urlRef.current) {
      URL.revokeObjectURL(urlRef.current);
      urlRef.current = null;
    }
  };

  async function onSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    const data = new FormData(form);

    setPending(true);
    revokeUrl();
    try {
      const res = await fetch("/api/solid-sku", { method: "POST", body: data });
      const contentType = res.headers.get("content-type") ?? "";

      if (!contentType.includes("spreadsheetml")) {
        const body = (await res.json().catch(() => null)) as {
          errors?: string[];
        } | null;
        setStatus({
          kind: "error",
          messages: body?.errors ?? [`Request failed (${res.status})`],
        });
        return;
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      urlRef.current = url;
      const disposition = res.headers.get("content-disposition") ?? "";
      const utf8Match = /filename\*=UTF-8''([^;]+)/i.exec(disposition);
      const fileName = utf8Match
        ? decodeURIComponent(utf8Match[1])
        : "SOLID_SKU_output.xlsx";

      setStatus({
        kind: "done",
        fileName,
        url,
        rowCount: Number(res.headers.get("x-row-count") ?? 0),
        mixedCount: Number(res.headers.get("x-mixed-count") ?? 0),
      });

      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      a.click();
    } catch {
      setStatus({ kind: "error", messages: ["Network request failed."] });
    } finally {
      setPending(false);
    }
  }

  return (
    <div className="flex w-full flex-col gap-6">
      <form onSubmit={onSubmit}
        className="flex flex-col gap-4 rounded-xl border border-black/[.08] bg-white p-5 dark:border-white/[.145] dark:bg-black"
      >
        <label className="flex flex-col gap-1.5 text-sm font-medium text-zinc-700 dark:text-zinc-300">
          INV-PL file (must contain a &ldquo;PL&rdquo; sheet)
          <input
            type="file"
            name="plFile"
            accept=".xlsx,.xlsm"
            required
            className="block w-full cursor-pointer rounded-lg border border-black/[.12] p-2 text-sm text-zinc-600 file:mr-3 file:rounded-md file:border-0 file:bg-zinc-900 file:px-3 file:py-1.5 file:text-sm file:font-medium file:text-white hover:file:bg-zinc-700 dark:border-white/[.145] dark:text-zinc-400 dark:file:bg-white dark:file:text-black"
          />
        </label>

        <label className="flex flex-col gap-1.5 text-sm font-medium text-zinc-700 dark:text-zinc-300">
          Barcode file (article / color / size &rarr; EAN)
          <input
            type="file"
            name="barcodeFile"
            accept=".xlsx,.xlsm"
            required
            className="block w-full cursor-pointer rounded-lg border border-black/[.12] p-2 text-sm text-zinc-600 file:mr-3 file:rounded-md file:border-0 file:bg-zinc-900 file:px-3 file:py-1.5 file:text-sm file:font-medium file:text-white hover:file:bg-zinc-700 dark:border-white/[.145] dark:text-zinc-400 dark:file:bg-white dark:file:text-black"
          />
        </label>

        <div className="flex flex-col gap-4 sm:flex-row">
          <label className="flex flex-1 flex-col gap-1.5 text-sm font-medium text-zinc-700 dark:text-zinc-300">
            LPN start number
            <input
              type="text"
              inputMode="numeric"
              name="lpn"
              required
              placeholder="1"
              className="rounded-lg border border-black/[.12] p-2 text-sm text-zinc-800 dark:border-white/[.145] dark:text-zinc-200"
            />
          </label>
          <label className="flex flex-[2] flex-col gap-1.5 text-sm font-medium text-zinc-700 dark:text-zinc-300">
            Output filename
            <input
              type="text"
              name="output"
              placeholder="SOLID_SKU_output.xlsx"
              className="rounded-lg border border-black/[.12] p-2 text-sm text-zinc-800 dark:border-white/[.145] dark:text-zinc-200"
            />
          </label>
        </div>

        <button
          type="submit"
          disabled={pending}
          className="h-11 shrink-0 self-start rounded-full bg-blue-600 px-8 text-sm font-semibold text-white transition-colors hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-50"
        >
          {pending ? "Working…" : "Generate →"}
        </button>
      </form>

      {status.kind === "error" && (
        <div className="rounded-lg bg-red-50 px-4 py-3 text-sm text-red-700 dark:bg-red-950/40 dark:text-red-300">
          <ul className="list-inside space-y-1 whitespace-pre-wrap">
            {status.messages.map((m) => (
              <li key={m}>{m}</li>
            ))}
          </ul>
        </div>
      )}

      {status.kind === "done" && (
        <div className="flex flex-col gap-2 rounded-lg bg-green-50 px-4 py-3 text-sm text-green-800 dark:bg-green-950/40 dark:text-green-300">
          <p>
            &#10003; {status.rowCount} rows &middot; {status.mixedCount} mixed
            &middot; saved to{" "}
            <span className="font-medium">{status.fileName}</span>
          </p>
          <p className="text-green-700 dark:text-green-400">
            Download didn&rsquo;t start?{" "}
            <a
              href={status.url}
              download={status.fileName}
              className="underline underline-offset-2"
            >
              Download {status.fileName}
            </a>
          </p>
        </div>
      )}
    </div>
  );
}
