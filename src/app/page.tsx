import { SolidSkuGenerator } from "@/components/SolidSkuGenerator";

const BUILD_SHA = (process.env.NEXT_PUBLIC_BUILD_SHA ?? "").slice(0, 7);
const BUILD_TIME = process.env.NEXT_PUBLIC_BUILD_TIME ?? "";

export default function Home() {
  return (
    <main className="mx-auto flex w-full max-w-4xl flex-1 flex-col gap-8 px-6 py-16">
      <header className="flex flex-col gap-2">
        <h1 className="text-3xl font-semibold tracking-tight text-zinc-900 dark:text-zinc-50">
          SOLID SKU Generator
        </h1>
        <p className="max-w-prose text-zinc-600 dark:text-zinc-400">
          Upload the INV-PL workbook (it must contain a &ldquo;PL&rdquo; sheet)
          and the barcode file, and the server builds one Excel row per carton
          with the EAN lookup and LPN numbering &mdash; mixed cartons become one
          polybag row per size cell (each with its own LPN/EAN), exactly like
          <code className="rounded bg-zinc-100 px-1 py-0.5 text-[0.85em] dark:bg-zinc-800">solid_sku_gen.py</code>.
          A carton-labels PDF (one A5 label sheet per carton) and a polybag LPN
          sticker PDF (version 1+2, mixed-carton rows only) are generated from
          that workbook as the second and third steps; mixed cartons get the
          supplier-manual version-2 box sticker without an LPN.
        </p>
      </header>
      <SolidSkuGenerator />
      <footer className="flex flex-wrap items-baseline gap-x-3 text-xs text-zinc-500">
        <span className="rounded-full border border-zinc-300 px-2 py-0.5 font-medium text-zinc-600 dark:border-zinc-700 dark:text-zinc-400">
          v8 &middot; one upload &rarr; workbook + carton + polybag PDFs
        </span>
        <span aria-live="polite">
          {BUILD_SHA
            ? `deployed build ${BUILD_SHA}${BUILD_TIME ? ` · ${BUILD_TIME}` : ""}`
            : "local dev build"}
        </span>
      </footer>
    </main>
  );
}
