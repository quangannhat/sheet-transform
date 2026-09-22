import { SolidSkuGenerator } from "@/components/SolidSkuGenerator";

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
          with the EAN lookup, LPN numbering and mixed-carton flags &mdash;
          exactly like <code className="rounded bg-zinc-100 px-1 py-0.5 text-[0.85em] dark:bg-zinc-800">solid_sku_gen.py</code>.
        </p>
      </header>
      <SolidSkuGenerator />
    </main>
  );
}
