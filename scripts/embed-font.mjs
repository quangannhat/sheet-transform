import fs from "fs";

const JOBS = [
  ["src/lib/fonts/Arial-Bold.ttf", "src/lib/fonts/arialBold.ts", "ARIAL_BOLD_TTF_BASE64"],
  ["src/lib/fonts/Arial-Regular.ttf", "src/lib/fonts/arialRegular.ts", "ARIAL_REGULAR_TTF_BASE64"],
];

for (const [ttfPath, outPath, exportName] of JOBS) {
  const ttf = fs.readFileSync(ttfPath);
  const b64 = ttf.toString("base64");
  const out =
    "// Arial-compatible TTF (Liberation Sans, metric-identical to Arial, OFL-licensed),\n" +
    "// embedded as base64 so it is always present in the server bundle (no fs tracing needed).\n" +
    `// Regenerate: node scripts/embed-font.mjs  (after replacing ${ttfPath})\n` +
    "// Keep as a single string literal: a `+`-concatenated chain overflows the TS checker.\n" +
    `export const ${exportName} = "${b64}";\n`;
  fs.writeFileSync(outPath, out);
  const decoded = Buffer.from(b64, "base64");
  console.log(ttfPath, "→", outPath, ttf.length, "bytes, magic", decoded.slice(0, 4).toString("hex"));
}
