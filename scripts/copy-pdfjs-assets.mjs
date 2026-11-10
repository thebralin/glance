// Copies PDF.js runtime assets (CMaps for CJK/Cyrillic fonts, standard fonts, wasm image
// decoders, ICC profiles) into public/ so they are bundled with the app and work offline.
import { cpSync, existsSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const src = join(root, "node_modules", "pdfjs-dist");
const dst = join(root, "public", "pdfjs");

mkdirSync(dst, { recursive: true });
for (const dir of ["cmaps", "standard_fonts", "wasm", "iccs"]) {
  const from = join(src, dir);
  if (existsSync(from)) cpSync(from, join(dst, dir), { recursive: true });
}
console.log("pdf.js assets copied to public/pdfjs");
