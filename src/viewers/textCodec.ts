// Decoding / encoding of plain-text files so that saving keeps the original encoding,
// byte-order mark and line endings. Editors work with "\n" only.

export type TextEncoding = "utf-8" | "utf-16le" | "utf-16be" | "windows-1251";

export interface TextFormat {
  encoding: TextEncoding;
  bom: boolean;
  eol: "\r\n" | "\n";
}

export function decodeText(data: ArrayBuffer): { text: string; format: TextFormat } {
  const bytes = new Uint8Array(data);
  let encoding: TextEncoding;
  let bom = false;
  let raw: string;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    encoding = "utf-8";
    bom = true;
    raw = new TextDecoder("utf-8").decode(bytes.subarray(3));
  } else if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    encoding = "utf-16le";
    bom = true;
    raw = new TextDecoder("utf-16le").decode(bytes.subarray(2));
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    encoding = "utf-16be";
    bom = true;
    raw = new TextDecoder("utf-16be").decode(bytes.subarray(2));
  } else {
    try {
      raw = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
      encoding = "utf-8";
    } catch {
      // Cyrillic text files saved by old Windows tools are very often in Windows-1251.
      raw = new TextDecoder("windows-1251").decode(bytes);
      encoding = "windows-1251";
    }
  }
  const eol = raw.includes("\r\n") ? "\r\n" : "\n";
  return { text: raw.replace(/\r\n?/g, "\n"), format: { encoding, bom, eol } };
}

export function encodeText(text: string, format: TextFormat): { data: Uint8Array; format: TextFormat; note?: string } {
  if (format.eol === "\r\n") text = text.replace(/\n/g, "\r\n");
  switch (format.encoding) {
    case "utf-16le":
    case "utf-16be":
      return { data: encodeUtf16(text, format.encoding === "utf-16le", format.bom), format };
    case "windows-1251": {
      const data = encode1251(text);
      if (data) return { data, format };
      const utf8 = { ...format, encoding: "utf-8" as const, bom: false };
      return {
        data: encodeUtf8(text, false),
        format: utf8,
        note: "В тексте есть символы, которых нет в Windows-1251, — файл сохранён в UTF-8",
      };
    }
    default:
      return { data: encodeUtf8(text, format.bom), format };
  }
}

export function formatLabel(f: TextFormat): string {
  const enc = { "utf-8": "UTF-8", "utf-16le": "UTF-16 LE", "utf-16be": "UTF-16 BE", "windows-1251": "Windows-1251" }[f.encoding];
  return `${enc}${f.bom && f.encoding === "utf-8" ? " BOM" : ""} · ${f.eol === "\r\n" ? "CRLF" : "LF"}`;
}

function encodeUtf8(text: string, bom: boolean): Uint8Array {
  const body = new TextEncoder().encode(text);
  if (!bom) return body;
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf]);
  out.set(body, 3);
  return out;
}

function encodeUtf16(text: string, le: boolean, bom: boolean): Uint8Array {
  const off = bom ? 2 : 0;
  const out = new Uint8Array(off + text.length * 2);
  const view = new DataView(out.buffer);
  if (bom) view.setUint16(0, 0xfeff, le);
  for (let i = 0; i < text.length; i++) view.setUint16(off + i * 2, text.charCodeAt(i), le);
  return out;
}

let table1251: Map<number, number> | null = null;

/** Returns null if the text has characters Windows-1251 cannot represent. */
function encode1251(text: string): Uint8Array | null {
  if (!table1251) {
    table1251 = new Map();
    const dec = new TextDecoder("windows-1251");
    for (let b = 0x80; b <= 0xff; b++) table1251.set(dec.decode(new Uint8Array([b])).charCodeAt(0), b);
  }
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    const b = c < 0x80 ? c : table1251.get(c);
    if (b === undefined) return null;
    out[i] = b;
  }
  return out;
}
