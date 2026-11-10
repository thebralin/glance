// Text editing of DOCX paragraphs. docx-preview renders HTML without any link back to the
// XML, so before rendering every <w:p> of the main document gets an invisible bookmark
// carrying its index; the rendered <p> is found by that bookmark. A paragraph is editable
// only when the text we can rebuild from its XML equals the rendered text, so a save can
// never touch content we do not understand. Saving rewrites the text inside the existing
// runs (changed middle part only), which keeps the formatting of everything else.
import JSZip from "jszip";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const MARK = "__glance_p";
/** Word packages have no folder entries; JSZip would add "word/" otherwise. */
const NO_FOLDERS = { createFolders: false };

/** Elements that never contribute visible paragraph text (or belong to nested content). */
const SKIP = new Set(["pPr", "rPr", "p", "drawing", "pict", "object", "del", "moveFrom", "sdtPr", "instrText", "delText"]);
/** Content we cannot round-trip: the paragraph stays read-only. */
const UNSUPPORTED = new Set(["fldChar", "fldSimple", "sym", "cr", "noBreakHyphen", "softHyphen", "ptab", "footnoteReference", "endnoteReference"]);

interface Atom {
  el: Element;
  kind: "t" | "tab" | "br";
  text: string;
}

export interface DocxSource {
  zip: JSZip;
  path: string; // e.g. "word/document.xml"
  xml: string;
  /** Same document with paragraph markers, for rendering only. */
  marked: ArrayBuffer;
}

export async function prepareDocx(data: ArrayBuffer): Promise<DocxSource> {
  const zip = await JSZip.loadAsync(data);
  const path = await mainDocumentPath(zip);
  const xml = await zip.file(path)!.async("string");
  const doc = parseXml(xml);
  const paras = doc.getElementsByTagNameNS(W, "p");
  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    if (insideAlternateContent(p)) continue;
    const bm = doc.createElementNS(W, `${p.prefix ? p.prefix + ":" : ""}bookmarkStart`);
    bm.setAttributeNS(W, attrName(p, "id"), String(900000000 + i));
    bm.setAttributeNS(W, attrName(p, "name"), MARK + i);
    const pPr = firstChild(p, "pPr");
    p.insertBefore(bm, pPr ? pPr.nextSibling : p.firstChild);
  }
  zip.file(path, serializeXml(doc, xml), NO_FOLDERS);
  const marked = await zip.generateAsync({ type: "arraybuffer", compression: "STORE" });
  zip.file(path, xml, NO_FOLDERS);
  return { zip, path, xml, marked };
}

/**
 * Finds rendered paragraphs by their markers and removes the markers. Returns the editable
 * ones keyed by XML paragraph index, with their current text.
 */
export function mapParagraphs(body: HTMLElement, xml: string, tabClass: string): Map<HTMLElement, { index: number; text: string }> {
  const out = new Map<HTMLElement, { index: number; text: string }>();
  const paras = parseXml(xml).getElementsByTagNameNS(W, "p");
  for (const span of body.querySelectorAll<HTMLElement>(`span[id^="${MARK}"]`)) {
    const index = Number(span.id.slice(MARK.length));
    const p = span.parentElement;
    span.remove();
    if (!p || p.tagName !== "P" || !paras[index]) continue;
    const atoms = collectAtoms(paras[index]);
    if (!atoms) continue;
    const text = atoms.map((a) => a.text).join("");
    if (text !== domText(p, tabClass)) continue;
    p.dataset.gp = String(index);
    out.set(p, { index, text });
  }
  return out;
}

/** Text of a rendered paragraph as the XML would spell it (tabs as \t, soft breaks as \n). */
export function domText(p: HTMLElement, tabClass: string): string {
  let s = "";
  let lastBr = false;
  const walk = (n: Node) => {
    for (const c of n.childNodes) {
      if (c.nodeType === Node.TEXT_NODE) {
        s += (c as Text).data;
        if ((c as Text).data) lastBr = false;
      } else if (c instanceof HTMLElement) {
        if (c.tagName === "BR") {
          s += "\n";
          lastBr = true;
        } else if (c.classList.contains(tabClass)) {
          s += "\t";
          lastBr = false;
        } else walk(c);
      }
    }
  };
  walk(p);
  // Chromium keeps a placeholder <br> at the end of an emptied or line-broken block.
  return lastBr ? s.slice(0, -1) : s;
}

/** Applies paragraph text changes to the XML and returns the new DOCX bytes and XML. */
export async function buildDocx(src: DocxSource, baseXml: string, changes: Map<number, { from: string; to: string }>) {
  const doc = parseXml(baseXml);
  const paras = doc.getElementsByTagNameNS(W, "p");
  for (const [index, { from, to }] of changes) {
    const p = paras[index];
    const atoms = p && collectAtoms(p);
    if (!atoms || atoms.map((a) => a.text).join("") !== from) throw new Error("структура документа изменилась, правку не удалось применить");
    applyText(p, atoms, from, to);
  }
  const xml = serializeXml(doc, baseXml);
  src.zip.file(src.path, xml, NO_FOLDERS);
  const data = await src.zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return { data, xml };
}

// ---------------------------------------------------------------------------------------

function collectAtoms(p: Element): Atom[] | null {
  const atoms: Atom[] = [];
  let ok = true;
  const walk = (el: Element) => {
    for (const c of el.children) {
      if (!ok) return;
      if (c.namespaceURI === MC) {
        ok = false; // alternate content: unclear which branch is shown
        return;
      }
      if (c.namespaceURI !== W) continue;
      const name = c.localName;
      if (UNSUPPORTED.has(name)) {
        ok = false;
        return;
      }
      if (SKIP.has(name)) continue;
      if (name === "t") atoms.push({ el: c, kind: "t", text: c.textContent ?? "" });
      else if (name === "tab" && el.localName === "r") atoms.push({ el: c, kind: "tab", text: "\t" });
      else if (name === "br") {
        const type = c.getAttributeNS(W, "type") || c.getAttribute("w:type");
        if (type && type !== "textWrapping") {
          ok = false; // page / column break: docx-preview splits the paragraph
          return;
        }
        atoms.push({ el: c, kind: "br", text: "\n" });
      } else walk(c);
    }
  };
  walk(p);
  return ok ? atoms : null;
}

/** Applies each changed stretch separately, so untouched words keep their own runs. */
function applyText(p: Element, atoms: Atom[], from: string, to: string) {
  const hunks = diffWords(from, to);
  // Right to left: earlier offsets stay valid while later parts change.
  for (let k = hunks.length - 1; k >= 0; k--) {
    const h = hunks[k];
    applyHunk(p, k === hunks.length - 1 ? atoms : collectAtoms(p)!, h.at, h.at + h.del, h.insert);
  }
}

interface Hunk {
  at: number; // offset in the old text
  del: number;
  insert: string;
}

/** Word-level diff (LCS over words, spaces and punctuation), as character offsets. */
export function diffWords(from: string, to: string): Hunk[] {
  let pre = 0;
  const max = Math.min(from.length, to.length);
  while (pre < max && from[pre] === to[pre]) pre++;
  let suf = 0;
  while (suf < max - pre && from[from.length - 1 - suf] === to[to.length - 1 - suf]) suf++;
  if (pre === from.length && pre === to.length) return [];
  const a = tokenize(from.slice(pre, from.length - suf));
  const b = tokenize(to.slice(pre, to.length - suf));
  const whole = (): Hunk[] => [{ at: pre, del: from.length - suf - pre, insert: to.slice(pre, to.length - suf) }];
  if (!a.length || !b.length || a.length * b.length > 4_000_000) return whole();

  // lcs[i][j] = common tokens of a[i..] and b[j..]
  const m = b.length + 1;
  const lcs = new Uint32Array((a.length + 1) * m);
  for (let i = a.length - 1; i >= 0; i--)
    for (let j = b.length - 1; j >= 0; j--)
      lcs[i * m + j] = a[i] === b[j] ? lcs[(i + 1) * m + j + 1] + 1 : Math.max(lcs[(i + 1) * m + j], lcs[i * m + j + 1]);

  const hunks: Hunk[] = [];
  let pos = pre;
  let cur: Hunk | null = null;
  const flush = () => {
    if (cur && (cur.del || cur.insert)) hunks.push(cur);
    cur = null;
  };
  let i = 0;
  let j = 0;
  while (i < a.length || j < b.length) {
    if (i < a.length && j < b.length && a[i] === b[j]) {
      flush();
      pos += a[i].length;
      i++;
      j++;
    } else if (j < b.length && (i >= a.length || lcs[i * m + j + 1] >= lcs[(i + 1) * m + j])) {
      cur ??= { at: pos, del: 0, insert: "" };
      cur.insert += b[j++];
    } else {
      cur ??= { at: pos, del: 0, insert: "" };
      cur.del += a[i].length;
      pos += a[i++].length;
    }
  }
  flush();
  return hunks;
}

function tokenize(s: string): string[] {
  return s.match(/[\p{L}\p{N}_]+|\s+|[^\p{L}\p{N}_\s]/gu) ?? [];
}

function applyHunk(p: Element, atoms: Atom[], pre: number, delEnd: number, insert: string) {
  insert = insert.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");

  // Text inserted at a run boundary takes the formatting of the text before it, as in Word.
  let host = -1;
  let pos = 0;
  const starts = atoms.map((a) => {
    const s = pos;
    pos += a.text.length;
    return s;
  });
  for (let i = 0; i < atoms.length; i++) {
    const s = starts[i];
    const e = s + atoms[i].text.length;
    if (atoms[i].kind === "t" && s <= pre && pre <= e && (host < 0 || s < pre)) host = i;
  }

  for (let i = 0; i < atoms.length; i++) {
    const a = atoms[i];
    const s = starts[i];
    const e = s + a.text.length;
    if (i === host) continue;
    if (e <= pre || s >= delEnd) continue;
    if (a.kind !== "t") a.el.remove();
    else {
      const keep = a.text.slice(0, Math.max(0, pre - s)) + a.text.slice(Math.min(a.text.length, delEnd - s));
      if (keep) setText(a.el, keep);
      else a.el.remove();
    }
  }

  if (host >= 0) {
    const a = atoms[host];
    const s = starts[host];
    const left = a.text.slice(0, pre - s);
    const right = a.text.slice(Math.max(pre - s, Math.min(a.text.length, delEnd - s)));
    replaceWithTokens(a.el, left + insert + right);
    return;
  }
  if (!insert) return;
  // No text run at the insertion point: attach to a neighbouring tab / break, or start a run.
  const before = atoms.filter((a, i) => starts[i] + a.text.length <= pre && a.el.parentNode).pop();
  const after = atoms.find((a, i) => starts[i] >= delEnd && a.el.parentNode);
  const nodes = tokens(p.ownerDocument, p, insert);
  if (before) before.el.after(...nodes);
  else if (after) after.el.before(...nodes);
  else {
    const doc = p.ownerDocument;
    const run = doc.createElementNS(W, qname(p, "r"));
    const markProps = firstChild(firstChild(p, "pPr"), "rPr");
    if (markProps) {
      const rPr = markProps.cloneNode(true) as Element;
      for (const c of [...rPr.children]) if (["ins", "del", "moveFrom", "moveTo", "rPrChange"].includes(c.localName)) c.remove();
      run.appendChild(rPr);
    }
    run.append(...nodes);
    p.appendChild(run);
  }
}

function replaceWithTokens(el: Element, text: string) {
  const nodes = tokens(el.ownerDocument, el, text);
  if (!nodes.length) {
    el.remove();
    return;
  }
  // Keep the original <w:t> (and its attributes) for the first text piece.
  const firstText = nodes.findIndex((n) => n.localName === "t");
  if (firstText >= 0) {
    setText(el, nodes[firstText].textContent ?? "");
    nodes[firstText] = el;
  }
  el.replaceWith(...nodes);
}

function tokens(doc: Document, ref: Element, text: string): Element[] {
  const out: Element[] = [];
  for (const piece of text.split(/(\t|\n)/)) {
    if (!piece) continue;
    if (piece === "\t") out.push(doc.createElementNS(W, qname(ref, "tab")));
    else if (piece === "\n") out.push(doc.createElementNS(W, qname(ref, "br")));
    else {
      const t = doc.createElementNS(W, qname(ref, "t"));
      setText(t, piece);
      out.push(t);
    }
  }
  return out;
}

function setText(t: Element, text: string) {
  t.textContent = text;
  if (/^\s|\s$|\s\s/.test(text)) t.setAttributeNS(XML_NS, "xml:space", "preserve");
}

const qname = (ref: Element, local: string) => (ref.prefix ? `${ref.prefix}:${local}` : local);
const attrName = qname;

function firstChild(el: Element | null, local: string): Element | null {
  if (!el) return null;
  for (const c of el.children) if (c.namespaceURI === W && c.localName === local) return c;
  return null;
}

function insideAlternateContent(el: Element) {
  for (let n = el.parentElement; n; n = n.parentElement) if (n.namespaceURI === MC) return true;
  return false;
}

async function mainDocumentPath(zip: JSZip): Promise<string> {
  const rels = await zip.file("_rels/.rels")?.async("string");
  if (rels) {
    for (const r of parseXml(rels).getElementsByTagName("Relationship")) {
      if ((r.getAttribute("Type") ?? "").endsWith("/officeDocument")) {
        const target = (r.getAttribute("Target") ?? "").replace(/^\//, "");
        if (zip.file(target)) return target;
      }
    }
  }
  if (zip.file("word/document.xml")) return "word/document.xml";
  throw new Error("в архиве нет основного документа Word");
}

function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("не удалось разобрать XML документа");
  return doc;
}

/** XMLSerializer drops the XML declaration; Word expects it, so carry the original one over. */
function serializeXml(doc: Document, original: string): string {
  const decl = original.match(/^\s*<\?xml[^>]*\?>/)?.[0].trim() ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  return decl + "\r\n" + new XMLSerializer().serializeToString(doc.documentElement);
}
