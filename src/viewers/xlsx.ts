import ExcelJS from "exceljs";
import JSZip from "jszip";
import * as SSF from "ssf";
import { clampZoom, type Viewer, type ViewerContext } from "./types";
import { DomSearch } from "./domSearch";
import { findCharts, renderChart, type ChartPlacement } from "./xlsxCharts";

type WS = ExcelJS.Worksheet;
type Color = { argb?: string; theme?: number; tint?: number; indexed?: number };

// ---------------------------------------------------------------------------------------
// Colors
// ---------------------------------------------------------------------------------------

// Legacy indexed palette (BIFF8 default).
const INDEXED = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

const DEFAULT_THEME = [
  "FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72",
];

function parseTheme(xml: string | undefined): string[] {
  if (!xml) return DEFAULT_THEME;
  try {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const scheme = doc.getElementsByTagNameNS("*", "clrScheme")[0];
    if (!scheme) return DEFAULT_THEME;
    const get = (name: string) => {
      const el = scheme.getElementsByTagNameNS("*", name)[0];
      const c = el?.firstElementChild;
      if (!c) return null;
      return c.getAttribute("lastClr") || c.getAttribute("val");
    };
    const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    return order.map((n, i) => get(n) || DEFAULT_THEME[i]);
  } catch {
    return DEFAULT_THEME;
  }
}

function applyTint(hex: string, tint: number): string {
  let r = parseInt(hex.slice(0, 2), 16) / 255,
    g = parseInt(hex.slice(2, 4), 16) / 255,
    b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  let h = 0,
    s = 0,
    l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  if (s === 0) r = g = b = l;
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s,
      p = 2 * l - q;
    r = hue(p, q, h + 1 / 3);
    g = hue(p, q, h);
    b = hue(p, q, h - 1 / 3);
  }
  const x = (v: number) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, "0");
  return x(r) + x(g) + x(b);
}

// ---------------------------------------------------------------------------------------
// Values
// ---------------------------------------------------------------------------------------

type Kind = "n" | "s" | "b" | "e" | "empty";
interface Display {
  kind: Kind;
  text: string;
  rich?: { text: string; font?: Partial<ExcelJS.Font> }[];
  link?: string;
}

function toSerial(d: Date, date1904: boolean) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (d.getTime() - epoch) / 86400000;
}

// Excel shows numbers with the separators of the user's system locale, whatever the file says.
const LOCALE_SEP = (() => {
  try {
    const parts = new Intl.NumberFormat(navigator.language).formatToParts(12345.6);
    return {
      group: parts.find((p) => p.type === "group")?.value ?? ",",
      decimal: parts.find((p) => p.type === "decimal")?.value ?? ".",
    };
  } catch {
    return { group: ",", decimal: "." };
  }
})();

function normalizeFmt(fmt: string): string {
  // Format codes are case-insensitive in Excel but SSF expects lower-case date/time tokens.
  let f = fmt.replace(/"[^"]*"|\[[^\]]*\]|AM\/PM|A\/P|[DMYHS]/g, (m) => (m.length === 1 ? m.toLowerCase() : m));
  // SSF mistakes "dd.mm" for fractional seconds; escape literal dots in date formats.
  if (/^(?:[^"]|"[^"]*")*?[dy]/.test(f.replace(/\[[^\]]*\]/g, ""))) {
    f = f.replace(/"[^"]*"|\\.|\.(?!0)/g, (m) => (m === "." ? "\\." : m));
  }
  return f;
}

// SSF only knows English month/day names; swap them for the user's language.
const NAME_MAP: [RegExp, string][] = (() => {
  const lang = navigator.language || "ru";
  if (lang.startsWith("en")) return [];
  const out: [RegExp, string][] = [];
  try {
    const en = (o: Intl.DateTimeFormatOptions, d: Date) => new Intl.DateTimeFormat("en-US", o).format(d);
    const loc = (o: Intl.DateTimeFormatOptions, d: Date) => new Intl.DateTimeFormat(lang, o).format(d);
    const gen = (d: Date) =>
      new Intl.DateTimeFormat(lang, { day: "numeric", month: "long" }).formatToParts(d).find((p) => p.type === "month")?.value;
    for (let m = 0; m < 12; m++) {
      const d = new Date(2026, m, 15);
      out.push([new RegExp(`\\b${en({ month: "long" }, d)}\\b`, "g"), gen(d) ?? loc({ month: "long" }, d)]);
      out.push([new RegExp(`\\b${en({ month: "short" }, d)}\\b`, "g"), loc({ month: "short" }, d).replace(/\.$/, "")]);
    }
    for (let w = 0; w < 7; w++) {
      const d = new Date(2026, 0, 4 + w);
      out.push([new RegExp(`\\b${en({ weekday: "long" }, d)}\\b`, "g"), loc({ weekday: "long" }, d)]);
      out.push([new RegExp(`\\b${en({ weekday: "short" }, d)}\\b`, "g"), loc({ weekday: "short" }, d)]);
    }
  } catch {
    /* keep English */
  }
  return out;
})();

function localizeNames(s: string) {
  if (!/[A-Z][a-z]/.test(s)) return s;
  for (const [re, name] of NAME_MAP) s = s.replace(re, name);
  return s;
}

function localizeNumber(s: string): string {
  if (LOCALE_SEP.group === "," && LOCALE_SEP.decimal === ".") return s;
  const group = LOCALE_SEP.group.replace(/\u00A0/, "\u202F");
  return s.replace(/\d[\d,]*(?:\.\d+)?/g, (num) =>
    num.replace(/[,.]/g, (ch) => (ch === "," ? group : LOCALE_SEP.decimal)),
  );
}

// Built-in formats 14 and 22 mean "system short date (and time)" in Excel.
const SHORT_DATE = new Intl.DateTimeFormat(navigator.language, { day: "2-digit", month: "2-digit", year: "numeric", timeZone: "UTC" });
const SHORT_DATE_TIME = new Intl.DateTimeFormat(navigator.language, {
  day: "2-digit",
  month: "2-digit",
  year: "numeric",
  hour: "2-digit",
  minute: "2-digit",
  timeZone: "UTC",
});
const serialToDate = (v: number) => new Date(Math.round((v - 25569) * 86400000));

function fmtNumber(v: number, fmt: string | undefined): string {
  if (fmt === "mm-dd-yy" || fmt === "m/d/yy" || fmt === "14") return SHORT_DATE.format(serialToDate(v));
  if (fmt === "m/d/yy h:mm" || fmt === "22") return SHORT_DATE_TIME.format(serialToDate(v));
  const f = normalizeFmt(fmt || "General");
  try {
    const out = SSF.format(f, v);
    return SSF.is_date(f) ? localizeNames(out) : localizeNumber(out);
  } catch {
    return localizeNumber(String(v));
  }
}

function display(cell: ExcelJS.Cell, date1904: boolean): Display {
  let v: any = cell.value;
  let link: string | undefined;
  if (v === null || v === undefined) return { kind: "empty", text: "" };
  if (typeof v === "object" && !(v instanceof Date)) {
    if ("hyperlink" in v) {
      link = v.hyperlink;
      v = v.text;
    }
    if (v && typeof v === "object" && ("formula" in v || "sharedFormula" in v)) v = v.result;
    if (v && typeof v === "object" && "richText" in v) {
      const rich = (v.richText as ExcelJS.RichText[]).map((r) => ({ text: r.text, font: r.font }));
      return { kind: "s", text: rich.map((r) => r.text).join(""), rich, link };
    }
    if (v && typeof v === "object" && "error" in v) return { kind: "e", text: String(v.error), link };
  }
  if (v === null || v === undefined) return { kind: "empty", text: "" };
  if (v instanceof Date) {
    const fmt = cell.numFmt && !/^general$/i.test(cell.numFmt) ? cell.numFmt : "mm-dd-yy";
    if (fmt === "mm-dd-yy" || fmt === "m/d/yy") return { kind: "n", text: SHORT_DATE.format(v), link };
    return { kind: "n", text: fmtNumber(toSerial(v, date1904), fmt), link };
  }
  if (typeof v === "number") return { kind: "n", text: fmtNumber(v, cell.numFmt), link };
  if (typeof v === "boolean") return { kind: "b", text: v ? "ИСТИНА" : "ЛОЖЬ", link };
  if (typeof v === "object") return { kind: "s", text: JSON.stringify(v), link };
  const s = String(v);
  // Text format with "@" or literal sections, e.g. '"Итого: "@'
  if (cell.numFmt && cell.numFmt.includes("@") && cell.numFmt !== "@") {
    try {
      return { kind: "s", text: SSF.format(cell.numFmt, s), link };
    } catch {
      /* fallthrough */
    }
  }
  return { kind: "s", text: s, link };
}

function colName(c: number) {
  let s = "";
  for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
}

const esc = (s: string) => s.replace(/[&<>"]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[ch]!);

// ---------------------------------------------------------------------------------------
// Sheet rendering
// ---------------------------------------------------------------------------------------

const BORDER: Record<string, string> = {
  thin: "1px solid",
  hair: "1px solid",
  dotted: "1px dotted",
  dashed: "1px dashed",
  dashDot: "1px dashed",
  dashDotDot: "1px dashed",
  medium: "2px solid",
  mediumDashed: "2px dashed",
  mediumDashDot: "2px dashed",
  mediumDashDotDot: "2px dashed",
  slantDashDot: "2px dashed",
  thick: "3px solid",
  double: "3px double",
};

const HEADER_W = 46;
const HEADER_H = 22;
const CHUNK = 400;

interface SheetView {
  el: HTMLElement;
  renderAll(): void;
  destroy(): void;
}

function renderSheet(
  ws: WS,
  wb: ExcelJS.Workbook,
  theme: string[],
  date1904: boolean,
  host: HTMLElement,
  onCell: (addr: string, cell: ExcelJS.Cell) => void,
  charts: ChartPlacement[] = [],
  zip?: JSZip,
): SheetView {
  const color = (c?: Partial<Color>): string | undefined => {
    if (!c) return undefined;
    if (c.argb && c.argb.length >= 6) return "#" + c.argb.slice(-6);
    if (c.theme !== undefined) {
      const base = theme[c.theme] ?? "000000";
      return "#" + (c.tint ? applyTint(base, c.tint) : base);
    }
    if (c.indexed !== undefined && c.indexed < 64) return "#" + INDEXED[c.indexed];
    return undefined;
  };

  // ---- used range -------------------------------------------------------------------
  let maxRow = 0,
    maxCol = 0;
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      if (cell.value !== null && cell.value !== undefined && cell.value !== "") {
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      } else if (cell.fill && (cell.fill as any).pattern && (cell.fill as any).pattern !== "none") {
        // Coloured but empty cells still belong to the visible design (up to a limit).
        if (r > maxRow && r < maxRow + 200) maxRow = r;
        if (c > maxCol && c < maxCol + 50) maxCol = c;
      }
    });
  });
  const merges: { top: number; left: number; bottom: number; right: number }[] = Object.values(
    (ws as any)._merges ?? {},
  ).map((m: any) => ({ top: m.top, left: m.left, bottom: m.bottom, right: m.right }));
  for (const m of merges) {
    maxRow = Math.max(maxRow, m.bottom);
    maxCol = Math.max(maxCol, m.right);
  }
  const images = ws.getImages?.() ?? [];
  for (const im of images) {
    const br: any = (im.range as any).br ?? (im.range as any).tl;
    maxRow = Math.max(maxRow, Math.ceil((br?.nativeRow ?? 0) + 1));
    maxCol = Math.max(maxCol, Math.ceil((br?.nativeCol ?? 0) + 1));
  }
  for (const ch of charts) {
    const end = ch.to ?? ch.from;
    maxRow = Math.max(maxRow, end.row + 1);
    maxCol = Math.max(maxCol, end.col + 1);
  }
  const rows = Math.max(maxRow + 5, 40);
  const cols = Math.max(maxCol + 2, 12);

  // ---- geometry ---------------------------------------------------------------------
  const props: any = ws.properties ?? {};
  const defColW = props.defaultColWidth ?? 8.43;
  const defRowH = props.defaultRowHeight ?? 15;
  const colPx: number[] = [0];
  const colHidden: boolean[] = [false];
  for (let c = 1; c <= cols; c++) {
    const col = ws.getColumn(c);
    colHidden[c] = !!col.hidden;
    colPx[c] = Math.max(8, Math.round((col.width ?? defColW) * 7 + 5));
  }
  const rowPx = (r: number) => {
    const row = ws.findRow(r);
    return Math.round(((row?.height as number | undefined) ?? defRowH) * (96 / 72));
  };
  const rowHidden = (r: number) => !!ws.findRow(r)?.hidden;

  const view: any = ws.views?.[0] ?? {};
  const showGrid = view.showGridLines !== false;
  const frozenRows = view.state === "frozen" ? view.ySplit ?? 0 : 0;
  const frozenCols = view.state === "frozen" ? view.xSplit ?? 0 : 0;

  // Merge lookup
  const mergeAt = new Map<string, { rs: number; cs: number }>();
  const covered = new Set<string>();
  for (const m of merges) {
    let rs = 0,
      cs = 0;
    for (let r = m.top; r <= m.bottom; r++) if (!rowHidden(r)) rs++;
    for (let c = m.left; c <= m.right; c++) if (!colHidden[c]) cs++;
    mergeAt.set(`${m.top}:${m.left}`, { rs: Math.max(1, rs), cs: Math.max(1, cs) });
    for (let r = m.top; r <= m.bottom; r++)
      for (let c = m.left; c <= m.right; c++) if (r !== m.top || c !== m.left) covered.add(`${r}:${c}`);
  }
  const mergeOf = (r: number, c: number) =>
    merges.find((m) => r >= m.top && r <= m.bottom && c >= m.left && c <= m.right);

  // Sticky offsets for frozen panes
  const stickyLeft: number[] = [];
  {
    let x = HEADER_W;
    for (let c = 1; c <= frozenCols; c++) {
      stickyLeft[c] = x;
      if (!colHidden[c]) x += colPx[c];
    }
  }
  const stickyTop: number[] = [];
  {
    let y = HEADER_H;
    for (let r = 1; r <= frozenRows; r++) {
      stickyTop[r] = y;
      if (!rowHidden(r)) y += rowPx(r);
    }
  }

  // ---- DOM --------------------------------------------------------------------------
  const wrap = document.createElement("div");
  wrap.className = "xl-sheet" + (showGrid ? "" : " no-grid");
  const table = document.createElement("table");
  table.className = "xl-grid";
  const colgroup = document.createElement("colgroup");
  let totalW = HEADER_W;
  colgroup.innerHTML =
    `<col style="width:${HEADER_W}px">` +
    Array.from({ length: cols }, (_, i) => i + 1)
      .filter((c) => !colHidden[c])
      .map((c) => {
        totalW += colPx[c];
        return `<col style="width:${colPx[c]}px">`;
      })
      .join("");
  table.style.width = `${totalW}px`;
  const thead = document.createElement("thead");
  thead.innerHTML =
    `<tr style="height:${HEADER_H}px"><th class="xl-corner"></th>` +
    Array.from({ length: cols }, (_, i) => i + 1)
      .filter((c) => !colHidden[c])
      .map((c) => `<th class="xl-ch" data-c="${c}"${c <= frozenCols ? ` style="left:${stickyLeft[c]}px;z-index:4"` : ""}>${colName(c)}</th>`)
      .join("") +
    `</tr>`;
  table.append(colgroup, thead);
  wrap.appendChild(table);

  const borderCss = (b?: Partial<ExcelJS.Border>) => (b?.style ? `${BORDER[b.style] ?? "1px solid"} ${color(b.color) ?? "#000"}` : "");

  const isEmpty = (r: number, c: number) => {
    if (covered.has(`${r}:${c}`) || mergeAt.has(`${r}:${c}`)) return false;
    const row = ws.findRow(r);
    if (!row) return true;
    const v = row.findCell(c)?.value;
    return v === null || v === undefined || v === "";
  };

  function cellHtml(r: number, c: number): string {
    const key = `${r}:${c}`;
    if (covered.has(key)) return "";
    const row = ws.findRow(r);
    // Don't materialise empty cells in the model; a stub is enough for rendering.
    const cell: ExcelJS.Cell = row?.findCell(c) ?? ({ value: null } as unknown as ExcelJS.Cell);
    const m = mergeAt.get(key);
    const span = m ? `${m.rs > 1 ? ` rowspan="${m.rs}"` : ""}${m.cs > 1 ? ` colspan="${m.cs}"` : ""}` : "";
    const st: string[] = [];
    const d = display(cell, date1904);

    const fill: any = cell.fill;
    let bg: string | undefined;
    if (fill?.type === "pattern" && fill.pattern && fill.pattern !== "none") {
      bg = color(fill.pattern === "solid" ? fill.fgColor : fill.fgColor ?? fill.bgColor);
    } else if (fill?.type === "gradient" && fill.stops?.length) {
      bg = color(fill.stops[0].color);
    }
    if (bg) st.push(`background:${bg}`);

    const font: Partial<ExcelJS.Font> = cell.font ?? {};
    const fontCss = (f: Partial<ExcelJS.Font>) => {
      const o: string[] = [];
      if (f.bold) o.push("font-weight:700");
      if (f.italic) o.push("font-style:italic");
      const deco = [f.underline ? "underline" : "", f.strike ? "line-through" : ""].filter(Boolean).join(" ");
      if (deco) o.push(`text-decoration:${deco}`);
      if (f.size && f.size !== 11) o.push(`font-size:${((f.size * 96) / 72).toFixed(1)}px`);
      const fc = color(f.color as Color);
      if (fc && fc.toUpperCase() !== "#000000") o.push(`color:${fc}`);
      if (f.name && !/^(calibri|aptos( narrow)?)$/i.test(f.name)) o.push(`font-family:'${f.name.replace(/'/g, "")}',var(--xl-font)`);
      return o;
    };
    st.push(...fontCss(font));

    // Borders: this cell's right/bottom, plus neighbours' left/top pushed onto us.
    const b: Partial<ExcelJS.Borders> = cell.border ?? {};
    const mm = m ? mergeOf(r, c) : undefined;
    const rightCell = mm ? ws.findRow(mm.top)?.findCell(mm.right) : cell;
    const bottomCell = mm ? ws.findRow(mm.bottom)?.findCell(mm.left) : cell;
    const nextC = (mm?.right ?? c) + 1,
      nextR = (mm?.bottom ?? r) + 1;
    const right = borderCss(rightCell?.border?.right) || borderCss(ws.findRow(r)?.findCell(nextC)?.border?.left);
    const bottom = borderCss(bottomCell?.border?.bottom) || borderCss(ws.findRow(nextR)?.findCell(c)?.border?.top);
    if (right) st.push(`border-right:${right}`);
    if (bottom) st.push(`border-bottom:${bottom}`);
    if (r === 1 && b.top?.style) st.push(`border-top:${borderCss(b.top)}`);
    if (c === 1 && b.left?.style) st.push(`border-left:${borderCss(b.left)}`);
    if (bg && !right && showGrid) st.push(`border-right-color:${bg}`);
    if (bg && !bottom && showGrid) st.push(`border-bottom-color:${bg}`);

    const al: Partial<ExcelJS.Alignment> = cell.alignment ?? {};
    let h = al.horizontal as string | undefined;
    if (!h || h === "general") h = d.kind === "n" ? "right" : d.kind === "b" || d.kind === "e" ? "center" : "left";
    if (h === "centerContinuous") h = "center";
    if (h === "fill" || h === "distributed" || h === "justify") h = "left";
    st.push(`text-align:${h}`);
    const v = al.vertical === "middle" ? "middle" : al.vertical === "top" ? "top" : "bottom";
    if (v !== "bottom") st.push(`vertical-align:${v}`);
    if (al.indent) st.push(`padding-left:${3 + al.indent * 9}px`);

    if (c <= frozenCols || r <= frozenRows) {
      st.push("position:sticky");
      if (c <= frozenCols) st.push(`left:${stickyLeft[c]}px`);
      if (r <= frozenRows) st.push(`top:${stickyTop[r]}px`);
      st.push(`z-index:${c <= frozenCols && r <= frozenRows ? 3 : 2}`);
      if (!bg) st.push("background:var(--xl-cell-bg)");
    }

    const wrapText = !!al.wrapText;
    let inner = "";
    if (d.rich) {
      inner = d.rich
        .map((run) => {
          const css = run.font ? fontCss(run.font).join(";") : "";
          const t = esc(run.text).replace(/\n/g, wrapText ? "<br>" : " ");
          return css ? `<span style="${css}">${t}</span>` : t;
        })
        .join("");
    } else {
      inner = esc(d.text);
      inner = wrapText ? inner.replace(/\n/g, "<br>") : inner.replace(/\n/g, " ");
    }
    if (d.link) inner = `<span class="xl-link">${inner}</span>`;

    // Excel lets non-wrapped left-aligned text spill over empty neighbours.
    let cls = "";
    if (d.kind === "s" && !wrapText && !m && h === "left" && isEmpty(r, c + 1)) cls = "spill";
    else if (wrapText) cls = "wrap";
    if (d.kind === "n") cls = "num";

    const rot = (al as any).textRotation;
    if (rot && rot !== 0) {
      const deg = rot === "vertical" ? 90 : rot > 90 ? rot - 90 : -rot;
      inner = `<span style="display:inline-block;transform:rotate(${deg}deg)">${inner}</span>`;
    }

    let maxH = "";
    if (wrapText) {
      let hpx = 0;
      const rs = mm ? mm.bottom - mm.top + 1 : 1;
      for (let k = 0; k < rs; k++) if (!rowHidden(r + k)) hpx += rowPx(r + k);
      maxH = ` style="max-height:${hpx - 1}px"`;
    }
    const content = inner ? `<div class="c ${cls}"${maxH}>${inner}</div>` : "";
    return `<td data-r="${r}" data-c="${c}"${span}${st.length ? ` style="${st.join(";")}"` : ""}>${content}</td>`;
  }

  function rowHtml(r: number): string {
    if (rowHidden(r)) return "";
    const h = rowPx(r);
    const sticky = r <= frozenRows ? ` style="position:sticky;top:${stickyTop[r]}px;z-index:3"` : "";
    let s = `<tr style="height:${h}px"><th class="xl-rh"${sticky}>${r}</th>`;
    for (let c = 1; c <= cols; c++) if (!colHidden[c]) s += cellHtml(r, c);
    return s + "</tr>";
  }

  let rendered = 0;
  function renderMore(n = CHUNK) {
    if (rendered >= rows) return;
    const end = Math.min(rows, rendered + n);
    let html = "";
    for (let r = rendered + 1; r <= end; r++) html += rowHtml(r);
    const tb = document.createElement("tbody");
    tb.innerHTML = html;
    table.appendChild(tb);
    rendered = end;
    if (wrap.isConnected) markOverflowingNumbers(tb);
    else requestAnimationFrame(() => markOverflowingNumbers(tb));
  }

  // Like Excel: a number that doesn't fit its column is shown as ###.
  function markOverflowingNumbers(tb: HTMLElement) {
    const cells = [...tb.querySelectorAll<HTMLElement>(".c.num")];
    const over = cells.filter((c) => c.scrollWidth > c.clientWidth + 1);
    for (const c of over) {
      c.title = c.textContent ?? "";
      const n = Math.max(1, Math.floor(c.clientWidth / 8.2));
      c.textContent = "#".repeat(n);
    }
  }
  renderMore();

  const onScroll = () => {
    if (rendered < rows && host.scrollTop + host.clientHeight * 3 > host.scrollHeight) renderMore();
  };
  host.addEventListener("scroll", onScroll, { passive: true });

  // ---- images -----------------------------------------------------------------------
  const urls: string[] = [];
  const layer = document.createElement("div");
  layer.className = "xl-images";
  wrap.appendChild(layer);
  const colLeft = (c0: number) => {
    // c0 is 0-based native column
    let x = HEADER_W;
    for (let c = 1; c <= c0; c++) if (!colHidden[c]) x += colPx[c] ?? Math.round(defColW * 7 + 5);
    return x;
  };
  const rowTop = (r0: number) => {
    let y = HEADER_H;
    for (let r = 1; r <= r0; r++) if (!rowHidden(r)) y += rowPx(r);
    return y;
  };
  for (const im of images) {
    try {
      const media: any = wb.getImage(Number(im.imageId));
      if (!media) continue;
      const bytes: Uint8Array | undefined = media.buffer ?? (media.base64 ? Uint8Array.from(atob(media.base64), (ch) => ch.charCodeAt(0)) : undefined);
      if (!bytes) continue;
      const ext = (media.extension || "png").replace("jpg", "jpeg");
      const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: `image/${ext}` }));
      urls.push(url);
      const rg: any = im.range;
      const tl = rg.tl,
        br = rg.br;
      const x = colLeft(tl.nativeCol) + (tl.nativeColOff ?? 0) / 9525;
      const y = rowTop(tl.nativeRow) + (tl.nativeRowOff ?? 0) / 9525;
      let w: number, h: number;
      if (br) {
        w = colLeft(br.nativeCol) + (br.nativeColOff ?? 0) / 9525 - x;
        h = rowTop(br.nativeRow) + (br.nativeRowOff ?? 0) / 9525 - y;
      } else {
        w = rg.ext?.width ?? 100;
        h = rg.ext?.height ?? 100;
      }
      const img = document.createElement("img");
      img.src = url;
      img.style.cssText = `left:${x}px;top:${y}px;width:${Math.max(1, w)}px;height:${Math.max(1, h)}px`;
      layer.appendChild(img);
    } catch (e) {
      console.warn("image", e);
    }
  }

  // ---- charts -----------------------------------------------------------------------
  let destroyed = false;
  const disposers: (() => void)[] = [];
  for (const ch of zip ? charts : []) {
    const x = colLeft(ch.from.col) + ch.from.colOff / 9525;
    const y = rowTop(ch.from.row) + ch.from.rowOff / 9525;
    const w = ch.to ? colLeft(ch.to.col) + ch.to.colOff / 9525 - x : (ch.ext?.cx ?? 4572000) / 9525;
    const h = ch.to ? rowTop(ch.to.row) + ch.to.rowOff / 9525 - y : (ch.ext?.cy ?? 2743200) / 9525;
    if (w < 4 || h < 4) continue;
    const box = document.createElement("div");
    box.className = "xl-chart";
    box.style.cssText = `left:${x}px;top:${y}px;width:${w}px;height:${h}px`;
    layer.appendChild(box);
    renderChart(zip!, ch.chartPath, w, h, box)
      .then((dispose) => (destroyed ? dispose() : disposers.push(dispose)))
      .catch((e) => {
        console.warn("chart", e);
        box.classList.add("failed");
        box.textContent = "Диаграмма";
      });
  }

  // ---- selection / cell info ---------------------------------------------------------
  let selected: HTMLElement | null = null;
  table.addEventListener("click", (e) => {
    const td = (e.target as HTMLElement).closest("td") as HTMLElement | null;
    if (!td) return;
    selected?.classList.remove("sel");
    selected = td;
    td.classList.add("sel");
    const r = Number(td.dataset.r),
      c = Number(td.dataset.c);
    onCell(colName(c) + r, ws.getRow(r).getCell(c));
  });

  return {
    el: wrap,
    renderAll() {
      if (rendered < rows) renderMore(rows - rendered);
    },
    destroy() {
      destroyed = true;
      disposers.forEach((d) => d());
      host.removeEventListener("scroll", onScroll);
      urls.forEach((u) => URL.revokeObjectURL(u));
      wrap.remove();
    },
  };
}

// ---------------------------------------------------------------------------------------
// CSV
// ---------------------------------------------------------------------------------------

function decodeText(data: ArrayBuffer): string {
  const bytes = new Uint8Array(data);
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes).replace(/^﻿/, "");
  } catch {
    // Cyrillic CSVs exported by Excel are very often in Windows-1251.
    return new TextDecoder("windows-1251").decode(bytes);
  }
}

function parseCsv(text: string, sep?: string): string[][] {
  if (!sep) {
    const head = text.slice(0, 5000);
    const counts = [",", ";", "\t", "|"].map((d) => [d, head.split(d).length] as const);
    sep = counts.sort((a, b) => b[1] - a[1])[0][0];
  }
  const rows: string[][] = [];
  let row: string[] = [],
    field = "",
    q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else q = false;
      } else field += ch;
    } else if (ch === '"' && field === "") q = true;
    else if (ch === sep) {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += ch;
  }
  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function csvWorkbook(data: ArrayBuffer, name: string): ExcelJS.Workbook {
  const rows = parseCsv(decodeText(data), name.toLowerCase().endsWith(".tsv") ? "\t" : undefined);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(name.replace(/\.[^.]+$/, "").slice(0, 31) || "CSV");
  const widths: number[] = [];
  rows.forEach((r) => {
    ws.addRow(
      r.map((v, i) => {
        widths[i] = Math.max(widths[i] ?? 6, Math.min(60, Math.ceil(v.length * 1.2) + 2));
        const t = v.trim();
        return /^-?\d+(\.\d+)?$/.test(t) && t.length < 16 ? Number(t) : v;
      }),
    );
  });
  widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));
  if (rows.length > 1) {
    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: "frozen", ySplit: 1, xSplit: 0 }];
  }
  return wb;
}

// ---------------------------------------------------------------------------------------
// Loading (with repairs for files ExcelJS chokes on)
// ---------------------------------------------------------------------------------------

async function loadWorkbook(data: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const attempt = async (buf: ArrayBuffer) => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    return wb;
  };
  let firstError: unknown;
  try {
    return await attempt(data);
  } catch (e) {
    firstError = e;
    console.warn("ExcelJS failed, trying repairs:", e);
  }
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw firstError;
  }
  const files = (re: RegExp) => Object.keys(zip.files).filter((n) => re.test(n));
  const edit = async (name: string, fn: (xml: string) => string) => {
    const xml = await zip.file(name)!.async("string");
    const out = fn(xml);
    if (out !== xml) zip.file(name, out);
  };
  const sheets = files(/^xl\/worksheets\/[^/]+\.xml$/);

  // Each step makes the file more "boring"; we stop at the first one ExcelJS accepts.
  const repairs: (() => Promise<void>)[] = [
    // Drawings written without the "xdr:" prefix (openpyxl, some online converters).
    async () => {
      for (const n of files(/^xl\/drawings\/[^/]+\.xml$/))
        await edit(n, (xml) =>
          /<wsDr[\s>]/.test(xml)
            ? xml
                .replace(/xmlns="(http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/spreadsheetDrawing)"/, 'xmlns:xdr="$1"')
                .replace(/<(\/?)(?![\w-]+:)([A-Za-z][\w-]*)/g, "<$1xdr:$2")
            : xml,
        );
    },
    // Tables (Ctrl+T ranges) with filters ExcelJS doesn't understand; cell data is unaffected.
    async () => {
      for (const n of files(/^xl\/tables\//)) zip.remove(n);
      for (const n of sheets) await edit(n, (xml) => xml.replace(/<(?:\w+:)?tableParts\b[\s\S]*?<\/(?:\w+:)?tableParts>|<(?:\w+:)?tableParts\b[^>]*\/>/g, ""));
    },
    // Conditional formatting, data validation and extension blocks.
    async () => {
      for (const n of sheets)
        await edit(n, (xml) =>
          xml
            .replace(/<(?:\w+:)?conditionalFormatting\b[\s\S]*?<\/(?:\w+:)?conditionalFormatting>/g, "")
            .replace(/<(?:\w+:)?dataValidations\b[\s\S]*?<\/(?:\w+:)?dataValidations>/g, "")
            .replace(/<(?:\w+:)?extLst\b[\s\S]*?<\/(?:\w+:)?extLst>/g, ""),
        );
    },
    // Drawings (images, charts, comments) — last resort, the data still shows.
    async () => {
      for (const n of files(/^xl\/(drawings|comments)/)) zip.remove(n);
      for (const n of sheets)
        await edit(n, (xml) => xml.replace(/<(?:\w+:)?(?:legacyDrawing|drawing)\b[^>]*\/>/g, ""));
    },
  ];
  for (const repair of repairs) {
    await repair();
    try {
      return await attempt(await zip.generateAsync({ type: "arraybuffer" }));
    } catch (e) {
      console.warn("repair step failed:", e);
    }
  }
  throw firstError;
}

// ---------------------------------------------------------------------------------------
// Viewer
// ---------------------------------------------------------------------------------------

export async function openSpreadsheet(data: ArrayBuffer, ctx: ViewerContext, fileName: string): Promise<Viewer> {
  const { host, footer } = ctx;
  let wb: ExcelJS.Workbook;
  let zip: JSZip | undefined;
  let charts = new Map<string, ChartPlacement[]>();
  if (/\.(csv|tsv)$/i.test(fileName)) {
    wb = csvWorkbook(data, fileName);
  } else {
    wb = await loadWorkbook(data);
    try {
      zip = await JSZip.loadAsync(data);
      charts = await findCharts(zip);
    } catch (e) {
      console.warn("charts", e);
    }
  }
  const theme = parseTheme((wb as any)._themes?.theme1);
  const date1904 = !!(wb.properties as any)?.date1904;
  const sheets = wb.worksheets.filter((ws) => ws.state !== "hidden" && ws.state !== "veryHidden");
  if (!sheets.length) throw new Error("В книге нет видимых листов");

  const root = document.createElement("div");
  root.className = "xl-doc";
  host.appendChild(root);

  // Footer: sheet tabs + selected cell info
  const tabs = document.createElement("div");
  tabs.className = "sheet-tabs";
  const info = document.createElement("div");
  info.className = "cell-info";
  footer.replaceChildren(tabs, info);

  let zoom = 1;
  let active = -1;
  let view: SheetView | null = null;
  const scrollPos = new Map<number, [number, number]>();

  const showCell = (addr: string, cell: ExcelJS.Cell) => {
    const v: any = cell.value;
    let raw = "";
    if (v && typeof v === "object" && "formula" in v) raw = "=" + v.formula;
    else if (v && typeof v === "object" && "sharedFormula" in v) raw = "=" + ((cell as any).formula ?? v.sharedFormula);
    else raw = display(cell, date1904).text;
    info.innerHTML = `<b>${addr}</b><span>${esc(raw)}</span>`;
    info.title = raw;
  };

  function select(i: number) {
    if (i === active) return;
    if (view && active >= 0) scrollPos.set(active, [host.scrollLeft, host.scrollTop]);
    view?.destroy();
    active = i;
    view = renderSheet(sheets[i], wb, theme, date1904, host, showCell, charts.get(sheets[i].name), zip);
    view.el.style.zoom = String(zoom);
    root.replaceChildren(view.el);
    const [x, y] = scrollPos.get(i) ?? [0, 0];
    host.scrollTo(x, y);
    info.textContent = "";
    [...tabs.children].forEach((t, k) => t.classList.toggle("active", k === i));
    search.clear();
    ctx.changed();
  }

  sheets.forEach((ws, i) => {
    const b = document.createElement("button");
    b.className = "sheet-tab";
    b.textContent = ws.name;
    const tc: any = (ws.properties as any)?.tabColor;
    if (tc?.argb) b.style.setProperty("--tab-color", "#" + tc.argb.slice(-6));
    b.onclick = () => select(i);
    tabs.appendChild(b);
  });

  const search = new DomSearch(() => root, host, () => view?.renderAll());
  select(0);

  return {
    get zoom() {
      return zoom;
    },
    setZoom(z) {
      zoom = clampZoom(z);
      if (view) view.el.style.zoom = String(zoom);
      ctx.changed();
    },
    search,
    summary: () => `${sheets.length} ${plural(sheets.length, "лист", "листа", "листов")}`,
    destroy() {
      view?.destroy();
      root.remove();
      footer.replaceChildren();
    },
  };
}

function plural(n: number, one: string, few: string, many: string) {
  const m10 = n % 10,
    m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return one;
  if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few;
  return many;
}
