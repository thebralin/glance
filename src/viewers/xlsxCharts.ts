// Excel charts. ExcelJS ignores them, but the chart XML (DrawingML "c:chartSpace") is the same
// format PowerPoint uses, so we wrap each chart into a tiny one-slide PPTX and let the PPTX
// renderer (which draws charts with ECharts) render it.
import JSZip from "jszip";

export interface CellAnchor {
  col: number; // 0-based
  colOff: number; // EMU
  row: number;
  rowOff: number;
}

export interface ChartPlacement {
  from: CellAnchor;
  to?: CellAnchor;
  ext?: { cx: number; cy: number };
  chartPath: string; // zip path, e.g. xl/charts/chart1.xml
}

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function parseXml(xml: string) {
  return new DOMParser().parseFromString(xml, "application/xml");
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

async function readRels(zip: JSZip, part: string): Promise<Map<string, string>> {
  const relsPath = part.replace(/([^/]+)$/, "_rels/$1.rels");
  const file = zip.file(relsPath);
  const map = new Map<string, string>();
  if (!file) return map;
  const doc = parseXml(await file.async("string"));
  for (const r of Array.from(doc.getElementsByTagName("Relationship"))) {
    if (r.getAttribute("TargetMode") === "External") continue;
    map.set(r.getAttribute("Id")!, resolvePath(part, r.getAttribute("Target")!));
  }
  return map;
}

const local = (el: Element, name: string) => Array.from(el.children).find((c) => c.localName === name);
const num = (el: Element | undefined, name: string) => Number(el && local(el, name)?.textContent) || 0;

function anchorOf(el: Element | undefined): CellAnchor | undefined {
  if (!el) return undefined;
  return { col: num(el, "col"), colOff: num(el, "colOff"), row: num(el, "row"), rowOff: num(el, "rowOff") };
}

/** Map of sheet name -> charts placed on it. */
export async function findCharts(zip: JSZip): Promise<Map<string, ChartPlacement[]>> {
  const result = new Map<string, ChartPlacement[]>();
  const wbFile = zip.file("xl/workbook.xml");
  if (!wbFile || !Object.keys(zip.files).some((n) => n.startsWith("xl/charts/"))) return result;
  const wbRels = await readRels(zip, "xl/workbook.xml");
  const wb = parseXml(await wbFile.async("string"));
  for (const sheet of Array.from(wb.getElementsByTagNameNS("*", "sheet"))) {
    const name = sheet.getAttribute("name") ?? "";
    const rid = sheet.getAttributeNS(REL_NS, "id") ?? sheet.getAttribute("r:id");
    const sheetPath = rid ? wbRels.get(rid) : undefined;
    if (!sheetPath || !zip.file(sheetPath)) continue;
    const sheetRels = await readRels(zip, sheetPath);
    const sheetDoc = parseXml(await zip.file(sheetPath)!.async("string"));
    const drawingEl = sheetDoc.getElementsByTagNameNS("*", "drawing")[0];
    const drawingRid = drawingEl?.getAttributeNS(REL_NS, "id") ?? drawingEl?.getAttribute("r:id");
    const drawingPath = drawingRid ? sheetRels.get(drawingRid) : undefined;
    if (!drawingPath || !zip.file(drawingPath)) continue;
    const drawingRels = await readRels(zip, drawingPath);
    const drawing = parseXml(await zip.file(drawingPath)!.async("string"));
    const charts: ChartPlacement[] = [];
    for (const anchor of Array.from(drawing.documentElement.children)) {
      if (!/Anchor$/.test(anchor.localName)) continue;
      const chartEl = anchor.getElementsByTagNameNS("*", "chart")[0];
      const crid = chartEl?.getAttributeNS(REL_NS, "id") ?? chartEl?.getAttribute("r:id");
      const chartPath = crid ? drawingRels.get(crid) : undefined;
      if (!chartPath || !zip.file(chartPath)) continue;
      const ext = local(anchor, "ext");
      charts.push({
        from: anchorOf(local(anchor, "from")) ?? { col: 0, colOff: 0, row: 0, rowOff: 0 },
        to: anchorOf(local(anchor, "to")),
        ext: ext ? { cx: Number(ext.getAttribute("cx")) || 0, cy: Number(ext.getAttribute("cy")) || 0 } : undefined,
        chartPath,
      });
    }
    if (charts.length) result.set(name, charts);
  }
  return result;
}

// ---- minimal PPTX wrapper -------------------------------------------------------------

const CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`;

const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const rels = (items: [string, string, string][]) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items
    .map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`)
    .join("")}</Relationships>`;
const EMPTY_TREE = `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>`;

const DEFAULT_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements>
<a:clrScheme name="Office"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements></a:theme>`;

async function buildChartPptx(zip: JSZip, chartPath: string, cx: number, cy: number): Promise<ArrayBuffer> {
  const out = new JSZip();
  const chartXml = await zip.file(chartPath)!.async("string");
  const theme = (await zip.file("xl/theme/theme1.xml")?.async("string")) ?? DEFAULT_THEME;
  out.file("[Content_Types].xml", CT);
  out.file("_rels/.rels", rels([["rId1", "officeDocument", "ppt/presentation.xml"]]));
  out.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="${cx}" cy="${cy}"/><p:notesSz cx="${cy}" cy="${cx}"/></p:presentation>`,
  );
  out.file(
    "ppt/_rels/presentation.xml.rels",
    rels([
      ["rId1", "slideMaster", "slideMasters/slideMaster1.xml"],
      ["rId2", "slide", "slides/slide1.xml"],
      ["rId3", "theme", "theme/theme1.xml"],
    ]),
  );
  out.file(
    "ppt/slideMasters/slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster ${NS}>${EMPTY_TREE}<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
  );
  out.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    rels([
      ["rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"],
      ["rId2", "theme", "../theme/theme1.xml"],
    ]),
  );
  out.file("ppt/slideLayouts/slideLayout1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout ${NS}>${EMPTY_TREE}</p:sldLayout>`);
  out.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", rels([["rId1", "slideMaster", "../slideMasters/slideMaster1.xml"]]));
  out.file("ppt/theme/theme1.xml", theme);
  out.file(
    "ppt/slides/slide1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:bg><p:bgPr><a:noFill/><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame>` +
      `</p:spTree></p:cSld></p:sld>`,
  );
  out.file(
    "ppt/slides/_rels/slide1.xml.rels",
    rels([
      ["rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"],
      ["rId2", "chart", "../charts/chart1.xml"],
    ]),
  );
  out.file("ppt/charts/chart1.xml", chartXml);
  return out.generateAsync({ type: "arraybuffer" });
}

/** Render one chart into `container` (sized in CSS px). Returns a disposer. */
export async function renderChart(zip: JSZip, chartPath: string, widthPx: number, heightPx: number, container: HTMLElement) {
  const { PptxViewer } = await import("@aiden0z/pptx-renderer");
  const EMU = 9525;
  const buf = await buildChartPptx(zip, chartPath, Math.round(widthPx * EMU), Math.round(heightPx * EMU));
  const viewer = new PptxViewer(container, { fitMode: "none", zoomPercent: 100, pdfjs: false });
  await viewer.open(buf, { renderMode: "slide" });
  return () => viewer.destroy();
}
