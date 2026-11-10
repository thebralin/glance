// The "legacy" build ships polyfills for newer JS built-ins, so it also works on
// machines with an older WebView2 runtime.
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import type { PDFDocumentProxy, PDFPageProxy, RenderTask } from "pdfjs-dist";
import workerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";
import { clampZoom, type Viewer, type ViewerContext } from "./types";
import { findRanges, setHighlights, clearHighlights, scrollRangeIntoView } from "./domSearch";

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

const ASSETS = new URL("pdfjs/", document.baseURI).href;
const GAP = 16;
/** Pages further than this from the viewport get their canvases freed. */
const KEEP_AROUND = 6;

interface PageSlot {
  n: number;
  el: HTMLDivElement;
  w: number; // unscaled (PDF points, 1 = 1/72in rendered as CSS px at 100%)
  h: number;
  page?: PDFPageProxy;
  rendered: number; // zoom it was rendered at, 0 = not rendered
  task?: RenderTask;
  textDone?: Promise<void>;
  visible: boolean;
}

export async function openPdf(data: ArrayBuffer, ctx: ViewerContext): Promise<Viewer> {
  const { host } = ctx;
  const task = pdfjs.getDocument({
    data: new Uint8Array(data),
    cMapUrl: ASSETS + "cmaps/",
    cMapPacked: true,
    standardFontDataUrl: ASSETS + "standard_fonts/",
    wasmUrl: ASSETS + "wasm/",
    iccUrl: ASSETS + "iccs/",
  });
  let retried = false;
  task.onPassword = async (update: (pw: string) => void) => {
    const pw = await ctx.askPassword(retried);
    retried = true;
    if (pw === null) task.destroy();
    else update(pw);
  };
  const doc: PDFDocumentProxy = await task.promise;

  const root = document.createElement("div");
  root.className = "pdf-doc";
  host.appendChild(root);

  const first = await doc.getPage(1);
  const vp1 = first.getViewport({ scale: 1 });
  let destroyed = false;
  let current = 1;
  // Start at "fit width", but don't blow small pages up beyond a comfortable reading size.
  let zoom = Math.min(fitZoom(), 1.25);
  const slots: PageSlot[] = [];
  for (let n = 1; n <= doc.numPages; n++) {
    const el = document.createElement("div");
    el.className = "pdf-page paper";
    el.dataset.page = String(n);
    root.appendChild(el);
    slots.push({ n, el, w: vp1.width, h: vp1.height, rendered: 0, visible: false });
  }
  slots[0].page = first;

  // Load real sizes of all pages in the background (mixed page sizes are common in scans).
  (async () => {
    for (const s of slots.slice(1)) {
      if (destroyed) return;
      try {
        const p = s.page ?? (await doc.getPage(s.n));
        s.page = p;
        const vp = p.getViewport({ scale: 1 });
        if (vp.width !== s.w || vp.height !== s.h) {
          s.w = vp.width;
          s.h = vp.height;
          layoutSlot(s);
        }
      } catch {
        /* ignore broken page */
      }
    }
  })();

  function fitZoom() {
    const avail = host.clientWidth - 64;
    return clampZoom(Math.max(0.5, avail / (vp1.width * (96 / 72))));
  }

  const CSS_UNITS = 96 / 72; // PDF points -> CSS px at 100%

  function layoutSlot(s: PageSlot) {
    const scale = zoom * CSS_UNITS;
    s.el.style.width = `${Math.floor(s.w * scale)}px`;
    s.el.style.height = `${Math.floor(s.h * scale)}px`;
    s.el.style.setProperty("--total-scale-factor", String(scale));
  }

  function layoutAll() {
    root.style.gap = `${GAP}px`;
    slots.forEach(layoutSlot);
  }
  layoutAll();

  async function renderSlot(s: PageSlot) {
    if (s.rendered === zoom || destroyed) return;
    const targetZoom = zoom;
    s.task?.cancel();
    const page = s.page ?? (await doc.getPage(s.n));
    s.page = page;
    if (targetZoom !== zoom || destroyed) return;
    const scale = zoom * CSS_UNITS;
    const vp = page.getViewport({ scale });
    const dpr = window.devicePixelRatio || 1;
    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(vp.width * dpr);
    canvas.height = Math.floor(vp.height * dpr);
    canvas.style.width = `${Math.floor(vp.width)}px`;
    canvas.style.height = `${Math.floor(vp.height)}px`;
    const renderTask = page.render({
      canvas,
      viewport: vp,
      transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
    });
    s.task = renderTask;
    try {
      await renderTask.promise;
    } catch (e: any) {
      if (e?.name === "RenderingCancelledException") return;
      throw e;
    }
    if (targetZoom !== zoom || destroyed) return;
    // Swap in the fresh canvas + text layer atomically to avoid flicker.
    const textLayer = document.createElement("div");
    textLayer.className = "textLayer";
    s.el.replaceChildren(canvas, textLayer);
    s.rendered = targetZoom;
    s.textDone = new pdfjs.TextLayer({
      textContentSource: page.streamTextContent(),
      container: textLayer,
      viewport: vp,
    })
      .render()
      .then(() => {
        if (searchQuery) refreshHighlights();
      })
      .catch(() => {});
  }

  function freeSlot(s: PageSlot) {
    s.task?.cancel();
    s.el.replaceChildren();
    s.rendered = 0;
    s.textDone = undefined;
  }

  const io = new IntersectionObserver(
    (entries) => {
      for (const e of entries) {
        const s = slots[Number((e.target as HTMLElement).dataset.page) - 1];
        s.visible = e.isIntersecting;
      }
      schedule();
    },
    { root: host, rootMargin: "100% 0px" },
  );
  slots.forEach((s) => io.observe(s.el));

  let scheduled = false;
  function schedule() {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      const vis = slots.filter((s) => s.visible);
      for (const s of vis) renderSlot(s).catch(console.error);
      if (vis.length) {
        const lo = vis[0].n - KEEP_AROUND,
          hi = vis[vis.length - 1].n + KEEP_AROUND;
        for (const s of slots) if (s.rendered && (s.n < lo || s.n > hi)) freeSlot(s);
      }
    });
  }

  function updateCurrent() {
    const mid = host.scrollTop + host.clientHeight / 3;
    let lo = 0,
      hi = slots.length - 1;
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1;
      if (slots[m].el.offsetTop <= mid) lo = m;
      else hi = m - 1;
    }
    if (current !== lo + 1) {
      current = lo + 1;
      ctx.changed();
    }
  }
  host.addEventListener("scroll", updateCurrent, { passive: true });

  function setZoom(z: number) {
    z = clampZoom(z);
    if (Math.abs(z - zoom) < 0.001) return;
    // Anchor: keep the same relative point of the current page at the top of the view.
    const s = slots[current - 1];
    const rel = (host.scrollTop - s.el.offsetTop) / s.el.offsetHeight;
    zoom = z;
    layoutAll();
    host.scrollTop = s.el.offsetTop + rel * s.el.offsetHeight;
    for (const x of slots) if (x.rendered) x.rendered = -1; // stale: re-render when visible
    schedule();
    ctx.changed();
  }

  // ---- search ----------------------------------------------------------------------------
  let searchQuery = "";
  let pageTexts: string[] | null = null;
  let matches: { page: number; nth: number }[] = [];
  let active = -1;

  async function loadTexts() {
    if (pageTexts) return pageTexts;
    const texts: string[] = [];
    for (const s of slots) {
      const p = s.page ?? (await doc.getPage(s.n));
      s.page = p;
      const tc = await p.getTextContent();
      // Same concatenation the DOM search sees: text layer spans in order, no separators.
      texts.push(tc.items.map((i: any) => i.str ?? "").join(""));
    }
    return (pageTexts = texts);
  }

  function pageRanges(s: PageSlot): Range[] {
    const tl = s.el.querySelector(".textLayer");
    return tl ? findRanges(tl, searchQuery) : [];
  }

  function refreshHighlights() {
    const all: Range[] = [];
    let act: Range | null = null;
    for (const s of slots) {
      if (!s.rendered) continue;
      const rs = pageRanges(s);
      all.push(...rs);
      const m = matches[active];
      if (m && m.page === s.n) act = rs[m.nth] ?? null;
    }
    setHighlights(all, act);
    return act;
  }

  const viewer: Viewer = {
    get zoom() {
      return zoom;
    },
    setZoom,
    fitWidth() {
      setZoom(fitZoom());
    },
    pages: {
      label: "Страница",
      total: () => doc.numPages,
      current: () => current,
      goTo(n) {
        const s = slots[Math.min(Math.max(1, n), slots.length) - 1];
        host.scrollTo({ top: s.el.offsetTop - GAP });
      },
    },
    search: {
      async find(q) {
        searchQuery = q;
        active = -1;
        matches = [];
        if (!q) {
          clearHighlights();
          return 0;
        }
        const texts = await loadTexts();
        const lq = q.toLocaleLowerCase();
        texts.forEach((t, i) => {
          const l = t.toLocaleLowerCase();
          let nth = 0;
          for (let k = l.indexOf(lq); k >= 0; k = l.indexOf(lq, k + lq.length)) matches.push({ page: i + 1, nth: nth++ });
        });
        refreshHighlights();
        return matches.length;
      },
      step(dir) {
        if (!matches.length) return 0;
        active = (active + dir + matches.length) % matches.length;
        const m = matches[active];
        const s = slots[m.page - 1];
        const box = host.getBoundingClientRect();
        const r = s.el.getBoundingClientRect();
        if (r.bottom < box.top || r.top > box.bottom) host.scrollTo({ top: s.el.offsetTop - GAP });
        const reveal = async () => {
          await renderSlot(s);
          await s.textDone;
          const act = refreshHighlights();
          if (act) scrollRangeIntoView(act, host);
        };
        reveal();
        return active + 1;
      },
      clear() {
        searchQuery = "";
        matches = [];
        active = -1;
        clearHighlights();
      },
    },
    summary: () => `${doc.numPages} стр.`,
    destroy() {
      destroyed = true;
      io.disconnect();
      host.removeEventListener("scroll", updateCurrent);
      slots.forEach((s) => s.task?.cancel());
      task.destroy();
      root.remove();
    },
  };
  schedule();
  return viewer;
}
