import { PptxViewer, RECOMMENDED_ZIP_LIMITS } from "@aiden0z/pptx-renderer";
import pdfModuleUrl from "pdfjs-dist/legacy/build/pdf.min.mjs?url";
import pdfWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";
import { clampZoom, type Viewer, type ViewerContext } from "./types";
import { setFullscreen } from "../platform";
import { icons } from "../icons";

type Result = ReturnType<PptxViewer["searchText"]>[number];

export async function openPptx(data: ArrayBuffer, ctx: ViewerContext): Promise<Viewer> {
  const { host } = ctx;
  // The renderer fits slides to its container's width, so padding lives on a wrapper.
  const wrap = document.createElement("div");
  wrap.className = "pptx-wrap";
  const root = document.createElement("div");
  root.className = "pptx-doc";
  wrap.appendChild(root);
  host.appendChild(wrap);

  let current = 1;
  const viewer = await PptxViewer.open(data, root, {
    zipLimits: RECOMMENDED_ZIP_LIMITS,
    scrollContainer: host,
    fitMode: "contain",
    zoomPercent: 100,
    lazyMedia: true,
    lazySlides: true,
    listOptions: { windowed: true, initialSlides: 4, batchSize: 4, showSlideLabels: false },
    pdfjs: {
      moduleUrl: new URL(pdfModuleUrl, document.baseURI).href,
      workerUrl: new URL(pdfWorkerUrl, document.baseURI).href,
    },
    onSlideChange(index) {
      current = index + 1;
      ctx.changed();
    },
  });

  // ---- search -----------------------------------------------------------------------
  let results: Result[] = [];
  let active = -1;

  // ---- slideshow --------------------------------------------------------------------
  function slideshow(start = current - 1) {
    const overlay = document.createElement("div");
    overlay.className = "slideshow";
    const stage = document.createElement("div");
    stage.className = "slideshow-stage";
    const counter = document.createElement("div");
    counter.className = "slideshow-counter";
    overlay.append(stage, counter);
    document.body.appendChild(overlay);
    let idx = Math.max(0, Math.min(start, viewer.slideCount - 1));
    let handle: { dispose(): void } | null = null;

    const show = () => {
      handle?.dispose();
      stage.replaceChildren();
      const scale = Math.min(window.innerWidth / viewer.slideWidth, window.innerHeight / viewer.slideHeight);
      stage.style.width = `${viewer.slideWidth * scale}px`;
      stage.style.height = `${viewer.slideHeight * scale}px`;
      handle = viewer.renderSlideToContainer(idx, stage, scale);
      counter.textContent = `${idx + 1} / ${viewer.slideCount}`;
    };
    const go = (d: number) => {
      const n = Math.max(0, Math.min(viewer.slideCount - 1, idx + d));
      if (n !== idx) {
        idx = n;
        show();
      }
    };
    const close = () => {
      window.removeEventListener("keydown", onKey, true);
      window.removeEventListener("resize", show);
      handle?.dispose();
      overlay.remove();
      setFullscreen(false);
      viewer.goToSlide(idx);
    };
    const onKey = (e: KeyboardEvent) => {
      e.stopPropagation();
      e.preventDefault();
      if (["ArrowRight", "ArrowDown", "PageDown", " ", "Enter", "n"].includes(e.key)) go(1);
      else if (["ArrowLeft", "ArrowUp", "PageUp", "Backspace", "p"].includes(e.key)) go(-1);
      else if (e.key === "Home") go(-idx);
      else if (e.key === "End") go(viewer.slideCount);
      else if (e.key === "Escape" || e.key === "F5") close();
    };
    overlay.addEventListener("click", (e) => (e.clientX < window.innerWidth / 4 ? go(-1) : go(1)));
    overlay.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      go(-1);
    });
    window.addEventListener("keydown", onKey, true);
    window.addEventListener("resize", show);
    setFullscreen(true).then(() => setTimeout(show, 150));
    show();
  }

  return {
    get zoom() {
      return viewer.zoomPercent / 100;
    },
    setZoom(z) {
      viewer.setZoom(Math.round(clampZoom(z) * 100)).then(() => ctx.changed());
    },
    fitWidth() {
      viewer.setZoom(100).then(() => ctx.changed());
    },
    pages: {
      label: "Слайд",
      total: () => viewer.slideCount,
      current: () => current,
      goTo(n) {
        viewer.goToSlide(Math.min(Math.max(1, n), viewer.slideCount) - 1);
      },
    },
    search: {
      async find(q) {
        viewer.clearSearchHighlights();
        active = -1;
        results = q ? viewer.searchText(q) : [];
        return results.length;
      },
      step(dir) {
        if (!results.length) return 0;
        active = (active + dir + results.length) % results.length;
        viewer.clearSearchHighlights();
        viewer.highlightSearchResult(results[active], {
          scrollIntoView: { block: "center", behavior: "smooth" },
          borderColor: "var(--accent)",
          backgroundColor: "color-mix(in srgb, var(--accent) 22%, transparent)",
          borderRadius: 4,
        });
        return active + 1;
      },
      clear() {
        results = [];
        active = -1;
        viewer.clearSearchHighlights();
      },
    },
    actions: [{ id: "slideshow", title: "Показ слайдов (F5)", icon: icons.play, run: () => slideshow() }],
    summary: () => `${viewer.slideCount} слайд.`,
    destroy() {
      viewer.destroy();
      wrap.remove();
    },
  };
}
