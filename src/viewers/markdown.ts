import { Marked } from "marked";
import DOMPurify from "dompurify";
import { icons } from "../icons";
import { resolvePath, sourceFromPath } from "../platform";
import { clampZoom, type EditProvider, type SearchProvider, type Viewer, type ViewerContext } from "./types";
import { DomSearch } from "./domSearch";
import { decodeText, formatLabel } from "./textCodec";
import { TextEditor, textEditProvider } from "./textEditor";

const marked = new Marked({ gfm: true, async: false });

const MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  svg: "image/svg+xml",
  bmp: "image/bmp",
  ico: "image/x-icon",
  avif: "image/avif",
};

/** Markdown: rendered preview; Ctrl+E switches to a source editor with live preview. */
export async function openMarkdown(data: ArrayBuffer, ctx: ViewerContext): Promise<Viewer> {
  const { host } = ctx;
  const { text, format: initial } = decodeText(data);
  const format = { current: initial };

  const root = document.createElement("div");
  root.className = "md-view";
  root.innerHTML = `<div class="md-pane md-source"><div class="text-doc"></div></div><div class="md-pane md-preview"><article class="md-body"></article></div>`;
  const sourcePane = root.querySelector<HTMLElement>(".md-source")!;
  const sourceDoc = root.querySelector<HTMLElement>(".md-source .text-doc")!;
  const previewPane = root.querySelector<HTMLElement>(".md-preview")!;
  const article = root.querySelector<HTMLElement>(".md-body")!;
  host.appendChild(root);

  // Local images referenced by relative path, loaded once per viewer as blob: URLs.
  const images = new Map<string, Promise<string | null>>();
  const loadImage = (path: string) => {
    let p = images.get(path);
    if (!p) {
      const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
      p = sourceFromPath(path)
        .load()
        .then((buf) => URL.createObjectURL(new Blob([buf], { type: MIME[ext] ?? "" })))
        .catch(() => null);
      images.set(path, p);
    }
    return p;
  };

  let renderSeq = 0;
  const render = (src: string) => {
    const seq = ++renderSeq;
    article.innerHTML = DOMPurify.sanitize(marked.parse(src) as string);
    addHeadingIds(article);
    for (const img of article.querySelectorAll("img")) {
      const ref = img.getAttribute("src") ?? "";
      if (!ref || /^(data|blob):/i.test(ref)) continue;
      if (/^https?:/i.test(ref)) {
        img.classList.add("md-remote"); // blocked by CSP: the viewer works offline
        img.title = `Внешнее изображение не загружается: ${ref}`;
        continue;
      }
      img.removeAttribute("src");
      if (!ctx.path) continue;
      loadImage(resolvePath(ctx.path, ref)).then((url) => {
        if (url && seq === renderSeq) img.src = url;
      });
    }
  };
  render(text);

  let editor: TextEditor | null = null;
  let editing = false;
  let previewTimer = 0;
  const previewSearch = new DomSearch(() => article, host);

  const getEditor = () => {
    if (editor) return editor;
    editor = new TextEditor(text, sourcePane, {
      dirtyChanged: () => ctx.changed(),
      input: () => {
        clearTimeout(previewTimer);
        previewTimer = window.setTimeout(() => render(editor!.value), 200);
      },
    });
    sourceDoc.appendChild(editor.el);
    return editor;
  };

  // Keep the preview roughly aligned with the part of the source being edited.
  sourcePane.addEventListener(
    "scroll",
    () => {
      const max = sourcePane.scrollHeight - sourcePane.clientHeight;
      const ratio = max > 0 ? sourcePane.scrollTop / max : 0;
      previewPane.scrollTop = ratio * (previewPane.scrollHeight - previewPane.clientHeight);
    },
    { passive: true },
  );

  const toggleEdit = () => {
    search.clear();
    editing = !editing;
    root.classList.toggle("editing", editing);
    if (editing) getEditor().focus();
    else if (editor) {
      clearTimeout(previewTimer);
      render(editor.value);
      host.focus({ preventScroll: true });
    }
    ctx.changed();
  };

  // The shell keeps one search provider per viewer; route to whatever is on screen.
  const search: SearchProvider = {
    find: (q) => (editing ? getEditor().search : previewSearch).find(q),
    step: (dir) => (editing ? getEditor().search : previewSearch).step(dir),
    clear() {
      previewSearch.clear();
      editor?.search.clear();
    },
  };

  let editProvider: EditProvider | undefined;
  let zoom = 1;
  const viewer: Viewer = {
    get zoom() {
      return zoom;
    },
    setZoom(z) {
      zoom = clampZoom(z);
      article.style.zoom = sourceDoc.style.zoom = String(zoom);
      ctx.changed();
    },
    search,
    get actions() {
      return [
        {
          id: "edit",
          title: editing ? "Закрыть редактор (Ctrl+E)" : "Редактировать (Ctrl+E)",
          icon: icons.edit,
          active: editing,
          run: toggleEdit,
        },
      ];
    },
    get edit() {
      // Before the editor is first opened nothing can be dirty.
      if (editor && !editProvider) editProvider = textEditProvider(editor, format);
      return editProvider;
    },
    summary: () => formatLabel(format.current),
    destroy() {
      clearTimeout(previewTimer);
      renderSeq++;
      for (const p of images.values()) p.then((url) => url && URL.revokeObjectURL(url));
      root.remove();
    },
  };
  return viewer;
}

/** GitHub-style anchors so in-document links like [См. ниже](#установка) work. */
function addHeadingIds(root: HTMLElement) {
  const used = new Map<string, number>();
  for (const h of root.querySelectorAll<HTMLElement>("h1,h2,h3,h4,h5,h6")) {
    const base =
      (h.textContent ?? "")
        .trim()
        .toLowerCase()
        .replace(/[^\p{L}\p{N}\s_-]/gu, "")
        .replace(/\s+/g, "-") || "section";
    const n = used.get(base) ?? 0;
    used.set(base, n + 1);
    h.id = n ? `${base}-${n}` : base;
  }
}
