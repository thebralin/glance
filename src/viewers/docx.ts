import { renderAsync } from "docx-preview";
import { icons } from "../icons";
import { clampZoom, type EditProvider, type Viewer, type ViewerContext } from "./types";
import { DomSearch } from "./domSearch";
import { buildDocx, domText, mapParagraphs, prepareDocx, type DocxSource } from "./docxEdit";

const TAB_CLASS = "docx-tab-stop"; // docx-preview: `${className}-tab-stop`

export async function openDocx(data: ArrayBuffer, ctx: ViewerContext): Promise<Viewer> {
  const { host } = ctx;
  const root = document.createElement("div");
  root.className = "docx-doc";
  const banner = document.createElement("div");
  banner.className = "docx-edit-banner";
  banner.innerHTML = `${icons.edit}<span>Правка текста: абзацы и ячейки таблиц. <kbd>Enter</kbd> — перенос строки, новые абзацы и форматирование не создаются.</span>`;
  const styles = document.createElement("div");
  styles.hidden = true;
  const body = document.createElement("div");
  root.append(banner, styles, body);
  host.appendChild(root);

  // Editing needs paragraph markers in the rendered copy; if anything about that fails,
  // the document still opens read-only.
  let src: DocxSource | null = null;
  try {
    src = await prepareDocx(data);
  } catch (err) {
    console.warn("docx: editing unavailable", err);
  }

  await renderAsync(src?.marked ?? data, body, styles, {
    className: "docx",
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: true,
    experimental: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    renderComments: false,
    renderChanges: false,
    useBase64URL: false,
  });

  fixSymbolBullets(styles);

  const pages = () => [...body.querySelectorAll<HTMLElement>("section.docx")];
  pages().forEach((p) => p.classList.add("paper"));
  widenPages(pages());

  const editable = src ? mapParagraphs(body, src.xml, TAB_CLASS) : new Map<HTMLElement, { index: number; text: string }>();
  if (editable.size) lockReadOnlyParts(body);

  const pageWidth = () => pages()[0]?.offsetWidth || 816;
  const fitZoom = () => clampZoom(Math.min(1.5, (host.clientWidth - 64) / pageWidth()));

  let zoom = 1;
  let current = 1;
  const applyZoom = () => {
    // CSS zoom re-lays out content (crisp text), unlike transform: scale.
    body.style.zoom = String(zoom);
  };
  zoom = Math.min(1, fitZoom());
  applyZoom();

  const onScroll = () => {
    const mid = host.getBoundingClientRect().top + host.clientHeight / 3;
    const ps = pages();
    let c = 1;
    ps.forEach((p, i) => {
      if (p.getBoundingClientRect().top <= mid) c = i + 1;
    });
    if (c !== current) {
      current = c;
      ctx.changed();
    }
  };
  host.addEventListener("scroll", onScroll, { passive: true });

  const search = new DomSearch(() => body, host);

  // ---- editing --------------------------------------------------------------------------

  let editing = false;
  let everEdited = false;
  let wasDirty = false;
  let savedXml = src?.xml ?? "";
  const touched = new Set<HTMLElement>();

  const paraOf = (n: Node | null | undefined) => {
    const el = n instanceof Element ? n : n?.parentElement;
    const p = el?.closest<HTMLElement>("p[data-gp]");
    return p && body.contains(p) ? p : null;
  };
  const isDirty = () => [...touched].some((p) => domText(p, TAB_CLASS) !== editable.get(p)!.text);
  const refreshDirty = () => {
    const d = isDirty();
    if (d !== wasDirty) {
      wasDirty = d;
      ctx.changed();
    }
  };

  // Only edits that stay inside one mapped paragraph and only touch text are allowed;
  // anything else (merging paragraphs, deleting pictures, formatting) is refused.
  body.addEventListener("beforeinput", (e) => {
    if (!editing) return;
    const type = e.inputType;
    if (type === "historyUndo" || type === "historyRedo") return;
    if (
      type.startsWith("format") ||
      type === "insertParagraph" ||
      type === "insertFromDrop" ||
      type === "deleteByDrag" ||
      /^insert.*(List|HorizontalRule|Link)/.test(type)
    ) {
      e.preventDefault();
      return;
    }
    const ranges = e.getTargetRanges();
    let target: HTMLElement | null = null;
    const ok =
      ranges.length > 0 &&
      ranges.every((r) => {
        const p = paraOf(r.startContainer);
        if (!p || p !== paraOf(r.endContainer) || (target && target !== p)) return false;
        target = p;
        if (r.collapsed) return true;
        const range = document.createRange();
        range.setStart(r.startContainer, r.startOffset);
        range.setEnd(r.endContainer, r.endOffset);
        return !range.cloneContents().querySelector(`img, svg, [contenteditable="false"]:not(.${TAB_CLASS})`);
      });
    if (!ok) e.preventDefault();
    else touched.add(target!);
  });
  body.addEventListener("input", () => {
    search.clear(); // match ranges point into the old text
    const p = paraOf(getSelection()?.anchorNode);
    if (p) touched.add(p);
    refreshDirty();
  });
  body.addEventListener("keydown", (e) => {
    if (!editing || e.ctrlKey || e.metaKey || e.altKey) return;
    if (e.key === "Enter") {
      e.preventDefault();
      document.execCommand("insertLineBreak");
    } else if (e.key === "Tab" && !e.shiftKey) {
      e.preventDefault();
      document.execCommand("insertText", false, "\t");
    }
  });
  body.addEventListener("paste", (e) => {
    if (!editing) return;
    e.preventDefault();
    const text = (e.clipboardData?.getData("text/plain") ?? "").replace(/\r\n?/g, "\n");
    // Line by line: inserting "\n" as text would make the browser split the paragraph.
    text.split("\n").forEach((line, i) => {
      if (i) document.execCommand("insertLineBreak");
      if (line) document.execCommand("insertText", false, line);
    });
  });
  body.addEventListener("drop", (e) => editing && e.preventDefault());

  const toggleEdit = () => {
    search.clear();
    editing = !editing;
    everEdited ||= editing;
    root.classList.toggle("editing", editing);
    body.contentEditable = String(editing);
    body.spellcheck = false;
    if (editing) {
      // Put the caret into the first editable paragraph that is on screen.
      const top = host.getBoundingClientRect().top;
      const p = [...editable.keys()].find((el) => el.getBoundingClientRect().bottom > top + 40);
      body.focus({ preventScroll: true });
      if (p) getSelection()?.collapse(p, 0);
    } else host.focus({ preventScroll: true });
    ctx.changed();
  };

  let pending: { xml: string; texts: Map<HTMLElement, string> } | null = null;
  const editProvider: EditProvider = {
    dirty: isDirty,
    async serialize() {
      const changes = new Map<number, { from: string; to: string }>();
      const texts = new Map<HTMLElement, string>();
      for (const p of touched) {
        const info = editable.get(p)!;
        const now = domText(p, TAB_CLASS);
        if (now === info.text) continue;
        changes.set(info.index, { from: info.text, to: now });
        texts.set(p, now);
      }
      const r = await buildDocx(src!, savedXml, changes);
      pending = { xml: r.xml, texts };
      return { data: r.data };
    },
    markSaved() {
      if (!pending) return;
      savedXml = pending.xml;
      for (const [p, text] of pending.texts) editable.get(p)!.text = text;
      pending = null;
      for (const p of [...touched]) if (domText(p, TAB_CLASS) === editable.get(p)!.text) touched.delete(p);
      wasDirty = isDirty();
      ctx.changed();
    },
  };

  const viewer: Viewer = {
    get zoom() {
      return zoom;
    },
    setZoom(z) {
      z = clampZoom(z);
      const rel = host.scrollTop / Math.max(1, host.scrollHeight);
      zoom = z;
      applyZoom();
      host.scrollTop = rel * host.scrollHeight;
      ctx.changed();
    },
    fitWidth() {
      viewer.setZoom(fitZoom());
    },
    pages: {
      label: "Страница",
      total: () => pages().length,
      current: () => current,
      goTo(n) {
        const p = pages()[Math.min(Math.max(1, n), pages().length) - 1];
        p?.scrollIntoView({ block: "start" });
      },
    },
    search,
    get actions() {
      if (!editable.size) return [];
      return [
        {
          id: "edit",
          title: editing ? "Закончить правку (Ctrl+E)" : "Править текст (Ctrl+E)",
          icon: icons.edit,
          active: editing,
          run: toggleEdit,
        },
      ];
    },
    get edit() {
      // Before the first edit session nothing can be dirty.
      return everEdited ? editProvider : undefined;
    },
    summary: () => `${pages().length} стр.`,
    destroy() {
      host.removeEventListener("scroll", onScroll);
      root.remove();
    },
  };
  return viewer;
}

/**
 * docx-preview gives every page the paper width from the document and clips the rest
 * (overflow: hidden). Tables wider than the text area are common in real documents, so
 * such pages are widened to show the whole table instead of cutting it off.
 */
function widenPages(pages: HTMLElement[]) {
  for (const page of pages) {
    const box = page.getBoundingClientRect();
    const padRight = parseFloat(getComputedStyle(page).paddingRight) || 0;
    let right = 0;
    for (const el of page.querySelectorAll<HTMLElement>("table, img, svg")) {
      if (el.closest("table") !== el && el.closest("table")) continue; // the table is enough
      if (getComputedStyle(el).position === "absolute") continue; // floating shapes
      right = Math.max(right, el.getBoundingClientRect().right - box.left);
    }
    const need = Math.ceil(right + padRight);
    if (need > box.width + 1) page.style.width = `${need}px`;
  }
}

/** Everything that cannot be edited is excluded from the editing host. */
function lockReadOnlyParts(body: HTMLElement) {
  for (const el of body.querySelectorAll<HTMLElement>(`p:not([data-gp]), header, footer, img, svg, .${TAB_CLASS}`)) {
    el.contentEditable = "false";
  }
}

// Word list bullets are Private-Use-Area glyphs from Symbol/Wingdings fonts. Map them to
// Unicode so they render even if those fonts are missing or the WebView ignores the mapping.
const SYMBOL_BULLETS: Record<string, string> = {
  "": "•",
  "": "▪",
  "": "■",
  "": "◆",
  "": "❖",
  "": "➢",
  "": "➔",
  "": "✓",
  "": "❑",
  "": "o",
  "": "–",
};

function fixSymbolBullets(styles: HTMLElement) {
  for (const st of styles.querySelectorAll("style")) {
    const css = st.textContent ?? "";
    if (!/[-]/.test(css)) continue;
    st.textContent = css.replace(/(content:\s*"[^"]*")([^}]*?)font-family:\s*(?:Symbol|Wingdings)[^;]*;/g, "$1$2").replace(
      /[-]/g,
      (ch) => SYMBOL_BULLETS[ch] ?? "•",
    );
  }
}
