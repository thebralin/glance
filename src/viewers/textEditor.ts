import { DomSearch, scrollRangeIntoView } from "./domSearch";
import { encodeText, type TextFormat } from "./textCodec";
import type { EditProvider } from "./types";

/**
 * Plain-text editor: a transparent <textarea> laid over a "mirror" div with identical text
 * and metrics. The mirror sizes the editor (so the surrounding scroller scrolls, not the
 * textarea) and carries the text nodes the CSS Highlight API needs for search matches.
 */
export class TextEditor {
  readonly el: HTMLElement;
  readonly textarea: HTMLTextAreaElement;
  readonly search: DomSearch;
  private mirror: HTMLElement;
  private saved: string;
  private wasDirty = false;
  private frame = 0;

  constructor(
    text: string,
    private scroller: HTMLElement,
    private hooks: { input?(): void; dirtyChanged?(): void } = {},
  ) {
    this.saved = text;
    this.el = document.createElement("div");
    this.el.className = "text-editor";
    this.mirror = document.createElement("div");
    this.mirror.className = "text-mirror";
    this.mirror.setAttribute("aria-hidden", "true");
    this.textarea = document.createElement("textarea");
    this.textarea.spellcheck = false;
    this.textarea.wrap = "soft";
    this.textarea.value = text;
    this.el.append(this.mirror, this.textarea);
    this.syncMirror();

    this.search = new DomSearch(() => this.mirror, scroller, () => this.flush());

    this.textarea.addEventListener("input", () => {
      this.search.clear(); // match ranges point into the old mirror text
      cancelAnimationFrame(this.frame);
      this.frame = requestAnimationFrame(() => this.syncMirror());
      const dirty = this.dirty;
      if (dirty !== this.wasDirty) {
        this.wasDirty = dirty;
        this.hooks.dirtyChanged?.();
      }
      this.hooks.input?.();
    });
    this.textarea.addEventListener("keydown", (e) => {
      if (e.key === "Tab" && !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey) {
        e.preventDefault();
        // execCommand keeps the edit on the native undo stack, unlike setting .value.
        document.execCommand("insertText", false, "\t");
      }
    });
  }

  get value() {
    return this.textarea.value;
  }

  get dirty() {
    return this.textarea.value !== this.saved;
  }

  /** `text` is what was written to disk; edits made while saving keep the tab dirty. */
  markSaved(text: string) {
    this.saved = text;
    this.wasDirty = this.dirty;
    this.hooks.dirtyChanged?.();
  }

  focus() {
    this.textarea.focus({ preventScroll: true });
  }

  /** Replace part of the text. While the editor is visible the change goes on its undo stack. */
  replaceRange(start: number, end: number, text: string) {
    const ta = this.textarea;
    if (ta.offsetParent) {
      ta.focus({ preventScroll: true });
      ta.setSelectionRange(start, end);
      if (document.execCommand("insertText", false, text)) return;
    }
    ta.setRangeText(text, start, end, "end");
    ta.dispatchEvent(new Event("input"));
  }

  /** Select `pos`..`end` (or put the caret at `pos`) and scroll it into view. */
  reveal(pos: number, end = pos) {
    this.flush();
    this.focus();
    this.textarea.setSelectionRange(pos, end);
    const node = this.mirror.firstChild;
    if (!node) return;
    const r = document.createRange();
    const at = Math.min(pos, (node as Text).length - 1);
    r.setStart(node, at);
    r.setEnd(node, at + 1);
    scrollRangeIntoView(r, this.scroller);
  }

  private flush() {
    if (!this.frame) return;
    cancelAnimationFrame(this.frame);
    this.syncMirror();
  }

  private syncMirror() {
    this.frame = 0;
    // A zero-width tail gives a trailing empty line the same height the textarea shows.
    this.mirror.textContent = this.textarea.value + "​";
  }
}

/** Save support shared by the plain-text and Markdown viewers. */
export function textEditProvider(editor: TextEditor, format: { current: TextFormat }): EditProvider {
  let pending: { text: string; format: TextFormat } | null = null;
  return {
    dirty: () => editor.dirty,
    serialize() {
      const text = editor.value;
      const r = encodeText(text, format.current);
      pending = { text, format: r.format };
      return { data: r.data, note: r.note };
    },
    markSaved() {
      if (!pending) return;
      format.current = pending.format;
      editor.markSaved(pending.text);
      pending = null;
    },
  };
}
