import { clampZoom, type Viewer, type ViewerContext } from "./types";
import { decodeText, formatLabel } from "./textCodec";
import { TextEditor, textEditProvider } from "./textEditor";

/** Plain text: opens straight into the editor, like Notepad. */
export async function openText(data: ArrayBuffer, ctx: ViewerContext): Promise<Viewer> {
  const { host } = ctx;
  const { text, format: initial } = decodeText(data);
  const format = { current: initial };

  const root = document.createElement("div");
  root.className = "text-doc";
  const editor = new TextEditor(text, host, { dirtyChanged: () => ctx.changed() });
  root.appendChild(editor.el);
  host.appendChild(root);

  let zoom = 1;
  return {
    get zoom() {
      return zoom;
    },
    setZoom(z) {
      zoom = clampZoom(z);
      root.style.zoom = String(zoom);
      ctx.changed();
    },
    search: editor.search,
    edit: textEditProvider(editor, format),
    summary: () => formatLabel(format.current),
    destroy() {
      root.remove();
    },
  };
}
