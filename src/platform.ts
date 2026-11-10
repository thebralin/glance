// Thin layer over Tauri APIs, with a plain-browser fallback so the UI can also be
// developed and tested with `npm run dev` in a regular browser.
import { invoke } from "@tauri-apps/api/core";

export const isTauri = typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;

export type DocKind = "pdf" | "docx" | "pptx" | "xlsx" | "csv" | "md" | "txt" | "json";

export const EXTENSIONS: Record<DocKind, string[]> = {
  pdf: ["pdf"],
  docx: ["docx", "docm", "dotx"],
  pptx: ["pptx", "pptm", "ppsx", "potx"],
  xlsx: ["xlsx", "xlsm", "xltx"],
  csv: ["csv", "tsv"],
  md: ["md", "markdown"],
  txt: ["txt"],
  json: ["json"],
};

export const ALL_EXTENSIONS = Object.values(EXTENSIONS).flat();

/** Legacy binary formats we recognise only to show a friendly message. */
const LEGACY: Record<string, string> = {
  doc: "DOC (Word 97–2003)",
  ppt: "PPT (PowerPoint 97–2003)",
  xls: "XLS (Excel 97–2003)",
};

export function kindOf(name: string): DocKind | null {
  const ext = extOf(name);
  for (const [kind, list] of Object.entries(EXTENSIONS)) {
    if (list.includes(ext)) return kind as DocKind;
  }
  return null;
}

export function legacyFormat(name: string): string | null {
  return LEGACY[extOf(name)] ?? null;
}

export function extOf(name: string): string {
  const i = name.lastIndexOf(".");
  return i >= 0 ? name.slice(i + 1).toLowerCase() : "";
}

export function baseName(path: string): string {
  return path.split(/[\\/]/).pop() || path;
}

/** Resolve a link / image reference found inside the document at `fromFile` to a file path. */
export function resolvePath(fromFile: string, href: string): string {
  let rel = href.replace(/[?#].*$/, "");
  try {
    rel = decodeURIComponent(rel);
  } catch {
    /* keep as written */
  }
  const win = fromFile.includes("\\") || /^[a-zA-Z]:/.test(fromFile);
  const sep = win ? "\\" : "/";
  if (/^file:/i.test(rel)) rel = rel.replace(/^file:\/*/i, win ? "" : "/");
  const absolute = /^[a-zA-Z]:[\\/]/.test(rel) || /^[\\/]/.test(rel);
  const full = absolute ? rel : fromFile.replace(/[^\\/]*$/, "") + rel;
  const lead = full.match(/^[\\/]*/)![0].length;
  const out: string[] = [];
  for (const seg of full.slice(lead).split(/[\\/]+/)) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") {
      if (out.length > 1 || (out.length === 1 && !/^[a-zA-Z]:$/.test(out[0]))) out.pop();
    } else out.push(seg);
  }
  return sep.repeat(Math.min(lead, 2)) + out.join(sep);
}

/** Something that can be opened: a path on disk (Tauri) or a browser File. */
export interface Source {
  name: string;
  path?: string;
  size?: number;
  load(): Promise<ArrayBuffer>;
}

export function sourceFromPath(path: string): Source {
  return {
    name: baseName(path),
    path,
    async load() {
      const data = await invoke<ArrayBuffer>("read_file", { path });
      // invoke() returns an ArrayBuffer for tauri::ipc::Response payloads.
      return data instanceof ArrayBuffer ? data : new Uint8Array(data as unknown as number[]).buffer;
    },
  };
}

export function sourceFromFile(file: File): Source {
  return { name: file.name, size: file.size, load: () => file.arrayBuffer() };
}

export async function pickFiles(): Promise<Source[]> {
  if (isTauri) {
    const { open } = await import("@tauri-apps/plugin-dialog");
    const result = await open({
      multiple: true,
      directory: false,
      title: "Открыть документ",
      filters: [
        { name: "Все документы", extensions: ALL_EXTENSIONS },
        { name: "PDF", extensions: EXTENSIONS.pdf },
        { name: "Word", extensions: EXTENSIONS.docx },
        { name: "PowerPoint", extensions: EXTENSIONS.pptx },
        { name: "Excel / CSV", extensions: [...EXTENSIONS.xlsx, ...EXTENSIONS.csv] },
        { name: "Текст / Markdown / JSON", extensions: [...EXTENSIONS.txt, ...EXTENSIONS.md, ...EXTENSIONS.json] },
      ],
    });
    if (!result) return [];
    const paths = Array.isArray(result) ? result : [result];
    return paths.map((p) => sourceFromPath(String(p)));
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.multiple = true;
    input.accept = ALL_EXTENSIONS.map((e) => "." + e).join(",");
    input.onchange = () => resolve([...(input.files ?? [])].map(sourceFromFile));
    input.click();
  });
}

/** Overwrite a file on disk (Tauri only; the backend allows only formats with an editor). */
export async function writeFile(path: string, data: Uint8Array) {
  await invoke("write_file", data, { headers: { path: encodeURIComponent(path) } });
}

/** Browser fallback for saving: hand the bytes to the user as a download. */
export function downloadFile(name: string, data: Uint8Array) {
  const url = URL.createObjectURL(new Blob([data as BlobPart]));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/**
 * Intercept closing the app window. `canClose` may ask the user; resolving false keeps
 * the window open.
 */
export async function onCloseRequested(canClose: () => Promise<boolean>, hasUnsaved: () => boolean) {
  if (!isTauri) {
    window.addEventListener("beforeunload", (e) => {
      if (hasUnsaved()) e.preventDefault();
    });
    return;
  }
  const { getCurrentWindow } = await import("@tauri-apps/api/window");
  await getCurrentWindow().onCloseRequested(async (e) => {
    if (!(await canClose())) e.preventDefault();
  });
}

export async function fileInfo(path: string): Promise<{ size: number; modified: number | null } | null> {
  if (!isTauri) return null;
  try {
    return await invoke("file_info", { path });
  } catch {
    return null;
  }
}

export async function takeStartupFiles(): Promise<string[]> {
  if (!isTauri) return [];
  try {
    return await invoke<string[]>("take_startup_files");
  } catch {
    return [];
  }
}

export async function onOpenFilesEvent(cb: (paths: string[]) => void) {
  if (!isTauri) return;
  const { listen } = await import("@tauri-apps/api/event");
  await listen<string[]>("open-files", (e) => cb(e.payload));
}

export type DropState = "over" | "leave" | "drop";

/** Native drag & drop (Tauri gives full paths); falls back to HTML5 events in a browser. */
export async function onFileDrop(cb: (state: DropState, sources: Source[]) => void) {
  if (isTauri) {
    const { getCurrentWebview } = await import("@tauri-apps/api/webview");
    await getCurrentWebview().onDragDropEvent((event) => {
      const p = event.payload;
      if (p.type === "enter" || p.type === "over") cb("over", []);
      else if (p.type === "leave") cb("leave", []);
      else if (p.type === "drop") cb("drop", p.paths.map(sourceFromPath));
    });
    return;
  }
  let depth = 0;
  window.addEventListener("dragenter", (e) => {
    e.preventDefault();
    depth++;
    cb("over", []);
  });
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("dragleave", () => {
    if (--depth <= 0) {
      depth = 0;
      cb("leave", []);
    }
  });
  window.addEventListener("drop", (e) => {
    e.preventDefault();
    depth = 0;
    cb("drop", [...(e.dataTransfer?.files ?? [])].map(sourceFromFile));
  });
}

export async function revealInFolder(path: string) {
  if (!isTauri) return;
  const { revealItemInDir } = await import("@tauri-apps/plugin-opener");
  await revealItemInDir(path);
}

/** Open a web/mail link in the system browser instead of navigating the app's webview. */
export async function openExternal(url: string) {
  if (!isTauri) {
    window.open(url, "_blank", "noopener");
    return;
  }
  const { openUrl } = await import("@tauri-apps/plugin-opener");
  await openUrl(url);
}

export async function setWindowTitle(title: string) {
  document.title = title;
  if (!isTauri) return;
  const { getCurrentWindow } = await import("@tauri-apps/api/window");
  await getCurrentWindow().setTitle(title).catch(() => {});
}

export async function setFullscreen(on: boolean) {
  if (isTauri) {
    const { getCurrentWindow } = await import("@tauri-apps/api/window");
    await getCurrentWindow().setFullscreen(on).catch(() => {});
  } else if (on) {
    await document.documentElement.requestFullscreen?.().catch(() => {});
  } else if (document.fullscreenElement) {
    await document.exitFullscreen().catch(() => {});
  }
}
