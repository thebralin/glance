/** Common contract every format viewer implements. */
export interface Viewer {
  /** Current zoom, 1 = 100%. */
  readonly zoom: number;
  setZoom(zoom: number): void;
  /** Zoom so the content fits the window width. */
  fitWidth?(): void;
  /** Tooltip for the fit button, when it does something other than fit the width. */
  fitTitle?: string;

  /** Page / slide navigation, when the format has pages. */
  pages?: {
    label: string; // "Страница", "Слайд", "Лист"
    total(): number;
    current(): number; // 1-based
    goTo(n: number): void;
  };

  /** Text search. Returns number of matches; `step` moves between them. */
  search?: SearchProvider;

  /** Optional format-specific actions shown in the toolbar (e.g. slideshow). */
  actions?: ViewerAction[];

  /** Short description for the status bar ("12 страниц", "3 листа"). */
  summary?(): string;

  /** Editable formats: unsaved-changes tracking and serialisation for saving. */
  edit?: EditProvider;

  destroy(): void;
}

export interface ViewerAction {
  id: string;
  title: string;
  icon: string; // inline SVG
  /** Toggle actions: shown pressed while on. */
  active?: boolean;
  run(): void;
}

export interface EditProvider {
  dirty(): boolean;
  /** File contents to write. `note` is shown to the user (e.g. encoding changed). */
  serialize(): { data: Uint8Array; note?: string } | Promise<{ data: Uint8Array; note?: string }>;
  /** Called once the bytes from `serialize` are safely on disk. */
  markSaved(): void;
}

export interface SearchProvider {
  find(query: string): Promise<number>;
  /** Move to next (+1) or previous (-1) match. Returns 1-based index of active match. */
  step(dir: 1 | -1): number;
  clear(): void;
}

export interface ViewerContext {
  /** Scrollable element that hosts the document. */
  host: HTMLElement;
  /** Bar under the document (sheet tabs etc.). Hidden while empty. */
  footer: HTMLElement;
  /** Notify the shell that page / zoom / status changed. */
  changed(): void;
  /** Ask the user for a PDF password. Resolves null if cancelled. */
  askPassword(retry: boolean): Promise<string | null>;
  /** Full path of the file on disk, when known (resolving relative links / images). */
  path?: string;
}

export const ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];

export function nextZoom(current: number, dir: 1 | -1): number {
  if (dir > 0) return ZOOM_STEPS.find((z) => z > current + 0.001) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1];
  return [...ZOOM_STEPS].reverse().find((z) => z < current - 0.001) ?? ZOOM_STEPS[0];
}

export const clampZoom = (z: number) => Math.min(4, Math.max(0.25, z));
