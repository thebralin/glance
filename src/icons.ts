const svg = (body: string, size = 18) =>
  `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${body}</svg>`;

export const icons = {
  open: svg(`<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v1"/><path d="M3 7v11a2 2 0 0 0 2 2h12.5a2 2 0 0 0 1.9-1.4L22 12H7.5a2 2 0 0 0-1.9 1.4L3 20"/>`),
  plus: svg(`<path d="M12 5v14M5 12h14"/>`, 16),
  close: svg(`<path d="M6 6l12 12M18 6L6 18"/>`, 14),
  zoomIn: svg(`<circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3M11 8v6M8 11h6"/>`),
  zoomOut: svg(`<circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3M8 11h6"/>`),
  fit: svg(`<path d="M4 12h16M4 12l3-3M4 12l3 3M20 12l-3-3M20 12l-3 3"/><path d="M4 5v2M4 17v2M20 5v2M20 17v2"/>`),
  prev: svg(`<path d="M15 18l-6-6 6-6"/>`),
  next: svg(`<path d="M9 18l6-6-6-6"/>`),
  up: svg(`<path d="M18 15l-6-6-6 6"/>`, 16),
  down: svg(`<path d="M6 9l6 6 6-6"/>`, 16),
  search: svg(`<circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/>`),
  folder: svg(`<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>`),
  play: svg(`<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M10 8.5v5l4-2.5z" fill="currentColor"/><path d="M8 21h8M12 17v4"/>`),
  upload: svg(`<path d="M12 16V4M7 9l5-5 5 5"/><path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>`, 28),
  clock: svg(`<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>`, 16),
  lock: svg(`<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>`, 28),
  alert: svg(`<path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18h.01"/>`, 28),
  edit: svg(`<path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/><path d="M13.5 6.5l4 4"/>`),
  expand: svg(`<path d="M7 9l5-5 5 5M7 15l5 5 5-5"/>`),
  collapse: svg(`<path d="M7 4l5 5 5-5M7 20l5-5 5 5"/>`),
  braces: svg(`<path d="M8 4H7a2 2 0 0 0-2 2v4l-2 2 2 2v4a2 2 0 0 0 2 2h1M16 4h1a2 2 0 0 1 2 2v4l2 2-2 2v4a2 2 0 0 1-2 2h-1"/>`),
  graph: svg(`<rect x="2.5" y="9.5" width="6" height="5" rx="1.5"/><rect x="15.5" y="3.5" width="6" height="5" rx="1.5"/><rect x="15.5" y="15.5" width="6" height="5" rx="1.5"/><path d="M8.5 12h3M11.5 6v12M11.5 6h4M11.5 18h4"/>`),
  tree: svg(`<path d="M5 4v14a2 2 0 0 0 2 2h2M5 11h4M13 6h7M13 11h7M13 20h7"/><path d="M3 4h4"/>`),
  save: svg(`<path d="M5 3h11l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M7 3v5h8V3M7 21v-7h10v7"/>`),
};

const KIND_STYLE: Record<string, { label: string; color: string; fs?: number }> = {
  pdf: { label: "PDF", color: "#E5484D" },
  docx: { label: "W", color: "#2B6CEE" },
  pptx: { label: "P", color: "#E8702A" },
  xlsx: { label: "X", color: "#1E9E5A" },
  csv: { label: "CSV", color: "#1E9E5A" },
  md: { label: "MD", color: "#6E56CF" },
  txt: { label: "TXT", color: "#64748B" },
  json: { label: "{}", color: "#C27C0E", fs: 8.5 },
};

/** Small coloured document badge used in tabs and the recent list. */
export function fileBadge(kind: string, size = 18): string {
  const k = KIND_STYLE[kind] ?? { label: "?", color: "#888" };
  const fs = k.fs ?? (k.label.length > 1 ? 6.2 : 9);
  return `<svg class="badge" width="${size}" height="${size}" viewBox="0 0 24 24" aria-hidden="true">
    <path d="M6 2h8l6 6v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z" fill="${k.color}"/>
    <path d="M14 2v4a2 2 0 0 0 2 2h4" fill="#fff" fill-opacity=".35"/>
    <text x="12" y="17.2" text-anchor="middle" font-family="Segoe UI, Inter, system-ui, sans-serif" font-weight="700" font-size="${fs}" fill="#fff">${k.label}</text>
  </svg>`;
}
