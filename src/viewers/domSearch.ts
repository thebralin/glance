import type { SearchProvider } from "./types";

// Uses the CSS Custom Highlight API (Chromium/WebView2) so matches are highlighted
// without mutating the rendered document DOM.
const hl = (globalThis as any).CSS?.highlights as Map<string, any> | undefined;
const HighlightCtor = (globalThis as any).Highlight as { new (...r: Range[]): any } | undefined;

export function setHighlights(all: Range[], active: Range | null) {
  if (!hl || !HighlightCtor) return;
  if (all.length) hl.set("glance-search", new HighlightCtor(...all));
  else hl.delete("glance-search");
  if (active) hl.set("glance-search-active", new HighlightCtor(active));
  else hl.delete("glance-search-active");
}

export function clearHighlights() {
  hl?.delete("glance-search");
  hl?.delete("glance-search-active");
}

/** Find all case-insensitive occurrences of `query` in text under `root`, across text nodes. */
export function findRanges(root: Node, query: string): Range[] {
  const q = query.toLocaleLowerCase();
  if (!q) return [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const p = node.parentElement;
      if (!p || p.closest("style,script,noscript")) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });
  const nodes: Text[] = [];
  const starts: number[] = [];
  let text = "";
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const t = n as Text;
    nodes.push(t);
    starts.push(text.length);
    text += t.data;
  }
  const lower = text.toLocaleLowerCase();
  // toLocaleLowerCase can change string length for a few exotic characters; bail out to a
  // per-node search in that case to keep offsets valid.
  if (lower.length !== text.length) return perNode(nodes, q);

  const locate = (pos: number): [Text, number] => {
    let lo = 0,
      hi = starts.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (starts[mid] <= pos) lo = mid;
      else hi = mid - 1;
    }
    return [nodes[lo], pos - starts[lo]];
  };

  const ranges: Range[] = [];
  for (let i = lower.indexOf(q); i >= 0; i = lower.indexOf(q, i + q.length)) {
    const [sn, so] = locate(i);
    const [en, eo] = locate(i + q.length - 1);
    const r = document.createRange();
    r.setStart(sn, so);
    r.setEnd(en, eo + 1);
    ranges.push(r);
    if (ranges.length > 5000) break;
  }
  return ranges;
}

function perNode(nodes: Text[], q: string): Range[] {
  const out: Range[] = [];
  for (const n of nodes) {
    const l = n.data.toLocaleLowerCase();
    if (l.length !== n.data.length) continue;
    for (let i = l.indexOf(q); i >= 0; i = l.indexOf(q, i + q.length)) {
      const r = document.createRange();
      r.setStart(n, i);
      r.setEnd(n, i + q.length);
      out.push(r);
    }
  }
  return out;
}

export function scrollRangeIntoView(range: Range, scroller: HTMLElement) {
  const rect = range.getBoundingClientRect();
  const box = scroller.getBoundingClientRect();
  if (rect.width === 0 && rect.height === 0) {
    range.startContainer.parentElement?.scrollIntoView({ block: "center" });
    return;
  }
  const dy = rect.top - box.top - box.height / 2 + rect.height / 2;
  const dx = rect.left < box.left + 20 || rect.right > box.right - 20 ? rect.left - box.left - box.width / 2 : 0;
  scroller.scrollBy({ top: dy, left: dx, behavior: "smooth" });
}

/** Search provider for viewers whose whole content is in the DOM. */
export class DomSearch implements SearchProvider {
  private ranges: Range[] = [];
  private index = -1;

  constructor(
    private root: () => Node,
    private scroller: HTMLElement,
    private prepare?: () => Promise<void> | void,
  ) {}

  async find(query: string): Promise<number> {
    await this.prepare?.();
    this.ranges = findRanges(this.root(), query);
    this.index = -1;
    setHighlights(this.ranges, null);
    return this.ranges.length;
  }

  step(dir: 1 | -1): number {
    const n = this.ranges.length;
    if (!n) return 0;
    this.index = (this.index + dir + n) % n;
    const r = this.ranges[this.index];
    setHighlights(this.ranges, r);
    scrollRangeIntoView(r, this.scroller);
    return this.index + 1;
  }

  clear() {
    this.ranges = [];
    this.index = -1;
    clearHighlights();
  }
}
