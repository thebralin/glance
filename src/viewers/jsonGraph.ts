import type { SearchProvider } from "./types";
import { clearHighlights, findRanges, setHighlights } from "./domSearch";
import { childPath, countLabel, keyLabel, type JNode } from "./jsonParse";

export const GRAPH_ROWS = 50; // rows per card before "ещё N"
const GAP_X = 72; // between a card and its children
const GAP_Y = 18; // between sibling subtrees
const MIN_ZOOM = 0.1;
const MAX_ZOOM = 4;
const SVG = "http://www.w3.org/2000/svg";

interface Card {
  path: string;
  el: HTMLElement;
  kids: { row: HTMLElement; card: Card }[];
  x: number;
  y: number;
  w: number;
  h: number;
  /** Height of the card together with everything to its right. */
  sub: number;
  kidsH: number;
}

export interface GraphOptions {
  /** Containers whose child card is shown (shared with the tree's open nodes). */
  expanded: Set<string>;
  limits: Map<string, number>;
  /** Coloured value element, the same one the tree uses (inline editing hooks onto it). */
  primitive(node: JNode): HTMLElement;
  zoomChanged(): void;
}

/**
 * JSON as a mind map: every object / array is a card listing its keys; nested containers
 * hang off their row to the right, connected by curves. Drag the canvas to pan,
 * Ctrl + wheel zooms at the cursor. Layout is a left-to-right tidy tree, sized from the
 * rendered cards so long keys and values never overlap.
 */
export class JsonGraph implements SearchProvider {
  readonly el: HTMLElement;
  private world: HTMLElement;
  private edges: SVGSVGElement;
  private cards = new Map<string, Card>();
  private root: Card | null = null;
  private tx = 0;
  private ty = 0;
  private s = 1;
  private placed = false;
  private pending: { doc: JNode | null } | null = null;
  private suppressClick = false;
  private ranges: Range[] = [];
  private index = -1;

  constructor(private o: GraphOptions) {
    this.el = document.createElement("div");
    this.el.className = "json-graph";
    this.el.tabIndex = -1;
    this.world = document.createElement("div");
    this.world.className = "jg-world";
    this.edges = document.createElementNS(SVG, "svg");
    this.edges.classList.add("jg-edges");
    const hint = document.createElement("div");
    hint.className = "jg-hint";
    hint.textContent = "Тяните поле мышью · Ctrl + колесо — масштаб · двойной клик — правка";
    this.el.append(this.world, hint);
    this.bindPanZoom();
    // Rendering needs real sizes: a graph built while hidden is laid out once it shows.
    // A visible graph keeps its centre when the area changes (editor opened, window resized).
    let size = { w: 0, h: 0 };
    new ResizeObserver(() => {
      const w = this.el.clientWidth;
      const h = this.el.clientHeight;
      if (!w) return;
      if (this.pending) this.render(this.pending.doc);
      else if (size.w && this.root) {
        this.tx += (w - size.w) / 2;
        this.ty += (h - size.h) / 2;
        this.apply();
      }
      size = { w, h };
    }).observe(this.el);
  }

  get zoom() {
    return this.s;
  }

  /** Rebuild the cards. `anchor` (a card path) stays put on screen, e.g. after a toggle. */
  render(doc: JNode | null, anchor?: string) {
    if (!this.el.clientWidth) {
      this.pending = { doc };
      return;
    }
    this.pending = null;
    const before = anchor !== undefined ? this.cards.get(anchor)?.el.getBoundingClientRect() : undefined;
    this.cards.clear();
    this.root = null;
    this.edges.replaceChildren();
    this.world.replaceChildren(this.edges);
    if (!doc) return;

    const root = (this.root = this.build(doc, "", "$"));
    for (const c of this.cards.values()) {
      c.w = c.el.offsetWidth;
      c.h = c.el.offsetHeight;
    }
    measure(root);
    place(root, 0, 0);
    for (const c of this.cards.values()) {
      c.el.style.transform = `translate(${c.x}px, ${c.y}px)`;
      for (const k of c.kids) {
        const x1 = c.x + c.w;
        const y1 = c.y + k.row.offsetTop + k.row.offsetHeight / 2;
        const x2 = k.card.x;
        const y2 = k.card.y + (k.card.el.firstElementChild as HTMLElement).offsetHeight / 2;
        const p = document.createElementNS(SVG, "path");
        p.setAttribute("d", `M${x1},${y1} C${x1 + GAP_X / 2},${y1} ${x2 - GAP_X / 2},${y2} ${x2},${y2}`);
        p.dataset.to = k.card.path;
        this.edges.appendChild(p);
      }
    }

    if (!this.placed) {
      this.placed = true;
      this.home();
    } else if (before) {
      const after = this.cards.get(anchor!)?.el.getBoundingClientRect();
      if (after) {
        this.tx += before.left - after.left;
        this.ty += before.top - after.top;
        this.apply();
      }
    }
  }

  private build(node: JNode, path: string, label: string): Card {
    const el = document.createElement("div");
    el.className = `jg-card jg-${node.type}`;
    el.dataset.path = path;
    const head = document.createElement("div");
    head.className = "jg-head";
    head.append(span("jg-title", label), span("jg-type", node.entries ? countLabel(node) : typeName(node)));
    el.appendChild(head);
    this.world.appendChild(el);
    const card: Card = { path, el, kids: [], x: 0, y: 0, w: 0, h: 0, sub: 0, kidsH: 0 };
    this.cards.set(path, card);

    if (!node.entries) {
      const row = div("jg-row");
      row.dataset.path = path;
      row.appendChild(this.value(node));
      el.appendChild(row);
      return card;
    }
    if (!node.entries.length) {
      el.appendChild(div("jg-row jg-empty", node.type === "object" ? "пустой объект" : "пустой массив"));
      return card;
    }
    const limit = this.o.limits.get(path) ?? GRAPH_ROWS;
    for (const e of node.entries.slice(0, limit)) {
      const p = childPath(path, e.key);
      const row = div("jg-row");
      row.dataset.path = p;
      const k = span(typeof e.key === "number" ? "jk jk-index" : "jk", keyLabel(e.key));
      if (e.keyStart !== undefined) k.dataset.edit = "key";
      row.append(k, span("jcolon", ": "));
      if (e.node.entries) {
        const n = e.node.entries.length;
        const [ob, cb] = e.node.type === "object" ? ["{", "}"] : ["[", "]"];
        row.appendChild(span("jg-sum", n ? `${ob}${cb} ${countLabel(e.node)}` : ob + cb));
        if (n) {
          const open = this.o.expanded.has(p);
          row.classList.add("jg-link");
          row.classList.toggle("open", open);
          row.appendChild(span("jg-port", open ? "" : "+"));
          const childLabel = typeof e.key === "number" ? `${label === "$" ? "" : label}[${e.key}]` : keyLabel(e.key);
          if (open) card.kids.push({ row, card: this.build(e.node, p, childLabel) });
        }
      } else row.appendChild(this.value(e.node));
      el.appendChild(row);
    }
    const rest = node.entries.length - limit;
    if (rest > 0) {
      const more = document.createElement("button");
      more.type = "button";
      more.className = "jg-more";
      more.dataset.more = path;
      more.textContent = `ещё ${Math.min(GRAPH_ROWS, rest)} из ${rest}`;
      el.appendChild(more);
    }
    return card;
  }

  private value(node: JNode) {
    const v = this.o.primitive(node);
    if ((v.textContent ?? "").length > 40) v.title = v.textContent!;
    return v;
  }

  markSelected(path: string | null) {
    this.world.querySelectorAll(".selected").forEach((e) => e.classList.remove("selected"));
    this.edges.querySelectorAll(".hot").forEach((e) => e.classList.remove("hot"));
    if (path === null) return;
    const sel = CSS.escape(path);
    this.world.querySelector(`.jg-row[data-path="${sel}"]`)?.classList.add("selected");
    this.world.querySelector(`.jg-card[data-path="${sel}"]`)?.classList.add("selected");
    // The branch from the root to the selection.
    for (const p of this.edges.querySelectorAll<SVGPathElement>("path")) {
      const to = p.dataset.to!;
      if (path === to || path.startsWith(to + "/")) p.classList.add("hot");
    }
  }

  // ---- view -------------------------------------------------------------------------

  private apply() {
    this.world.style.transform = `translate(${this.tx}px, ${this.ty}px) scale(${this.s})`;
    let grid = 22 * this.s;
    while (grid < 11) grid *= 4; // keep the dots apart when zoomed far out
    this.el.style.backgroundSize = `${grid}px ${grid}px`;
    this.el.style.backgroundPosition = `${this.tx}px ${this.ty}px`;
  }

  /** Animate the next programmatic move (not used while dragging). */
  private smooth() {
    this.world.classList.add("smooth");
    clearTimeout((this.world as any)._t);
    (this.world as any)._t = setTimeout(() => this.world.classList.remove("smooth"), 300);
  }

  private zoomAt(z: number, cx: number, cy: number) {
    z = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z));
    this.tx = cx - ((cx - this.tx) / this.s) * z;
    this.ty = cy - ((cy - this.ty) / this.s) * z;
    this.s = z;
    this.apply();
    this.o.zoomChanged();
  }

  setZoom(z: number) {
    this.smooth();
    this.zoomAt(z, this.el.clientWidth / 2, this.el.clientHeight / 2);
  }

  private bounds() {
    let w = 0;
    let h = 0;
    for (const c of this.cards.values()) {
      w = Math.max(w, c.x + c.w);
      h = Math.max(h, c.y + c.h);
    }
    return { w, h };
  }

  /** Whole graph in the window (never enlarged past 100%). */
  fit(animate = true) {
    if (!this.root) return;
    const { w, h } = this.bounds();
    const W = this.el.clientWidth;
    const H = this.el.clientHeight;
    const s = Math.min(1, Math.max(MIN_ZOOM, Math.min((W - 64) / w, (H - 64) / h)));
    if (animate) this.smooth();
    this.s = s;
    this.tx = (W - w * s) / 2;
    this.ty = (H - h * s) / 2;
    this.apply();
    this.o.zoomChanged();
  }

  /** Everything if it is still readable at that size, otherwise the root at 100%. */
  home(minScale = 0.7, animate = false) {
    const root = this.root;
    if (!root) return;
    const { w, h } = this.bounds();
    const W = this.el.clientWidth;
    const H = this.el.clientHeight;
    if (Math.min((W - 64) / w, (H - 64) / h) >= minScale) return this.fit(animate);
    if (animate) this.smooth();
    this.s = 1;
    this.tx = 40 - root.x;
    this.ty = root.h > H - 48 ? 24 - root.y : H / 2 - root.y - root.h / 2;
    this.apply();
    this.o.zoomChanged();
  }

  /** Bring a card (or a row) into view if it is off screen. */
  reveal(path: string) {
    const el = this.world.querySelector<HTMLElement>(`.jg-card[data-path="${CSS.escape(path)}"]`);
    if (el) this.revealRect(el.getBoundingClientRect(), true);
  }

  private revealRect(r: DOMRect, keepLeft = false) {
    const box = this.el.getBoundingClientRect();
    const m = 32;
    let dx = 0;
    let dy = 0;
    if (r.right > box.right - m || r.left < box.left + m) {
      dx = keepLeft && r.width < box.width - 2 * m ? box.right - m - r.right : box.left + box.width / 2 - (r.left + r.width / 2);
      if (r.left + dx < box.left + m) dx = box.left + m - r.left;
    }
    if (r.bottom > box.bottom - m || r.top < box.top + m) {
      dy = r.height < box.height - 2 * m && !keepLeft ? box.top + box.height / 2 - (r.top + r.height / 2) : r.top < box.top + m ? box.top + m - r.top : box.bottom - m - r.bottom;
      if (r.top + dy < box.top + m) dy = box.top + m - r.top;
    }
    if (!dx && !dy) return;
    this.smooth();
    this.tx += dx;
    this.ty += dy;
    this.apply();
  }

  private bindPanZoom() {
    let drag: { x: number; y: number; tx: number; ty: number; id: number; moved: boolean } | null = null;
    this.el.addEventListener("pointerdown", (e) => {
      if (e.button !== 0 && e.button !== 1) return;
      if ((e.target as HTMLElement).closest("input, button")) return;
      drag = { x: e.clientX, y: e.clientY, tx: this.tx, ty: this.ty, id: e.pointerId, moved: false };
      if (e.button === 1) e.preventDefault(); // no autoscroll
    });
    this.el.addEventListener("pointermove", (e) => {
      if (!drag || e.pointerId !== drag.id) return;
      const dx = e.clientX - drag.x;
      const dy = e.clientY - drag.y;
      if (!drag.moved) {
        if (Math.hypot(dx, dy) < 4) return;
        drag.moved = true;
        this.el.setPointerCapture(e.pointerId);
        this.el.classList.add("panning");
        this.world.classList.remove("smooth");
      }
      this.tx = drag.tx + dx;
      this.ty = drag.ty + dy;
      this.apply();
    });
    const end = (e: PointerEvent) => {
      if (!drag || e.pointerId !== drag.id) return;
      if (drag.moved) {
        this.suppressClick = true;
        setTimeout(() => (this.suppressClick = false));
      }
      drag = null;
      this.el.classList.remove("panning");
    };
    this.el.addEventListener("pointerup", end);
    this.el.addEventListener("pointercancel", end);
    // A drag that ends over a card is not a click on it.
    this.el.addEventListener(
      "click",
      (e) => {
        if (!this.suppressClick) return;
        this.suppressClick = false;
        e.stopPropagation();
      },
      true,
    );

    this.el.addEventListener(
      "wheel",
      (e) => {
        e.preventDefault();
        e.stopPropagation(); // the shell's Ctrl+wheel would zoom in steps
        const k = e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? this.el.clientHeight : 1;
        if (e.ctrlKey || e.metaKey) {
          const r = this.el.getBoundingClientRect();
          this.world.classList.remove("smooth");
          this.zoomAt(this.s * Math.exp(-e.deltaY * k * 0.0018), e.clientX - r.left, e.clientY - r.top);
          return;
        }
        this.tx -= (e.shiftKey ? e.deltaY : e.deltaX) * k;
        this.ty -= (e.shiftKey ? 0 : e.deltaY) * k;
        this.world.classList.remove("smooth");
        this.apply();
      },
      { passive: false },
    );
  }

  // ---- search -----------------------------------------------------------------------

  async find(q: string) {
    this.ranges = findRanges(this.world, q);
    this.index = -1;
    setHighlights(this.ranges, null);
    return this.ranges.length;
  }

  step(dir: 1 | -1) {
    const n = this.ranges.length;
    if (!n) return 0;
    this.index = (this.index + dir + n) % n;
    const r = this.ranges[this.index];
    setHighlights(this.ranges, r);
    this.revealRect(r.getBoundingClientRect());
    return this.index + 1;
  }

  clear() {
    this.ranges = [];
    this.index = -1;
    clearHighlights();
  }
}

function measure(c: Card): number {
  c.kidsH = c.kids.reduce((a, k) => a + measure(k.card), 0) + GAP_Y * Math.max(0, c.kids.length - 1);
  return (c.sub = Math.max(c.h, c.kidsH));
}

/** A card is centred on the block of its children; children stack top to bottom. */
function place(c: Card, x: number, top: number) {
  c.x = x;
  c.y = top + (c.sub - c.h) / 2;
  let y = top + (c.sub - c.kidsH) / 2;
  for (const k of c.kids) {
    place(k.card, x + c.w + GAP_X, y);
    y += k.card.sub + GAP_Y;
  }
}

function typeName(node: JNode) {
  return { string: "строка", number: "число", boolean: "логическое", null: "null", object: "объект", array: "массив" }[node.type];
}

function span(cls: string, text: string) {
  const s = document.createElement("span");
  s.className = cls;
  s.textContent = text;
  return s;
}

function div(cls: string, text?: string) {
  const d = document.createElement("div");
  d.className = cls;
  if (text) d.textContent = text;
  return d;
}
