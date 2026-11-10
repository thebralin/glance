import { icons } from "../icons";
import { clampZoom, type EditProvider, type SearchProvider, type Viewer, type ViewerContext } from "./types";
import { DomSearch } from "./domSearch";
import { decodeText, formatLabel } from "./textCodec";
import { TextEditor, textEditProvider } from "./textEditor";
import {
  childPath,
  countLabel,
  detectIndent,
  formatJson,
  JsonError,
  keyLabel,
  lineCol,
  parentOf,
  parseJson,
  splitPath,
  type JEntry,
  type JNode,
} from "./jsonParse";
import { GRAPH_ROWS, JsonGraph } from "./jsonGraph";

const CHUNK = 500; // array / object entries rendered per "show more"
const AUTO_EXPAND_ROWS = 400;
const VIEW_KEY = "glance.jsonView";

/**
 * JSON: a mind-map graph (default) or a collapsible tree. Values and keys can be edited in
 * place (double click); Ctrl+E opens the source next to the view, which follows the text as
 * you type. All edits go through the text, so saving writes exactly what is in the editor.
 */
export async function openJson(data: ArrayBuffer, ctx: ViewerContext): Promise<Viewer> {
  const { host } = ctx;
  const { text, format: initial } = decodeText(data);
  const format = { current: initial };

  const root = document.createElement("div");
  root.className = "md-view json-view";
  root.innerHTML = `<div class="md-pane md-source"><div class="text-doc"></div></div>
    <div class="md-pane json-pane"><div class="json-error" hidden></div><div class="json-tree" tabindex="-1"></div></div>`;
  const sourcePane = root.querySelector<HTMLElement>(".md-source")!;
  const sourceDoc = root.querySelector<HTMLElement>(".md-source .text-doc")!;
  const treePane = root.querySelector<HTMLElement>(".json-pane")!;
  const errorBox = root.querySelector<HTMLElement>(".json-error")!;
  const tree = root.querySelector<HTMLElement>(".json-tree")!;
  host.appendChild(root);

  // Path bar under the document: where the selected value lives, with copy buttons.
  const pathBar = document.createElement("div");
  pathBar.className = "json-path";
  pathBar.innerHTML = `<span class="jp-text"></span><button type="button" data-copy="path">Копировать путь</button><button type="button" data-copy="value">Копировать значение</button>`;
  const pathText = pathBar.querySelector<HTMLElement>(".jp-text")!;

  let source = text;
  let doc: JNode | null = null;
  let error: JsonError | null = null;
  const expanded = new Set<string>();
  const limits = new Map<string, number>(); // tree
  const graphLimits = new Map<string, number>();
  let selected: string | null = null;
  let view: "graph" | "tree" = readView();
  root.classList.toggle("graph", view === "graph");
  const lim = () => (view === "graph" ? graphLimits : limits);
  const chunk = () => (view === "graph" ? GRAPH_ROWS : CHUNK);

  const graph = new JsonGraph({
    expanded,
    limits: graphLimits,
    primitive: (node) => primitive(node),
    zoomChanged: () => ctx.changed(),
  });
  treePane.appendChild(graph.el);

  const parse = () => {
    try {
      doc = parseJson(source);
      error = null;
    } catch (e) {
      if (!(e instanceof JsonError)) throw e;
      error = e; // keep the last good tree on screen
    }
  };

  // ---- tree rendering -------------------------------------------------------------------

  const nodeAt = (path: string): JNode | null => {
    let node = doc;
    for (const seg of splitPath(path)) {
      const e = node?.entries?.find((x) => String(x.key) === seg);
      if (!e) return null;
      node = e.node;
    }
    return node;
  };

  /** Redraw the current view; `anchor` is a graph card that should stay where it is. */
  const renderTree = (anchor?: string) => {
    renderError();
    if (view === "graph") {
      tree.replaceChildren();
      graph.render(doc, anchor);
    } else {
      const top = treePane.scrollTop;
      tree.replaceChildren();
      if (doc) tree.appendChild(renderNode(null, doc, "", 0));
      treePane.scrollTop = top;
    }
    if (selected !== null && !nodeAt(selected)) select(null);
    else if (selected !== null) {
      markSelected();
      pathText.textContent = displayPath(doc, selected); // a key may have been renamed
    }
  };

  const renderNode = (key: string | number | null, node: JNode, path: string, depth: number, entry?: JEntry): HTMLElement => {
    const el = document.createElement("div");
    el.className = "jn";
    el.dataset.path = path;
    const row = document.createElement("div");
    row.className = "jr";
    row.style.setProperty("--d", String(depth));
    el.appendChild(row);

    const container = node.type === "object" || node.type === "array";
    const open = container && expanded.has(path);
    const tog = document.createElement("span");
    tog.className = "jt";
    if (container && node.entries!.length) tog.classList.add(open ? "open" : "closed");
    row.appendChild(tog);

    if (key !== null) {
      const k = document.createElement("span");
      k.className = typeof key === "number" ? "jk jk-index" : "jk";
      k.textContent = keyLabel(key);
      if (entry?.keyStart !== undefined) k.dataset.edit = "key";
      row.append(k, span("jcolon", ": "));
    }

    if (!container) {
      row.appendChild(primitive(node));
      return el;
    }

    const [ob, cb] = node.type === "object" ? ["{", "}"] : ["[", "]"];
    const count = node.entries!.length;
    row.appendChild(span("jb", ob));
    if (!open || !count) {
      if (count) row.appendChild(span("jsum", countLabel(node)));
      row.appendChild(span("jb", cb));
      return el;
    }
    row.appendChild(span("jcount", String(count)));
    const kids = document.createElement("div");
    kids.className = "jc";
    const limit = limits.get(path) ?? CHUNK;
    node.entries!.slice(0, limit).forEach((e) => kids.appendChild(renderNode(e.key, e.node, childPath(path, e.key), depth + 1, e)));
    if (count > limit) {
      const more = document.createElement("button");
      more.type = "button";
      more.className = "jmore";
      more.style.setProperty("--d", String(depth + 1));
      more.dataset.more = path;
      more.textContent = `Показать ещё ${Math.min(CHUNK, count - limit)} из ${count - limit}`;
      kids.appendChild(more);
    }
    el.appendChild(kids);
    const close = document.createElement("div");
    close.className = "jr jr-close";
    close.style.setProperty("--d", String(depth));
    close.append(span("jt", ""), span("jb", cb));
    el.appendChild(close);
    return el;
  };

  const primitive = (node: JNode): HTMLElement => {
    const raw = source.slice(node.start, node.end);
    let v: HTMLElement;
    if (node.type === "string" && /^https?:\/\/\S+$/.test(node.value!)) {
      const a = document.createElement("a");
      a.href = node.value!;
      a.title = "Открыть ссылку (Ctrl + клик в режиме правки)";
      a.textContent = `"${node.value}"`;
      v = a;
    } else v = span("", node.type === "string" ? `"${node.value}"` : raw);
    v.classList.add("jv", `j-${node.type}`);
    v.dataset.edit = "value";
    return v;
  };

  const rerenderNode = (path: string) => {
    const old = tree.querySelector<HTMLElement>(`.jn[data-path="${CSS.escape(path)}"]`);
    const node = nodeAt(path);
    if (!old || !node || view === "graph") return renderTree();
    const parentPath = path === "" ? null : parentOf(path);
    const parent = parentPath === null ? null : nodeAt(parentPath);
    const entry = parent?.entries?.find((e) => childPath(parentPath!, e.key) === path);
    const depth = Number(old.querySelector<HTMLElement>(".jr")!.style.getPropertyValue("--d"));
    old.replaceWith(renderNode(entry ? entry.key : null, node, path, depth, entry));
    if (selected !== null) markSelected();
  };

  const renderError = () => {
    errorBox.hidden = !error;
    tree.classList.toggle("stale", !!error);
    graph.el.classList.toggle("stale", !!error);
    if (!error) return;
    const { line, col } = lineCol(source, error.pos);
    errorBox.innerHTML = `${icons.alert}<span></span>`;
    errorBox.querySelector("span")!.textContent = `Ошибка JSON (строка ${line}, столбец ${col}): ${error.message}`;
    errorBox.title = "Показать в тексте";
  };

  /** Expand from the top, breadth first, while the tree stays small enough to scan. */
  const autoExpand = (budget: number) => {
    if (!doc) return;
    let rows = 1;
    let level: [JNode, string][] = [[doc, ""]];
    while (level.length && rows < budget) {
      const next: [JNode, string][] = [];
      for (const [node, path] of level) {
        if (!node.entries?.length) continue;
        const shown = Math.min(node.entries.length, lim().get(path) ?? chunk());
        if (path !== "" && rows + shown > budget) continue;
        expanded.add(path);
        rows += shown;
        for (const e of node.entries.slice(0, shown)) next.push([e.node, childPath(path, e.key)]);
      }
      level = next;
    }
  };

  const setAll = (open: boolean) => {
    expanded.clear();
    if (open) autoExpand(20000);
    else expanded.add("");
    renderTree();
    if (view === "graph") graph.home(0.35, true);
    ctx.changed();
  };

  // ---- selection & path bar -------------------------------------------------------------

  const markSelected = () => {
    if (view === "graph") return graph.markSelected(selected);
    tree.querySelector(".jr.selected")?.classList.remove("selected");
    if (selected === null) return;
    tree.querySelector(`.jn[data-path="${CSS.escape(selected)}"] > .jr`)?.classList.add("selected");
  };

  const select = (path: string | null) => {
    selected = path;
    markSelected();
    if (path === null) {
      pathBar.remove();
    } else {
      pathText.textContent = displayPath(doc, path);
      if (!pathBar.isConnected) ctx.footer.appendChild(pathBar);
    }
    ctx.changed();
  };

  pathBar.addEventListener("click", (e) => {
    const what = (e.target as HTMLElement).closest<HTMLElement>("[data-copy]")?.dataset.copy;
    if (!what || selected === null) return;
    const node = nodeAt(selected);
    const value = what === "path" ? displayPath(doc, selected) : node ? (node.type === "string" ? node.value! : source.slice(node.start, node.end)) : "";
    navigator.clipboard.writeText(value).catch(() => {});
  });

  // ---- interaction ----------------------------------------------------------------------

  tree.addEventListener("click", (e) => {
    const t = e.target as HTMLElement;
    if (t.closest("input")) return;
    const more = t.closest<HTMLElement>("[data-more]");
    if (more) {
      const p = more.dataset.more!;
      limits.set(p, (limits.get(p) ?? CHUNK) + CHUNK);
      rerenderNode(p);
      return;
    }
    const row = t.closest<HTMLElement>(".jr");
    if (!row || row.classList.contains("jr-close")) return;
    const path = row.parentElement!.dataset.path!;
    const node = nodeAt(path);
    if (t.closest(".jt, .jb, .jsum, .jcount") || (t === row && node?.entries?.length)) {
      if (node?.entries?.length) {
        if (expanded.has(path)) expanded.delete(path);
        else expanded.add(path);
        rerenderNode(path);
      }
    }
    select(path);
    // Show the value in the source as well.
    if (editing && node && !error) editor!.reveal(node.start, node.end);
  });

  graph.el.addEventListener("click", (e) => {
    const t = e.target as HTMLElement;
    if (t.closest("input")) return;
    const more = t.closest<HTMLElement>("[data-more]");
    if (more) {
      const p = more.dataset.more!;
      graphLimits.set(p, (graphLimits.get(p) ?? GRAPH_ROWS) + GRAPH_ROWS);
      renderTree(p);
      return;
    }
    const row = t.closest<HTMLElement>(".jg-row[data-path], .jg-head");
    if (!row) return;
    const path = row.closest<HTMLElement>("[data-path]")!.dataset.path!;
    const card = row.closest<HTMLElement>(".jg-card")!.dataset.path!;
    if (row.classList.contains("jg-link") && (t === row || t.closest(".jg-sum, .jg-port"))) {
      const opening = !expanded.has(path);
      if (opening) expanded.add(path);
      else expanded.delete(path);
      select(path);
      renderTree(card);
      if (opening) graph.reveal(path);
      return;
    }
    select(path);
    if (editing && !error) {
      const node = nodeAt(path);
      if (node) editor!.reveal(node.start, node.end);
    }
  });

  const onDblClick = (e: MouseEvent) => {
    const target = (e.target as HTMLElement).closest<HTMLElement>("[data-edit]");
    if (!target || error) return;
    e.preventDefault();
    const path = target.closest<HTMLElement>("[data-path]")!.dataset.path!;
    startInlineEdit(target, path, target.dataset.edit as "key" | "value");
  };
  tree.addEventListener("dblclick", onDblClick);
  graph.el.addEventListener("dblclick", onDblClick);

  const startInlineEdit = (target: HTMLElement, path: string, what: "key" | "value") => {
    const node = nodeAt(path);
    const parentPath = parentOf(path);
    const entry = nodeAt(parentPath)?.entries?.find((x) => childPath(parentPath, x.key) === path);
    if (!node || !entry) return;
    const isString = what === "key" || node.type === "string";
    const input = document.createElement("input");
    input.className = "jedit";
    input.spellcheck = false;
    input.value = what === "key" ? String(entry.key) : isString ? node.value! : source.slice(node.start, node.end);
    input.style.width = `${Math.min(60, Math.max(6, input.value.length + 2))}ch`;
    target.replaceWith(input);
    input.focus();
    input.select();

    let done = false;
    const finish = (commit: boolean) => {
      if (done) return;
      done = true;
      const v = input.value;
      const home = view === "graph" ? graph.el : tree;
      if (!commit || v === (what === "key" ? String(entry.key) : isString ? node.value : source.slice(node.start, node.end))) {
        input.replaceWith(target);
        home.focus({ preventScroll: true });
        return;
      }
      // Strings stay strings; other values accept any JSON (e.g. 42 → "abc" or → {"a": 1}).
      const replacement = isString ? JSON.stringify(v) : (literal(v) ?? JSON.stringify(v));
      const [start, end] = what === "key" ? [entry.keyStart!, entry.keyEnd!] : [node.start, node.end];
      if (what === "key") {
        // Keep the subtree open under its new name.
        const newPath = childPath(parentPath, v);
        for (const p of [...expanded]) if (p === path || p.startsWith(path + "/")) expanded.add(newPath + p.slice(path.length));
        if (selected === path) selected = newPath;
      }
      getEditor().replaceRange(start, end, replacement);
      updateFromSource(true);
      home.focus({ preventScroll: true });
    };
    input.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") finish(true);
      else if (e.key === "Escape") finish(false);
    });
    input.addEventListener("blur", () => finish(true));
  };

  errorBox.addEventListener("click", () => {
    if (!error) return;
    if (!editing) toggleEdit();
    getEditor().reveal(error.pos);
  });

  // ---- source editor --------------------------------------------------------------------

  let editor: TextEditor | null = null;
  let editing = false;
  let timer = 0;

  const updateFromSource = (now = false) => {
    clearTimeout(timer);
    const run = () => {
      source = editor!.value;
      parse();
      renderTree();
      ctx.changed();
    };
    if (now) run();
    else timer = window.setTimeout(run, 250);
  };

  const getEditor = () => {
    if (editor) return editor;
    editor = new TextEditor(source, sourcePane, {
      dirtyChanged: () => ctx.changed(),
      input: () => updateFromSource(),
    });
    sourceDoc.appendChild(editor.el);
    return editor;
  };

  const toggleEdit = () => {
    search.clear();
    editing = !editing;
    root.classList.toggle("editing", editing);
    if (editing) getEditor().focus();
    else {
      if (editor) updateFromSource(true);
      (view === "graph" ? graph.el : tree).focus({ preventScroll: true });
    }
    ctx.changed();
  };

  const toggleView = () => {
    search.clear();
    view = view === "graph" ? "tree" : "graph";
    root.classList.toggle("graph", view === "graph");
    try {
      localStorage.setItem(VIEW_KEY, view);
    } catch {}
    if (view === "tree") graph.render(null); // drop the cards, the tree takes over
    renderTree();
    if (view === "graph" && selected !== null) graph.reveal(selected);
    ctx.changed();
  };

  const reformat = () => {
    if (!doc || error) return;
    const pretty = formatJson(source, doc, detectIndent(source));
    if (pretty === source) return;
    getEditor().replaceRange(0, source.length, pretty);
    updateFromSource(true);
  };

  // ---- search ---------------------------------------------------------------------------

  const treeSearch = new DomSearch(() => tree, host);
  const revealMatches = (q: string) => {
    if (!doc) return;
    const needle = q.toLocaleLowerCase();
    let found = 0;
    const walk = (node: JNode, path: string): boolean => {
      let hit = false;
      node.entries?.forEach((e, idx) => {
        if (found > 2000) return;
        const p = childPath(path, e.key);
        const text = e.node.type === "string" ? e.node.value! : e.node.entries ? "" : source.slice(e.node.start, e.node.end);
        let h = (typeof e.key === "string" && e.key.toLocaleLowerCase().includes(needle)) || text.toLocaleLowerCase().includes(needle);
        if (h) found++;
        if (walk(e.node, p)) h = true;
        if (h) {
          hit = true;
          if (idx >= (lim().get(path) ?? chunk())) lim().set(path, Math.ceil((idx + 1) / chunk()) * chunk());
        }
      });
      if (hit) expanded.add(path);
      return hit;
    };
    walk(doc, "");
    renderTree();
  };
  const search: SearchProvider = {
    async find(q) {
      if (editing) return getEditor().search.find(q);
      revealMatches(q);
      return (view === "graph" ? graph : treeSearch).find(q);
    },
    step: (dir) => (editing ? getEditor().search : view === "graph" ? graph : treeSearch).step(dir),
    clear() {
      treeSearch.clear();
      graph.clear();
      editor?.search.clear();
    },
  };

  // ---- start ----------------------------------------------------------------------------

  parse();
  autoExpand(AUTO_EXPAND_ROWS);
  expanded.add("");
  renderTree();
  if (error) toggleEdit(); // nothing to show as a tree: open the text

  let editProvider: EditProvider | undefined;
  let zoom = 1;
  const viewer: Viewer = {
    get zoom() {
      return view === "graph" ? graph.zoom : zoom;
    },
    setZoom(z) {
      if (view === "graph") return graph.setZoom(z);
      zoom = clampZoom(z);
      tree.style.zoom = sourceDoc.style.zoom = String(zoom);
      ctx.changed();
    },
    get fitWidth() {
      return view === "graph" ? () => graph.fit() : undefined;
    },
    fitTitle: "Вписать схему в окно",
    search,
    get actions() {
      return [
        {
          id: "view",
          title: view === "graph" ? "Показать деревом" : "Показать схемой",
          icon: view === "graph" ? icons.tree : icons.graph,
          run: toggleView,
        },
        {
          id: "edit",
          title: editing ? "Закрыть редактор (Ctrl+E)" : "Редактировать текст (Ctrl+E)",
          icon: icons.edit,
          active: editing,
          run: toggleEdit,
        },
        { id: "expand", title: "Развернуть всё", icon: icons.expand, run: () => setAll(true) },
        { id: "collapse", title: "Свернуть всё", icon: icons.collapse, run: () => setAll(false) },
        { id: "format", title: "Отформатировать (выровнять отступы)", icon: icons.braces, run: reformat },
      ];
    },
    get edit() {
      // Before the editor exists nothing can be dirty.
      if (editor && !editProvider) editProvider = textEditProvider(editor, format);
      return editProvider;
    },
    summary() {
      const what = !doc
        ? "Ошибка JSON"
        : doc.type === "object"
          ? `Объект · ${countLabel(doc)}`
          : doc.type === "array"
            ? `Массив · ${countLabel(doc)}`
            : "JSON";
      return `${error && doc ? "Ошибка JSON · " : ""}${what} · ${formatLabel(format.current)}`;
    },
    destroy() {
      clearTimeout(timer);
      pathBar.remove();
      root.remove();
    },
  };
  return viewer;
}

// ---------------------------------------------------------------------------------------

function readView(): "graph" | "tree" {
  try {
    return localStorage.getItem(VIEW_KEY) === "tree" ? "tree" : "graph";
  } catch {
    return "graph";
  }
}

/** `$.users[3]["first name"]` — the form most tools accept. */
function displayPath(root: JNode | null, path: string): string {
  let out = "$";
  let node = root;
  for (const seg of splitPath(path)) {
    if (node?.type === "array") out += `[${seg}]`;
    else out += /^[A-Za-z_$][\w$]*$/.test(seg) ? `.${seg}` : `[${JSON.stringify(seg)}]`;
    node = node?.entries?.find((e) => String(e.key) === seg)?.node ?? null;
  }
  return out;
}

/** A valid JSON value typed by the user, normalised; null if it is not one. */
function literal(v: string): string | null {
  const t = v.trim();
  try {
    parseJson(t);
    return t;
  } catch {
    return null;
  }
}

function span(cls: string, text: string) {
  const s = document.createElement("span");
  if (cls) s.className = cls;
  s.textContent = text;
  return s;
}
