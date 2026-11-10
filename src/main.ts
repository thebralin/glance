import "./styles/app.css";
import { icons, fileBadge } from "./icons";
import {
  isTauri,
  kindOf,
  legacyFormat,
  pickFiles,
  sourceFromPath,
  takeStartupFiles,
  onOpenFilesEvent,
  onFileDrop,
  revealInFolder,
  openExternal,
  setWindowTitle,
  fileInfo,
  writeFile,
  downloadFile,
  onCloseRequested,
  resolvePath,
  type DocKind,
  type Source,
} from "./platform";
import { nextZoom, type Viewer, type ViewerContext } from "./viewers/types";

// ---------------------------------------------------------------------------------------
// Layout
// ---------------------------------------------------------------------------------------

const app = document.getElementById("app")!;
app.innerHTML = `
  <header class="topbar">
    <div class="tabs" role="tablist"></div>
    <button class="icon-btn new-tab" title="Открыть файл (Ctrl+O)">${icons.plus}</button>
  </header>
  <div class="toolbar" hidden>
    <div class="tb-group pager">
      <button class="icon-btn" data-act="prev" title="Предыдущая (PgUp)">${icons.prev}</button>
      <span class="pager-label"></span>
      <input class="pager-input" inputmode="numeric" aria-label="Номер страницы" />
      <span class="pager-total"></span>
      <button class="icon-btn" data-act="next" title="Следующая (PgDn)">${icons.next}</button>
    </div>
    <div class="tb-summary"></div>
    <div class="tb-spacer"></div>
    <div class="tb-group zoom">
      <button class="icon-btn" data-act="zoom-out" title="Уменьшить (Ctrl −)">${icons.zoomOut}</button>
      <button class="zoom-value" data-act="zoom-reset" title="Сбросить до 100% (Ctrl 0)">100%</button>
      <button class="icon-btn" data-act="zoom-in" title="Увеличить (Ctrl +)">${icons.zoomIn}</button>
      <button class="icon-btn" data-act="fit" title="По ширине окна">${icons.fit}</button>
    </div>
    <div class="tb-sep"></div>
    <div class="tb-group actions"></div>
    <button class="icon-btn" data-act="save" title="Сохранить (Ctrl+S)">${icons.save}</button>
    <button class="icon-btn" data-act="search" title="Поиск (Ctrl+F)">${icons.search}</button>
    <button class="icon-btn" data-act="reveal" title="Показать в папке">${icons.folder}</button>
  </div>
  <div class="searchbar" hidden>
    ${icons.search}
    <input type="search" placeholder="Найти в документе" spellcheck="false" />
    <span class="search-count"></span>
    <button class="icon-btn" data-act="search-prev" title="Предыдущее (Shift+Enter)">${icons.up}</button>
    <button class="icon-btn" data-act="search-next" title="Следующее (Enter)">${icons.down}</button>
    <button class="icon-btn" data-act="search-close" title="Закрыть (Esc)">${icons.close}</button>
  </div>
  <main class="stage">
    <section class="welcome">
      <div class="welcome-inner">
        <div class="brand">
          <img src="/icon.png" alt="" width="72" height="72" />
          <div>
            <h1>Glance</h1>
            <p>Быстрый просмотр документов без Microsoft Office</p>
          </div>
        </div>
        <button class="dropzone">
          <span class="dz-icon">${icons.upload}</span>
          <span class="dz-title">Перетащите файлы сюда</span>
          <span class="dz-sub">или нажмите, чтобы выбрать · Ctrl+O</span>
          <span class="formats">
            <span>${fileBadge("docx", 16)} Word</span>
            <span>${fileBadge("pptx", 16)} PowerPoint</span>
            <span>${fileBadge("xlsx", 16)} Excel</span>
            <span>${fileBadge("pdf", 16)} PDF</span>
            <span>${fileBadge("csv", 16)} CSV</span>
            <span>${fileBadge("md", 16)} Markdown</span>
            <span>${fileBadge("txt", 16)} Текст</span>
            <span>${fileBadge("json", 16)} JSON</span>
          </span>
        </button>
        <div class="recent" hidden>
          <h2>${icons.clock} Недавние</h2>
          <ul></ul>
        </div>
      </div>
    </section>
  </main>
  <div class="drop-overlay" hidden><div>${icons.upload}<span>Отпустите, чтобы открыть</span></div></div>
  <div class="toasts"></div>
`;

const $ = <T extends HTMLElement = HTMLElement>(sel: string, root: ParentNode = app) => root.querySelector(sel) as T;
const tabsEl = $(".tabs");
const toolbar = $(".toolbar");
const stage = $(".stage");
const welcome = $(".welcome");
const searchbar = $(".searchbar");
const searchInput = $<HTMLInputElement>(".searchbar input");
const searchCount = $(".search-count");
const pagerInput = $<HTMLInputElement>(".pager-input");

// ---------------------------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------------------------

interface Tab {
  id: number;
  source: Source;
  kind: DocKind | null;
  tabEl: HTMLElement;
  pane: HTMLElement;
  host: HTMLElement;
  footer: HTMLElement;
  viewer?: Viewer;
  size?: number;
  query: string;
  saving?: boolean;
}

const tabs: Tab[] = [];
let active: Tab | null = null;
let nextId = 1;

async function openSources(sources: Source[]) {
  let last: Tab | null = null;
  for (const s of sources) last = (await openSource(s)) ?? last;
  if (last) activate(last);
}

async function openSource(source: Source): Promise<Tab | null> {
  if (source.path) {
    const existing = tabs.find((t) => t.source.path && samePath(t.source.path, source.path!));
    if (existing) return existing;
  }
  const kind = kindOf(source.name);
  const legacy = legacyFormat(source.name);
  if (!kind && !legacy) {
    toast(`«${source.name}» — неподдерживаемый формат`);
    return null;
  }

  const tabEl = document.createElement("div");
  tabEl.className = "tab loading";
  tabEl.setAttribute("role", "tab");
  tabEl.title = source.path ?? source.name;
  tabEl.innerHTML = `${fileBadge(kind ?? "", 16)}<span class="tab-name"></span><span class="spinner"></span><button class="tab-close" title="Закрыть (Ctrl+W)">${icons.close}</button>`;
  tabEl.querySelector(".tab-name")!.textContent = source.name;
  tabsEl.appendChild(tabEl);

  const pane = document.createElement("section");
  pane.className = `pane kind-${kind ?? "legacy"}`;
  pane.hidden = true;
  pane.innerHTML = `<div class="doc-host" tabindex="-1"></div><div class="doc-footer" hidden></div>`;
  stage.appendChild(pane);

  const tab: Tab = {
    id: nextId++,
    source,
    kind,
    tabEl,
    pane,
    host: $(".doc-host", pane),
    footer: $(".doc-footer", pane),
    query: "",
  };
  tabs.push(tab);

  tabEl.addEventListener("mousedown", (e) => {
    if (e.button === 1) {
      e.preventDefault();
      requestClose(tab);
    }
  });
  tabEl.addEventListener("click", (e) => {
    if ((e.target as HTMLElement).closest(".tab-close")) requestClose(tab);
    else activate(tab);
  });

  activate(tab);
  if (source.path) remember(source.path, kind);

  if (!kind) {
    showMessage(
      tab,
      icons.alert,
      "Старый формат не поддерживается",
      `${legacy} — это бинарный формат до 2007 года. Откройте файл в Google Docs / Sheets / Slides или сохраните его как .docx / .xlsx / .pptx.`,
    );
    tabEl.classList.remove("loading");
    return tab;
  }

  const loader = document.createElement("div");
  loader.className = "loader";
  loader.innerHTML = `<span class="spinner big"></span><span>Открываю ${escapeHtml(source.name)}…</span>`;
  tab.host.appendChild(loader);

  try {
    const [data, info] = await Promise.all([source.load(), source.path ? fileInfo(source.path) : Promise.resolve(null)]);
    tab.size = info?.size ?? source.size ?? data.byteLength;
    if (!tabs.includes(tab)) return null; // closed while loading
    const ctx: ViewerContext = {
      host: tab.host,
      footer: tab.footer,
      changed: () => {
        tab.footer.hidden = !tab.footer.childElementCount;
        refreshDirty(tab);
        if (tab === active) updateToolbar();
      },
      askPassword: (retry) => askPassword(tab, retry),
      path: source.path,
    };
    // Load each renderer on demand so start-up stays instant.
    let viewer: Viewer;
    switch (kind) {
      case "pdf":
        viewer = await (await import("./viewers/pdf")).openPdf(data, ctx);
        break;
      case "docx":
        viewer = await (await import("./viewers/docx")).openDocx(data, ctx);
        break;
      case "pptx":
        viewer = await (await import("./viewers/pptx")).openPptx(data, ctx);
        break;
      case "xlsx":
      case "csv":
        viewer = await (await import("./viewers/xlsx")).openSpreadsheet(data, ctx, source.name);
        break;
      case "md":
        viewer = await (await import("./viewers/markdown")).openMarkdown(data, ctx);
        break;
      case "txt":
        viewer = await (await import("./viewers/text")).openText(data, ctx);
        break;
      case "json":
        viewer = await (await import("./viewers/json")).openJson(data, ctx);
        break;
    }
    loader.remove();
    if (!tabs.includes(tab)) {
      viewer.destroy();
      return null;
    }
    tab.viewer = viewer;
    ctx.changed();
  } catch (err: any) {
    console.error(err);
    loader.remove();
    const msg = String(err?.message ?? err);
    if (/password/i.test(msg) || err?.name === "PasswordException") {
      showMessage(tab, icons.lock, "Документ защищён паролем", "Не удалось открыть: пароль не введён или неверный.");
    } else if (/encrypt|EncryptionInfo|CDF|ole/i.test(msg)) {
      showMessage(tab, icons.lock, "Файл зашифрован", "Офисные файлы с паролем на открытие не поддерживаются.");
    } else {
      showMessage(tab, icons.alert, "Не удалось открыть файл", msg);
    }
  } finally {
    tabEl.classList.remove("loading");
    if (tab === active) updateToolbar();
  }
  return tab;
}

function showMessage(tab: Tab, icon: string, title: string, text: string) {
  const box = document.createElement("div");
  box.className = "message";
  box.innerHTML = `<div class="message-icon">${icon}</div><h3></h3><p></p>`;
  box.querySelector("h3")!.textContent = title;
  box.querySelector("p")!.textContent = text;
  tab.host.replaceChildren(box);
}

function activate(tab: Tab) {
  if (active && active !== tab) {
    active.query = searchInput.value;
    active.viewer?.search?.clear();
  }
  active = tab;
  for (const t of tabs) {
    const on = t === tab;
    t.pane.hidden = !on;
    t.tabEl.classList.toggle("active", on);
    t.tabEl.setAttribute("aria-selected", String(on));
  }
  welcome.hidden = true;
  tab.tabEl.scrollIntoView({ block: "nearest", inline: "nearest" });
  updateTitle();
  searchInput.value = tab.query;
  searchCount.textContent = "";
  if (!searchbar.hidden && tab.query) runSearch();
  updateToolbar();
  tab.host.focus({ preventScroll: true });
}

function closeTab(tab: Tab) {
  const i = tabs.indexOf(tab);
  if (i < 0) return;
  tab.viewer?.search?.clear();
  tab.viewer?.destroy();
  tab.pane.remove();
  tab.tabEl.remove();
  tabs.splice(i, 1);
  if (active === tab) {
    active = null;
    const next = tabs[i] ?? tabs[i - 1];
    if (next) activate(next);
    else showWelcome();
  }
}

const isDirty = (tab: Tab) => !!tab.viewer?.edit?.dirty();

function updateTitle() {
  if (active) setWindowTitle(`${isDirty(active) ? "● " : ""}${active.source.name} — Glance`);
}

function refreshDirty(tab: Tab) {
  tab.tabEl.classList.toggle("dirty", isDirty(tab));
  if (tab === active) updateTitle();
}

/** Close a tab, first offering to save unsaved edits. */
async function requestClose(tab: Tab) {
  if (modalOpen) return;
  if (isDirty(tab)) {
    activate(tab);
    const choice = await askUnsaved(`Сохранить изменения в «${tab.source.name}»?`);
    if (choice === "cancel" || (choice === "save" && !(await saveTab(tab)))) return;
  }
  closeTab(tab);
}

async function saveTab(tab: Tab): Promise<boolean> {
  const edit = tab.viewer?.edit;
  if (!edit || tab.saving) return false;
  tab.saving = true;
  try {
    const { data, note } = await edit.serialize();
    if (tab.source.path && isTauri) await writeFile(tab.source.path, data);
    else downloadFile(tab.source.name, data);
    edit.markSaved();
    tab.size = data.byteLength;
    if (note) toast(note);
    return true;
  } catch (err) {
    toast(`Не удалось сохранить «${tab.source.name}»: ${err}`);
    return false;
  } finally {
    tab.saving = false;
    refreshDirty(tab);
    if (tab === active) updateToolbar();
  }
}

function showWelcome() {
  active = null;
  welcome.hidden = false;
  toolbar.hidden = true;
  closeSearch();
  setWindowTitle("Glance");
  renderRecent();
}

// ---------------------------------------------------------------------------------------
// Toolbar
// ---------------------------------------------------------------------------------------

function updateToolbar() {
  const tab = active;
  toolbar.hidden = !tab;
  if (!tab) return;
  const v = tab.viewer;
  const pager = $(".pager", toolbar);
  pager.hidden = !v?.pages;
  if (v?.pages) {
    $(".pager-label", toolbar).textContent = v.pages.label;
    if (document.activeElement !== pagerInput) pagerInput.value = String(v.pages.current());
    pagerInput.style.width = `${Math.max(2, String(v.pages.total()).length) + 1.5}ch`;
    $(".pager-total", toolbar).textContent = `из ${v.pages.total()}`;
  }
  const parts = [v?.summary?.(), tab.size ? formatSize(tab.size) : ""].filter(Boolean);
  $(".tb-summary", toolbar).textContent = parts.join(" · ");
  $(".zoom", toolbar).hidden = !v;
  $(".tb-sep", toolbar).hidden = !v;
  $(".zoom-value", toolbar).textContent = v ? `${Math.round(v.zoom * 100)}%` : "";
  const fit = $('[data-act="fit"]', toolbar) as HTMLElement;
  fit.hidden = !v?.fitWidth;
  fit.title = v?.fitTitle ?? "По ширине окна";
  ($('[data-act="search"]', toolbar) as HTMLElement).hidden = !v?.search;
  ($('[data-act="reveal"]', toolbar) as HTMLElement).hidden = !tab.source.path;
  const save = $<HTMLButtonElement>('[data-act="save"]', toolbar);
  save.hidden = !v?.edit;
  save.disabled = !isDirty(tab);
  const actions = $(".actions", toolbar);
  actions.replaceChildren(
    ...(v?.actions ?? []).map((a) => {
      const b = document.createElement("button");
      b.className = "icon-btn";
      b.classList.toggle("on", !!a.active);
      b.setAttribute("aria-pressed", String(!!a.active));
      b.title = a.title;
      b.innerHTML = a.icon;
      b.onclick = () => a.run();
      return b;
    }),
  );
}

toolbar.addEventListener("click", (e) => {
  const act = (e.target as HTMLElement).closest<HTMLElement>("[data-act]")?.dataset.act;
  const v = active?.viewer;
  if (!act || !active) return;
  switch (act) {
    case "prev":
      v?.pages?.goTo(v.pages.current() - 1);
      break;
    case "next":
      v?.pages?.goTo(v.pages.current() + 1);
      break;
    case "zoom-in":
      v?.setZoom(nextZoom(v.zoom, 1));
      break;
    case "zoom-out":
      v?.setZoom(nextZoom(v.zoom, -1));
      break;
    case "zoom-reset":
      v?.setZoom(1);
      break;
    case "fit":
      v?.fitWidth?.();
      break;
    case "search":
      openSearch();
      break;
    case "save":
      saveTab(active);
      break;
    case "reveal":
      if (active.source.path) revealInFolder(active.source.path).catch((err) => toast(String(err)));
      break;
  }
});

pagerInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    const n = parseInt(pagerInput.value, 10);
    if (!isNaN(n)) active?.viewer?.pages?.goTo(n);
    pagerInput.blur();
  } else if (e.key === "Escape") pagerInput.blur();
});
pagerInput.addEventListener("focus", () => pagerInput.select());
pagerInput.addEventListener("blur", () => updateToolbar());

$(".new-tab").addEventListener("click", () => openDialog());
$(".dropzone").addEventListener("click", () => openDialog());

async function openDialog() {
  try {
    openSources(await pickFiles());
  } catch (e) {
    toast(String(e));
  }
}

// ---------------------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------------------

let searchSeq = 0;
function openSearch() {
  if (!active?.viewer?.search) return;
  searchbar.hidden = false;
  searchInput.focus();
  searchInput.select();
}
function closeSearch() {
  searchbar.hidden = true;
  active?.viewer?.search?.clear();
  searchCount.textContent = "";
}
let searchTotal = 0;
async function runSearch() {
  const s = active?.viewer?.search;
  if (!s) return;
  const seq = ++searchSeq;
  const q = searchInput.value.trim();
  if (!q) {
    s.clear();
    searchCount.textContent = "";
    searchTotal = 0;
    return;
  }
  searchCount.textContent = "…";
  const n = await s.find(q);
  if (seq !== searchSeq) return;
  searchTotal = n;
  if (n) {
    const i = s.step(1);
    searchCount.textContent = `${i} из ${n}`;
  } else searchCount.textContent = "Нет совпадений";
  searchbar.classList.toggle("nomatch", !n);
}
function stepSearch(dir: 1 | -1) {
  const s = active?.viewer?.search;
  if (!s || !searchTotal) return;
  const i = s.step(dir);
  searchCount.textContent = `${i} из ${searchTotal}`;
}
let searchTimer = 0;
searchInput.addEventListener("input", () => {
  clearTimeout(searchTimer);
  searchTimer = window.setTimeout(runSearch, 250);
});
searchInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    clearTimeout(searchTimer);
    if (searchCount.textContent === "" || searchCount.textContent === "…") runSearch();
    else stepSearch(e.shiftKey ? -1 : 1);
  } else if (e.key === "Escape") {
    closeSearch();
    active?.host.focus();
  }
});
searchbar.addEventListener("click", (e) => {
  const act = (e.target as HTMLElement).closest<HTMLElement>("[data-act]")?.dataset.act;
  if (act === "search-next") stepSearch(1);
  else if (act === "search-prev") stepSearch(-1);
  else if (act === "search-close") closeSearch();
});

// ---------------------------------------------------------------------------------------
// Password prompt (PDF)
// ---------------------------------------------------------------------------------------

function askPassword(tab: Tab, retry: boolean): Promise<string | null> {
  return new Promise((resolve) => {
    const box = document.createElement("form");
    box.className = "message password";
    box.innerHTML = `<div class="message-icon">${icons.lock}</div>
      <h3>Документ защищён паролем</h3>
      <p>${retry ? "Неверный пароль, попробуйте ещё раз." : "Введите пароль, чтобы открыть файл."}</p>
      <div class="pw-row"><input type="password" autocomplete="off" placeholder="Пароль" /><button type="submit" class="primary">Открыть</button></div>
      <button type="button" class="link">Отмена</button>`;
    tab.host.querySelector(".loader")?.setAttribute("hidden", "");
    tab.host.appendChild(box);
    const input = box.querySelector("input")!;
    setTimeout(() => input.focus(), 50);
    box.addEventListener("submit", (e) => {
      e.preventDefault();
      box.remove();
      tab.host.querySelector(".loader")?.removeAttribute("hidden");
      resolve(input.value);
    });
    box.querySelector(".link")!.addEventListener("click", () => {
      box.remove();
      resolve(null);
    });
  });
}

// ---------------------------------------------------------------------------------------
// Unsaved changes prompt
// ---------------------------------------------------------------------------------------

let modalOpen = false;

function askUnsaved(question: string): Promise<"save" | "discard" | "cancel"> {
  modalOpen = true;
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.innerHTML = `<div class="modal" role="alertdialog" aria-modal="true">
        <h3></h3>
        <p>Если не сохранить, изменения будут потеряны.</p>
        <div class="modal-buttons">
          <button type="button" class="primary" data-choice="save">Сохранить</button>
          <button type="button" class="secondary" data-choice="discard">Не сохранять</button>
          <button type="button" class="secondary" data-choice="cancel">Отмена</button>
        </div>
      </div>`;
    overlay.querySelector("h3")!.textContent = question;
    const done = (choice: "save" | "discard" | "cancel") => {
      overlay.remove();
      modalOpen = false;
      resolve(choice);
    };
    overlay.addEventListener("click", (e) => {
      const choice = (e.target as HTMLElement).closest<HTMLElement>("[data-choice]")?.dataset.choice;
      if (choice) done(choice as "save" | "discard" | "cancel");
      else if (e.target === overlay) done("cancel");
    });
    overlay.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        e.preventDefault();
        done("cancel");
      }
    });
    app.appendChild(overlay);
    overlay.querySelector<HTMLElement>(".primary")!.focus();
  });
}

/** Window close: resolves true when it is fine to close (saved, discarded or nothing to save). */
async function confirmQuit(): Promise<boolean> {
  if (modalOpen) return false;
  const dirty = tabs.filter(isDirty);
  if (!dirty.length) return true;
  activate(dirty[0]);
  const question =
    dirty.length === 1 ? `Сохранить изменения в «${dirty[0].source.name}»?` : `Сохранить изменения в файлах (${dirty.length})?`;
  const choice = await askUnsaved(question);
  if (choice === "cancel") return false;
  if (choice === "discard") return true;
  for (const t of dirty) if (!(await saveTab(t))) return false;
  return true;
}

// ---------------------------------------------------------------------------------------
// Recent files
// ---------------------------------------------------------------------------------------

interface Recent {
  path: string;
  name: string;
  kind: string;
  time: number;
}
const RECENT_KEY = "glance.recent";

function loadRecent(): Recent[] {
  try {
    return JSON.parse(localStorage.getItem(RECENT_KEY) || "[]");
  } catch {
    return [];
  }
}
function saveRecent(list: Recent[]) {
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0, 15)));
  } catch {
    /* storage unavailable */
  }
}
function remember(path: string, kind: string | null) {
  const list = loadRecent().filter((r) => !samePath(r.path, path));
  list.unshift({ path, name: path.split(/[\\/]/).pop() || path, kind: kind ?? "", time: Date.now() });
  saveRecent(list);
}
function renderRecent() {
  const box = $(".recent");
  const list = loadRecent();
  box.hidden = !isTauri || !list.length;
  const ul = $("ul", box);
  ul.replaceChildren(
    ...list.map((r) => {
      const li = document.createElement("li");
      li.innerHTML = `<button class="recent-item">${fileBadge(r.kind, 20)}<span class="ri-text"><span class="ri-name"></span><span class="ri-path"></span></span><span class="ri-time"></span></button>`;
      $(".ri-name", li).textContent = r.name;
      $(".ri-path", li).textContent = r.path.replace(/[\\/][^\\/]*$/, "");
      $(".ri-time", li).textContent = relTime(r.time);
      li.querySelector("button")!.onclick = async () => {
        const info = await fileInfo(r.path);
        if (!info) {
          toast(`Файл не найден: ${r.name}`);
          saveRecent(loadRecent().filter((x) => x.path !== r.path));
          renderRecent();
          return;
        }
        openSources([sourceFromPath(r.path)]);
      };
      return li;
    }),
  );
}

// ---------------------------------------------------------------------------------------
// Drag & drop, keyboard, wheel zoom
// ---------------------------------------------------------------------------------------

const dropOverlay = $(".drop-overlay");
onFileDrop((state, sources) => {
  dropOverlay.hidden = state !== "over";
  if (state === "drop" && sources.length) openSources(sources);
});

window.addEventListener("keydown", (e) => {
  if (modalOpen) {
    if (e.key !== "Tab" && e.key !== "Enter" && e.key !== " " && e.key !== "Escape") e.preventDefault();
    return;
  }
  const mod = e.ctrlKey || e.metaKey;
  const v = active?.viewer;
  // e.code keeps Ctrl+S / Ctrl+E working with the Russian keyboard layout.
  const k = e.code.startsWith("Key") ? e.code.slice(3).toLowerCase() : e.key.toLowerCase();
  if (mod && (k === "o" || k === "t")) {
    e.preventDefault();
    openDialog();
  } else if (mod && k === "w") {
    e.preventDefault();
    if (active) requestClose(active);
  } else if (mod && k === "s") {
    e.preventDefault();
    if (active && isDirty(active)) saveTab(active);
  } else if (mod && k === "e") {
    e.preventDefault();
    v?.actions?.find((a) => a.id === "edit")?.run();
  } else if (mod && e.key === "Tab") {
    e.preventDefault();
    if (tabs.length > 1 && active) {
      const i = tabs.indexOf(active);
      activate(tabs[(i + (e.shiftKey ? -1 : 1) + tabs.length) % tabs.length]);
    }
  } else if (mod && /^[1-9]$/.test(e.key)) {
    e.preventDefault();
    const t = e.key === "9" ? tabs[tabs.length - 1] : tabs[Number(e.key) - 1];
    if (t) activate(t);
  } else if (mod && k === "f") {
    e.preventDefault();
    openSearch();
  } else if (e.key === "F3") {
    e.preventDefault();
    if (searchbar.hidden) openSearch();
    else stepSearch(e.shiftKey ? -1 : 1);
  } else if (e.key === "Escape" && !searchbar.hidden) {
    closeSearch();
  } else if (mod && (e.key === "=" || e.key === "+")) {
    e.preventDefault();
    v?.setZoom(nextZoom(v.zoom, 1));
  } else if (mod && e.key === "-") {
    e.preventDefault();
    v?.setZoom(nextZoom(v.zoom, -1));
  } else if (mod && e.key === "0") {
    e.preventDefault();
    v?.setZoom(1);
  } else if (e.key === "F5") {
    e.preventDefault(); // never reload the app window
    v?.actions?.find((a) => a.id === "slideshow")?.run();
  } else if (mod && (k === "r" || k === "p")) {
    e.preventDefault();
  } else if (!mod && v?.pages && (e.key === "PageDown" || e.key === "PageUp") && active?.kind === "pptx") {
    e.preventDefault();
    v.pages.goTo(v.pages.current() + (e.key === "PageDown" ? 1 : -1));
  }
});

let wheelAcc = 0;
stage.addEventListener(
  "wheel",
  (e) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    e.preventDefault();
    const v = active?.viewer;
    if (!v) return;
    wheelAcc += e.deltaY;
    if (Math.abs(wheelAcc) >= 60) {
      v.setZoom(nextZoom(v.zoom, wheelAcc < 0 ? 1 : -1));
      wheelAcc = 0;
    }
  },
  { passive: false },
);

// Links inside documents must never navigate the app window itself.
document.addEventListener(
  "click",
  (e) => {
    const a = (e.target as HTMLElement).closest?.("a[href]") as HTMLAnchorElement | null;
    if (!a) return;
    const href = a.getAttribute("href") || "";
    if (href.startsWith("#")) return; // in-document bookmark
    e.preventDefault();
    if (a.isContentEditable && !(e.ctrlKey || e.metaKey)) return; // editing: a click places the caret
    if (/^(https?:|mailto:|tel:)/i.test(href)) openExternal(href).catch((err) => toast(String(err)));
    else if (active?.source.path && !/^(javascript|data|blob|about):/i.test(href)) openLinkedFile(resolvePath(active.source.path, href));
  },
  true,
);

/** Relative links between local documents (e.g. README.md → docs/setup.md) open as tabs. */
async function openLinkedFile(path: string) {
  if (!(await fileInfo(path))) {
    toast(`Файл не найден: ${path}`);
    return;
  }
  openSources([sourceFromPath(path)]);
}

// The WebView's own context menu (Reload, Save as…) makes no sense here; keep it only for copying text.
window.addEventListener("contextmenu", (e) => {
  const sel = window.getSelection();
  const editable = (e.target as HTMLElement).closest("input, textarea");
  if (!editable && (!sel || sel.isCollapsed)) e.preventDefault();
});

// ---------------------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------------------

function toast(text: string) {
  const el = document.createElement("div");
  el.className = "toast";
  el.textContent = text;
  $(".toasts").appendChild(el);
  setTimeout(() => el.classList.add("out"), 3800);
  setTimeout(() => el.remove(), 4300);
}

function samePath(a: string, b: string) {
  const norm = (p: string) => p.replace(/\\/g, "/").toLowerCase();
  return norm(a) === norm(b);
}

function formatSize(n: number) {
  if (n < 1024) return `${n} Б`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} КБ`;
  return `${(n / 1024 / 1024).toFixed(1).replace(".", ",")} МБ`;
}

function relTime(t: number) {
  const d = (Date.now() - t) / 1000;
  if (d < 60) return "только что";
  if (d < 3600) return `${Math.floor(d / 60)} мин назад`;
  if (d < 86400) return `${Math.floor(d / 3600)} ч назад`;
  if (d < 86400 * 7) return `${Math.floor(d / 86400)} дн назад`;
  return new Date(t).toLocaleDateString("ru-RU", { day: "numeric", month: "short" });
}

function escapeHtml(s: string) {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
}

// ---------------------------------------------------------------------------------------
// Start-up
// ---------------------------------------------------------------------------------------

showWelcome();
onCloseRequested(confirmQuit, () => tabs.some(isDirty));
onOpenFilesEvent((paths) => openSources(paths.map(sourceFromPath)));
takeStartupFiles().then((paths) => {
  if (paths.length) openSources(paths.map(sourceFromPath));
});

// Test hook: lets automated UI tests open files without a native dialog.
(window as any).__glanceOpen = (files: File[]) => openSources(files.map((f) => ({ name: f.name, size: f.size, load: () => f.arrayBuffer() })));
