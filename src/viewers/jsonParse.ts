// JSON parser that keeps source positions, so the tree can edit the text in place and
// numbers are shown exactly as written (JSON.parse would round big integers).

export type JType = "object" | "array" | "string" | "number" | "boolean" | "null";

export interface JNode {
  type: JType;
  start: number;
  end: number;
  /** Decoded value of strings. */
  value?: string;
  entries?: JEntry[];
}

export interface JEntry {
  key: string | number;
  /** Source range of an object key (including quotes). */
  keyStart?: number;
  keyEnd?: number;
  node: JNode;
}

export class JsonError extends Error {
  constructor(
    message: string,
    readonly pos: number,
  ) {
    super(message);
  }
}

const NUMBER = /-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/y;

export function parseJson(text: string): JNode {
  let i = 0;
  const n = text.length;

  const ws = () => {
    while (i < n) {
      const c = text.charCodeAt(i);
      if (c === 32 || c === 9 || c === 10 || c === 13) i++;
      else break;
    }
  };
  const fail = (msg: string, at = i): never => {
    throw new JsonError(msg, at);
  };
  const describe = () => (i >= n ? "конец файла" : `«${text[i]}»`);

  const str = (): { value: string; start: number; end: number } => {
    const start = i++;
    let escaped = false;
    while (i < n) {
      const c = text.charCodeAt(i);
      if (c === 34) break;
      if (c === 92) {
        escaped = true;
        const e = text[i + 1];
        if (e === "u") {
          if (!/^[0-9a-fA-F]{4}$/.test(text.slice(i + 2, i + 6))) fail("неверная последовательность \\u", i);
          i += 6;
        } else if (e !== undefined && '"\\/bfnrt'.includes(e)) i += 2;
        else fail("неверный escape-символ", i);
        continue;
      }
      if (c < 32) fail(c === 10 ? "строка не закрыта кавычкой" : "управляющий символ внутри строки");
      i++;
    }
    if (i >= n) fail("строка не закрыта кавычкой", start);
    i++;
    const raw = text.slice(start, i);
    return { value: escaped ? JSON.parse(raw) : raw.slice(1, -1), start, end: i };
  };

  const value = (depth: number): JNode => {
    if (depth > 5000) fail("слишком глубокая вложенность");
    ws();
    const start = i;
    const c = text[i];
    if (c === "{") {
      i++;
      const entries: JEntry[] = [];
      ws();
      if (text[i] === "}") {
        i++;
        return { type: "object", start, end: i, entries };
      }
      for (;;) {
        ws();
        if (text[i] !== '"') fail(text[i] === "}" ? "лишняя запятая перед «}»" : `ожидалось имя свойства в кавычках, а не ${describe()}`);
        const k = str();
        ws();
        if (text[i] !== ":") fail(`ожидалось «:», а не ${describe()}`);
        i++;
        const node = value(depth + 1);
        entries.push({ key: k.value, keyStart: k.start, keyEnd: k.end, node });
        ws();
        if (text[i] === ",") {
          i++;
          continue;
        }
        if (text[i] === "}") {
          i++;
          return { type: "object", start, end: i, entries };
        }
        fail(`ожидалось «,» или «}», а не ${describe()}`);
      }
    }
    if (c === "[") {
      i++;
      const entries: JEntry[] = [];
      ws();
      if (text[i] === "]") {
        i++;
        return { type: "array", start, end: i, entries };
      }
      for (;;) {
        ws();
        if (text[i] === "]") fail("лишняя запятая перед «]»");
        entries.push({ key: entries.length, node: value(depth + 1) });
        ws();
        if (text[i] === ",") {
          i++;
          continue;
        }
        if (text[i] === "]") {
          i++;
          return { type: "array", start, end: i, entries };
        }
        fail(`ожидалось «,» или «]», а не ${describe()}`);
      }
    }
    if (c === '"') {
      const s = str();
      return { type: "string", start, end: s.end, value: s.value };
    }
    for (const [word, type] of [
      ["true", "boolean"],
      ["false", "boolean"],
      ["null", "null"],
    ] as const) {
      if (text.startsWith(word, i)) {
        i += word.length;
        return { type, start, end: i };
      }
    }
    NUMBER.lastIndex = i;
    if (NUMBER.test(text)) {
      i = NUMBER.lastIndex;
      return { type: "number", start, end: i };
    }
    if (i >= n) fail("неожиданный конец файла");
    if (c === "'") fail("строки в JSON берутся в двойные кавычки");
    if (c === "/") fail("комментарии в JSON не допускаются");
    return fail(`неожиданный символ ${describe()}`);
  };

  const root = value(0);
  ws();
  if (i < n) fail(`лишние символы после конца данных: ${describe()}`);
  return root;
}

/** 1-based line and column of a text offset. */
export function lineCol(text: string, pos: number): { line: number; col: number } {
  let line = 1;
  let last = -1;
  for (let k = text.indexOf("\n"); k >= 0 && k < pos; k = text.indexOf("\n", k + 1)) {
    line++;
    last = k;
  }
  return { line, col: pos - last };
}

/** Re-indents the document; strings and numbers keep their exact source spelling. */
export function formatJson(text: string, root: JNode, indent: string): string {
  const out: string[] = [];
  const write = (node: JNode, level: number) => {
    if (node.type !== "object" && node.type !== "array") {
      out.push(text.slice(node.start, node.end));
      return;
    }
    const [open, close] = node.type === "object" ? ["{", "}"] : ["[", "]"];
    const entries = node.entries!;
    if (!entries.length) {
      out.push(open + close);
      return;
    }
    const pad = indent ? "\n" + indent.repeat(level + 1) : "";
    out.push(open);
    entries.forEach((e, k) => {
      if (k) out.push(",");
      out.push(pad);
      if (node.type === "object") out.push(text.slice(e.keyStart!, e.keyEnd!) + (indent ? ": " : ":"));
      write(e.node, level + 1);
    });
    out.push((indent ? "\n" + indent.repeat(level) : "") + close);
  };
  write(root, 0);
  return out.join("");
}

/** Indentation used by the document (first indented line), two spaces by default. */
export function detectIndent(text: string): string {
  const m = text.match(/\n([ \t]+)\S/);
  if (!m) return "  ";
  return m[1][0] === "\t" ? "\t" : " ".repeat(Math.min(8, m[1].length));
}

// ---- paths & labels shared by the tree and the graph ----------------------------------

/** Paths are "/"-joined keys (JSON Pointer escaping), "" is the root. */
export function childPath(parent: string, key: string | number) {
  return `${parent}/${String(key).replace(/~/g, "~0").replace(/\//g, "~1")}`;
}
export function splitPath(path: string): string[] {
  return path === "" ? [] : path.slice(1).split("/").map((s) => s.replace(/~1/g, "/").replace(/~0/g, "~"));
}
export function parentOf(path: string) {
  return path.slice(0, path.lastIndexOf("/"));
}

/** Keys that are not plain identifiers are quoted, so "first name" reads as one key. */
export function keyLabel(key: string | number) {
  return typeof key === "number" || /^[\p{L}_$][\p{L}\p{N}_$-]*$/u.test(key) ? String(key) : JSON.stringify(key);
}

/** "3 ключа" / "12 элементов". */
export function countLabel(node: JNode) {
  const n = node.entries?.length ?? 0;
  return node.type === "object" ? plural(n, "ключ", "ключа", "ключей") : plural(n, "элемент", "элемента", "элементов");
}

export function plural(n: number, one: string, few: string, many: string) {
  const m10 = n % 10;
  const m100 = n % 100;
  const w = m10 === 1 && m100 !== 11 ? one : m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14) ? few : many;
  return `${n.toLocaleString("ru-RU")} ${w}`;
}
