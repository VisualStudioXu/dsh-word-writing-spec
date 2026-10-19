/**
 * Minimal namespace-aware XML DOM for Office Open XML parts.
 *
 * Why not a real XML library: this plugin must work inside the harness
 * process with zero third-party dependencies, and OOXML parts are a narrow
 * subset (elements, attributes, text, comments, CDATA, processing
 * instructions, self-closing tags, no DTD). A full DOM would also risk
 * rewriting prefixes or dropping ignorable attributes, which is exactly what
 * corrupts a Word file.
 *
 * Design rules that keep round-trips safe:
 *  - prefixes are preserved verbatim (w:, a:, mc:, r:, ...), never renamed;
 *  - attribute order is preserved;
 *  - only elements we intentionally touch are re-serialized; untouched parts
 *    are written back as the original text when they were not parsed;
 *  - text is escaped on serialize, and already-escaped entities are decoded on
 *    parse, so no double escaping.
 */

const VOID_OK = new Set(['w:br', 'w:tab', 'w:cr']);

export class XmlText {
  constructor(text) {
    this.type = 'text';
    this.text = text;
  }
}

export class XmlElement {
  constructor(name) {
    this.type = 'element';
    this.name = name;
    /** @type {Array<{name: string, value: string}>} */
    this.attrs = [];
    /** @type {Array<XmlElement|XmlText>} */
    this.children = [];
    this.parent = null;
  }

  get attrMap() {
    const m = new Map();
    for (const a of this.attrs) m.set(a.name, a.value);
    return m;
  }

  localName() {
    const i = this.name.indexOf(':');
    return i < 0 ? this.name : this.name.slice(i + 1);
  }

  attr(name) {
    const hit = this.attrs.find((a) => a.name === name);
    return hit ? hit.value : undefined;
  }

  setAttr(name, value) {
    const hit = this.attrs.find((a) => a.name === name);
    if (hit) hit.value = value;
    else this.attrs.push({ name, value });
    return this;
  }

  removeAttr(name) {
    const i = this.attrs.findIndex((a) => a.name === name);
    if (i >= 0) this.attrs.splice(i, 1);
    return this;
  }

  append(child) {
    if (child && typeof child === 'object') child.parent = this;
    this.children.push(child);
    return child;
  }

  insertBefore(child, ref) {
    const i = ref ? this.children.indexOf(ref) : -1;
    if (child && typeof child === 'object') child.parent = this;
    if (i < 0) this.children.push(child);
    else this.children.splice(i, 0, child);
    return child;
  }

  removeChild(child) {
    const i = this.children.indexOf(child);
    if (i >= 0) this.children.splice(i, 1);
    return child;
  }

  elements() {
    return this.children.filter((c) => c.type === 'element');
  }

  child(name) {
    return this.children.find((c) => c.type === 'element' && c.name === name) ?? null;
  }

  childrenNamed(name) {
    return this.children.filter((c) => c.type === 'element' && c.name === name);
  }

  /** Depth-first search for the first descendant (or self) matching a predicate. */
  find(pred) {
    if (pred(this)) return this;
    for (const c of this.children) {
      if (c.type !== 'element') continue;
      const hit = c.find(pred);
      if (hit) return hit;
    }
    return null;
  }

  findAll(pred, out = []) {
    if (pred(this)) out.push(this);
    for (const c of this.children) {
      if (c.type !== 'element') continue;
      c.findAll(pred, out);
    }
    return out;
  }

  get text() {
    let s = '';
    for (const c of this.children) {
      if (c.type === 'text') s += c.text;
      else if (c.type === 'element') {
        if (c.name === 'w:t') s += c.textContent;
        else if (c.name === 'w:tab') s += '\t';
        else if (c.name === 'w:br' || c.name === 'w:cr') s += '\n';
        else s += c.text;
      }
    }
    return s;
  }

  get textContent() {
    let s = '';
    for (const c of this.children) {
      if (c.type === 'text') s += c.text;
      else if (c.type === 'element') s += c.textContent;
    }
    return s;
  }
}

/* ------------------------------------------------------------------ */
/* Parsing                                                             */
/* ------------------------------------------------------------------ */

function decodeEntities(s) {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, '&');
}

export function escapeText(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function escapeAttr(s) {
  return escapeText(s).replace(/"/g, '&quot;');
}

/**
 * Parse an XML document into a document node.
 * @param {string} src
 * @returns {XmlElement} a synthetic '#document' element
 */
export function parseXml(src) {
  const doc = new XmlElement('#document');
  let cur = doc;
  let i = 0;
  const n = src.length;

  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt < 0) {
      const tail = src.slice(i);
      if (tail.trim()) cur.append(new XmlText(decodeEntities(tail)));
      break;
    }
    if (lt > i) {
      const chunk = src.slice(i, lt);
      if (chunk.trim()) cur.append(new XmlText(decodeEntities(chunk)));
    }

    if (src.startsWith('<!--', lt)) {
      const end = src.indexOf('-->', lt + 4);
      if (end < 0) throw new Error('XML 解析失败：注释未闭合');
      i = end + 3;
      continue;
    }
    if (src.startsWith('<![CDATA[', lt)) {
      const end = src.indexOf(']]>', lt + 9);
      if (end < 0) throw new Error('XML 解析失败：CDATA 未闭合');
      cur.append(new XmlText(src.slice(lt + 9, end)));
      i = end + 3;
      continue;
    }
    if (src.startsWith('<?', lt) || src.startsWith('<!', lt)) {
      const end = src.indexOf('>', lt);
      if (end < 0) throw new Error('XML 解析失败：声明未闭合');
      i = end + 1;
      continue;
    }

    // closing tag
    if (src[lt + 1] === '/') {
      const end = src.indexOf('>', lt);
      if (end < 0) throw new Error(`XML 解析失败：闭合标签未结束（偏移 ${lt}）`);
      const closeName = src.slice(lt + 2, end).trim();
      if (cur === doc || cur.name !== closeName) {
        throw new Error(
          `XML 解析失败：闭合标签 </${closeName}> 与当前元素 <${cur === doc ? '#document' : cur.name}> 不匹配（偏移 ${lt}）`,
        );
      }
      cur = cur.parent ?? doc;
      i = end + 1;
      continue;
    }

    // element
    let j = lt + 1;
    let name = '';
    while (j < n && !/[\s/>]/.test(src[j])) name += src[j++];
    if (!name) throw new Error(`XML 解析失败：标签名为空（偏移 ${lt}）`);

    const el = new XmlElement(name);
    // attributes
    for (;;) {
      while (j < n && /\s/.test(src[j])) j++;
      if (j >= n) throw new Error(`XML 解析失败：<${name}> 未闭合`);
      if (src[j] === '/' && src[j + 1] === '>') {
        j += 2;
        cur.append(el);
        i = j;
        break;
      }
      if (src[j] === '>') {
        j += 1;
        cur.append(el);
        cur = el;
        i = j;
        break;
      }
      let an = '';
      while (j < n && !/[\s=/>]/.test(src[j])) an += src[j++];
      while (j < n && /\s/.test(src[j])) j++;
      if (src[j] !== '=') throw new Error(`XML 解析失败：属性 ${an} 缺少取值（<${name}>）`);
      j++;
      while (j < n && /\s/.test(src[j])) j++;
      const quote = src[j];
      if (quote !== '"' && quote !== "'") throw new Error(`XML 解析失败：属性 ${an} 的引号不合法`);
      const end = src.indexOf(quote, j + 1);
      if (end < 0) throw new Error(`XML 解析失败：属性 ${an} 未闭合`);
      el.attrs.push({ name: an, value: decodeEntities(src.slice(j + 1, end)) });
      j = end + 1;
    }
    continue;
  }

  if (cur !== doc) throw new Error(`XML 解析失败：<${cur.name}> 未闭合`);
  return doc;
}

/* ------------------------------------------------------------------ */
/* Serializing                                                         */
/* ------------------------------------------------------------------ */

export function serializeXml(node) {
  if (node.type === 'text') return escapeText(node.text);
  let s = '';
  for (const c of node.children) s += serializeXml(c);
  return s;
}

export function serializeElement(el) {
  if (el.type === 'text') return escapeText(el.text);
  let head = `<${el.name}`;
  for (const a of el.attrs) head += ` ${a.name}="${escapeAttr(a.value)}"`;
  if (!el.children.length) return `${head}/>`;
  let body = '';
  for (const c of el.children) body += serializeElement(c);
  return `${head}>${body}</${el.name}>`;
}

/** Serialize the whole document including the XML declaration and root attrs. */
export function serializeDocument(doc, { standalone = true } = {}) {
  let s = '<?xml version="1.0" encoding="UTF-8"' + (standalone ? ' standalone="yes"' : '') + '?>\r\n';
  for (const c of doc.children) {
    if (c.type === 'text') continue;
    s += serializeElement(c);
  }
  return s;
}

/** True when the part already carried its own XML declaration. */
export function hasDeclaration(src) {
  return /^\uFEFF?\s*<\?xml/.test(src);
}

/* ------------------------------------------------------------------ */
/* Small helpers                                                       */
/* ------------------------------------------------------------------ */

export function makeElement(name, attrs = {}) {
  const el = new XmlElement(name);
  for (const [k, v] of Object.entries(attrs)) el.attrs.push({ name: k, value: String(v) });
  return el;
}

/** First element child, tolerating any prefix (w:sectPr vs sectPr). */
export function child(el, localName) {
  if (!el) return null;
  return el.children.find((c) => c.type === 'element' && c.localName() === localName) ?? null;
}

export function childrenNamed(el, localName) {
  if (!el) return [];
  return el.children.filter((c) => c.type === 'element' && c.localName() === localName);
}

/** BFS/DFS by local name, prefix-agnostic. */
export function findAllByLocal(el, localName) {
  if (!el) return [];
  return el.findAll((e) => e.localName() === localName);
}

export function findFirstByLocal(el, localName) {
  return findAllByLocal(el, localName)[0] ?? null;
}
