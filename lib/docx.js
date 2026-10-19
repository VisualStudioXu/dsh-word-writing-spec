/**
 * DOCX model: open a .docx container, resolve its style inheritance chains,
 * and expose the facts a format audit needs (sections, paragraphs, runs,
 * tables, headers/footers).
 *
 * The important subtlety this module exists for: Word formatting is a four
 * level cascade — docDefaults → table style → paragraph style chain (basedOn)
 * → direct paragraph/run formatting. Reading only the run's own rPr (what most
 * naive scripts do) reports "no font set" for a document that actually renders
 * in the theme font, which is exactly the trap found in the source spec file
 * (ascii=Times New Roman only, eastAsia inherited from a theme with no East
 * Asian font). So everything here resolves the *effective* value, and records
 * *where* it came from (`source`) so a fix report can say which layer to patch.
 */

import fs from 'node:fs';
import { parseXml, child, childrenNamed, findAllByLocal, findFirstByLocal, serializeDocument } from './xml.js';
import { readZip, writeZip } from './zip.js';
import { effectiveDpi, imagePixelSize } from './image-size.js';

const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/* ------------------------------------------------------------------ */
/* Units                                                               */
/* ------------------------------------------------------------------ */

export const TWIP_PER_CM = 566.9291338582677;
export const TWIP_PER_PT = 20;
export const EMU_PER_TWIP = 635;

export const twipToCm = (t) => (t == null ? null : t / TWIP_PER_CM);
export const twipToPt = (t) => (t == null ? null : t / TWIP_PER_PT);
export const ptToTwip = (p) => Math.round(p * TWIP_PER_PT);
export const cmToTwip = (c) => Math.round(c * TWIP_PER_CM);

export const round = (v, d = 2) => (v == null ? null : Math.round(v * 10 ** d) / 10 ** d);

/* ------------------------------------------------------------------ */
/* Zoom chain helpers                                                  */
/* ------------------------------------------------------------------ */

/** w:b / w:sz / font keys use on/off semantics where a missing val means true. */
function onOff(el) {
  if (!el) return undefined;
  const v = el.attr('w:val');
  if (v == null) return true;
  return !(v === '0' || v === 'false' || v === 'off');
}

function halfPtToPt(sz) {
  if (sz == null) return null;
  const n = Number(sz);
  return Number.isFinite(n) ? n / 2 : null;
}

/* ------------------------------------------------------------------ */
/* Style sheet                                                         */
/* ------------------------------------------------------------------ */

class StyleSheet {
  constructor(stylesDoc, themeFonts) {
    this.themeFonts = themeFonts;
    this.byId = new Map();
    this.docDefaultsRPr = null;
    this.docDefaultsPPr = null;
    if (!stylesDoc) return;
    const root = stylesDoc.elements()[0];
    if (!root) return;
    const dd = child(root, 'docDefaults');
    if (dd) {
      const rPrDefault = child(dd, 'rPrDefault');
      const pPrDefault = child(dd, 'pPrDefault');
      this.docDefaultsRPr = rPrDefault ? child(rPrDefault, 'rPr') : null;
      this.docDefaultsPPr = pPrDefault ? child(pPrDefault, 'pPr') : null;
    }
    for (const s of childrenNamed(root, 'style')) {
      const id = s.attr('w:styleId');
      if (!id) continue;
      this.byId.set(id, {
        el: s,
        id,
        type: s.attr('w:type') ?? 'paragraph',
        name: child(s, 'name')?.attr('w:val') ?? id,
        basedOn: child(s, 'basedOn')?.attr('w:val') ?? null,
        next: child(s, 'next')?.attr('w:val') ?? null,
        rPr: child(s, 'rPr'),
        pPr: child(s, 'pPr'),
        default: boolAttr(s, 'w:default'),
        qFormat: !!child(s, 'qFormat'),
        uiPriority: child(s, 'uiPriority')?.attr('w:val'),
      });
    }
  }

  get(id) {
    return id ? this.byId.get(id) ?? null : null;
  }

  /** Style chain from base to leaf; guards against basedOn cycles. */
  chain(id) {
    const out = [];
    const seen = new Set();
    let cur = this.get(id);
    while (cur && !seen.has(cur.id)) {
      seen.add(cur.id);
      out.unshift(cur);
      cur = cur.basedOn ? this.get(cur.basedOn) : null;
    }
    return out;
  }

  /** Merge rPr across the chain, leaf wins. */
  effectiveRPr(styleId, directRPr) {
    const layers = [];
    if (this.docDefaultsRPr) layers.push({ el: this.docDefaultsRPr, from: 'docDefaults' });
    for (const s of this.chain(styleId)) {
      if (s.rPr) layers.push({ el: s.rPr, from: `style:${s.id}` });
    }
    if (directRPr) layers.push({ el: directRPr, from: 'direct' });
    return mergeRPr(layers, this.themeFonts);
  }

  /** Merge pPr across the chain, leaf wins. */
  effectivePPr(styleId, directPPr) {
    const layers = [];
    if (this.docDefaultsPPr) layers.push({ el: this.docDefaultsPPr, from: 'docDefaults' });
    for (const s of this.chain(styleId)) {
      if (s.pPr) layers.push({ el: s.pPr, from: `style:${s.id}` });
    }
    if (directPPr) layers.push({ el: directPPr, from: 'direct' });
    return mergePPr(layers);
  }

  /** Does the style chain declare an outline level? (TOC / navigation) */
  outlineLevel(styleId) {
    for (const s of this.chain(styleId).reverse()) {
      const ol = s.pPr ? child(s.pPr, 'outlineLvl') : null;
      if (ol) {
        const v = Number(ol.attr('w:val'));
        if (Number.isFinite(v)) return v;
      }
    }
    return null;
  }
}

function boolAttr(el, name) {
  const v = el.attr(name);
  return v != null && !(v === '0' || v === 'false');
}

/**
 * Resolve fonts across layers, including the theme indirection that makes
 * "ascii font looks right, Chinese renders in Cambria" possible.
 */
function mergeRPr(layers, themeFonts) {
  const out = {
    fontAscii: null,
    fontAsciiTheme: null,
    fontEastAsia: null,
    fontEastAsiaTheme: null,
    fontHAnsi: null,
    fontCS: null,
    sizePt: null,
    sizeCsPt: null,
    bold: null,
    italic: null,
    underline: null,
    color: null,
    /** Which layer supplied each of the four audited properties. */
    source: {},
  };

  for (const { el, from } of layers) {
    if (!el) continue;
    const fonts = child(el, 'rFonts');
    if (fonts) {
      const a = fonts.attr('w:ascii');
      const at = fonts.attr('w:asciiTheme');
      const e = fonts.attr('w:eastAsia');
      const et = fonts.attr('w:eastAsiaTheme');
      const h = fonts.attr('w:hAnsi');
      if (a != null) { out.fontAscii = a; delete out.fontAsciiTheme; out.source.fontAscii = from; }
      if (at != null) { out.fontAsciiTheme = at; out.source.fontAsciiTheme = from; }
      if (e != null) { out.fontEastAsia = e; delete out.fontEastAsiaTheme; out.source.fontEastAsia = from; }
      if (et != null) { out.fontEastAsiaTheme = et; out.source.fontEastAsiaTheme = from; }
      if (h != null) out.fontHAnsi = h;
    }
    const sz = child(el, 'sz');
    if (sz) {
      const v = halfPtToPt(sz.attr('w:val'));
      if (v != null) { out.sizePt = v; out.source.sizePt = from; }
    }
    const szCs = child(el, 'szCs');
    if (szCs) {
      const v = halfPtToPt(szCs.attr('w:val'));
      if (v != null) out.sizeCsPt = v;
    }
    const b = child(el, 'b');
    if (b) { out.bold = onOff(b); out.source.bold = from; }
    const i = child(el, 'i');
    if (i) out.italic = onOff(i);
    const u = child(el, 'u');
    if (u) out.underline = u.attr('w:val') ?? 'single';
    const c = child(el, 'color');
    if (c) out.color = c.attr('w:val') ?? null;
  }

  // A theme reference with no explicit font must resolve through theme1.xml.
  if (themeFonts) {
    if (!out.fontAscii && out.fontAsciiTheme) {
      out.fontAscii = themeFonts.resolve(out.fontAsciiTheme, 'latin');
      out.source.fontAscii = `theme:${out.fontAsciiTheme}`;
    }
    if (!out.fontEastAsia && out.fontEastAsiaTheme) {
      out.fontEastAsia = themeFonts.resolve(out.fontEastAsiaTheme, 'ea');
      out.source.fontEastAsia = `theme:${out.fontEastAsiaTheme}`;
    }
  }
  // A part that sets only w:ascii leaves East Asian text on the theme path.
  if (!out.fontEastAsia && themeFonts) {
    out.fontEastAsia = themeFonts.defaultEastAsia();
    out.source.fontEastAsia = out.source.fontEastAsia ?? 'theme:default';
  }
  return out;
}

function mergePPr(layers) {
  const out = {
    alignment: null,
    alignFrom: null,
    spaceBeforePt: null,
    spaceAfterPt: null,
    lineSpacing: null, // { type: 'auto'|'exact'|'atLeast', value, pt }
    firstLineIndentTwip: null,
    firstLineIndentChars: null,
    leftIndentTwip: null,
    leftIndentChars: null,
    lineRuleFrom: null,
    outlineLevel: null,
    pageBreakBefore: null,
    widowControl: null,
    keepNext: null,
    numberingId: null,
    numberingLevel: null,
    styleId: null,
    source: {},
  };
  for (const { el, from } of layers) {
    if (!el) continue;
    const jc = child(el, 'jc');
    if (jc) { out.alignment = jc.attr('w:val') ?? null; out.alignFrom = from; out.source.alignment = from; }
    const spacing = child(el, 'spacing');
    if (spacing) {
      const before = spacing.attr('w:before');
      const after = spacing.attr('w:after');
      const line = spacing.attr('w:line');
      const rule = spacing.attr('w:lineRule');
      if (before != null) { out.spaceBeforePt = twipToPt(Number(before)); out.source.spaceBeforePt = from; }
      if (after != null) { out.spaceAfterPt = twipToPt(Number(after)); out.source.spaceAfterPt = from; }
      if (line != null) {
        const lv = Number(line);
        out.lineSpacing = {
          type: rule ?? 'auto',
          value: lv,
          // `pt` is a convenience for point-based rules only: for lineRule=auto
          // the stored value is a multiple of 240, so dividing it by 20 would
          // report a meaningless "18pt 行距" for 1.5 倍. Callers that compare
          // modes use `lines` there; `pt` stays null so a wrong comparison is
          // impossible rather than merely unlikely.
          pt: rule === 'auto' ? null : twipToPt(lv),
          lines: rule === 'auto' ? lv / 240 : null,
        };
        out.lineRuleFrom = from;
        out.source.lineSpacing = from;
      }
    }
    const ind = child(el, 'ind');
    if (ind) {
      const fl = ind.attr('w:firstLine');
      const flc = ind.attr('w:firstLineChars');
      const lf = ind.attr('w:left');
      const lfc = ind.attr('w:leftChars');
      const hang = ind.attr('w:hanging');
      // Word writes both `firstLine` (twips) and `firstLineChars` (hundredths
      // of a character) when "首行缩进 N 字符" is set, and the char value wins
      // at layout time. Keep both so a reader can validate either way.
      if (fl != null) { out.firstLineIndentTwip = Number(fl); out.source.firstLineIndent = from; }
      if (flc != null) { out.firstLineIndentChars = Number(flc) / 100; out.source.firstLineIndent = from; }
      if (fl == null && flc == null && hang != null) {
        out.firstLineIndentTwip = -Number(hang);
        out.source.firstLineIndent = from;
      }
      if (lf != null) out.leftIndentTwip = Number(lf);
      else if (lfc != null) out.leftIndentChars = Number(lfc) / 100;
    }
    const ol = child(el, 'outlineLvl');
    if (ol) {
      const v = Number(ol.attr('w:val'));
      if (Number.isFinite(v) && v < 9) out.outlineLevel = v;
    }
    const pbb = child(el, 'pageBreakBefore');
    if (pbb) out.pageBreakBefore = onOff(pbb);
    const wc = child(el, 'widowControl');
    if (wc) out.widowControl = onOff(wc);
    const kn = child(el, 'keepNext');
    if (kn) out.keepNext = onOff(kn);
    const numPr = child(el, 'numPr');
    if (numPr) {
      const numId = child(numPr, 'numId')?.attr('w:val');
      const ilvl = child(numPr, 'ilvl')?.attr('w:val');
      if (numId != null) out.numberingId = numId;
      if (ilvl != null) out.numberingLevel = ilvl;
    }
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* Theme fonts                                                         */
/* ------------------------------------------------------------------ */

function loadThemeFonts(themeDoc) {
  if (!themeDoc) {
    return {
      majorLatin: null, minorLatin: null, majorEa: null, minorEa: null,
      resolve: () => null,
      defaultEastAsia: () => null,
    };
  }
  const root = themeDoc.elements()[0];
  const fontScheme = findFirstByLocal(root, 'fontScheme');
  const grab = (which) => {
    const group = fontScheme ? child(fontScheme, which) : null;
    const latin = group ? child(group, 'latin') : null;
    const ea = group ? child(group, 'ea') : null;
    return {
      latin: latin?.attr('typeface') || null,
      ea: ea?.attr('typeface') || null,
    };
  };
  const major = grab('majorFont');
  const minor = grab('minorFont');
  const api = {
    majorLatin: major.latin,
    minorLatin: minor.latin,
    majorEa: major.ea,
    minorEa: minor.ea,
    resolve(name, script) {
      if (!name) return null;
      const minorKey = /minor/i.test(name);
      const group = minorKey ? minor : major;
      if (script === 'latin') return group.latin;
      return group.ea || null;
    },
    defaultEastAsia() {
      return minor.ea || null;
    },
  };
  return api;
}

/* ------------------------------------------------------------------ */
/* Document                                                            */
/* ------------------------------------------------------------------ */

export class Docx {
  constructor(zipEntries, order) {
    this.zip = zipEntries;
    this.order = order;
    this.dirty = new Set();
    this.parts = new Map();
    this.docs = new Map();
  }

  has(part) {
    return this.zip.has(part);
  }

  text(part) {
    const buf = this.zip.get(part);
    return buf ? buf.toString('utf8') : null;
  }

  doc(part) {
    if (this.docs.has(part)) return this.docs.get(part);
    const src = this.text(part);
    if (src == null) return null;
    const parsed = parseXml(src);
    this.docs.set(part, parsed);
    return parsed;
  }

  root(part) {
    const d = this.doc(part);
    return d ? d.elements()[0] ?? null : null;
  }

  markDirty(part) {
    this.dirty.add(part);
  }

  /** Relationship targets for a part, keyed by rId. */
  rels(part) {
    const dir = part.includes('/') ? part.slice(0, part.lastIndexOf('/')) : '';
    const base = part.slice(part.lastIndexOf('/') + 1);
    const relPart = `${dir}/_rels/${base}.rels`;
    if (!this.has(relPart)) return new Map();
    if (!this._relsCache) this._relsCache = new Map();
    if (this._relsCache.has(relPart)) return this._relsCache.get(relPart);
    const root = this.root(relPart);
    const map = new Map();
    for (const r of childrenNamed(root, 'Relationship')) {
      map.set(r.attr('Id'), {
        type: r.attr('Type'),
        target: r.attr('Target'),
        mode: r.attr('TargetMode'),
      });
    }
    this._relsCache.set(relPart, map);
    return map;
  }

  resolveTarget(part, target) {
    const dir = part.includes('/') ? part.slice(0, part.lastIndexOf('/')) : '';
    const segs = `${dir}/${target}`.split('/');
    const out = [];
    for (const s of segs) {
      if (s === '..') out.pop();
      else if (s !== '.' && s !== '') out.push(s);
    }
    return out.join('/');
  }

  /** All body-level block elements in document order. */
  body() {
    const root = this.root('word/document.xml');
    return root ? child(root, 'body') : null;
  }

  /** All <w:sectPr>, in document order (body last + paragraph-level ones). */
  sections() {
    const out = [];
    const body = this.body();
    if (!body) return out;
    // Paragraph-level sectPr ends the *preceding* section, and the body-level
    // one is the last section. Emitting the body-level sectPr first (as a
    // breadth-first walk does) reverses the document order, so every finding
    // that names "第 N 节" pointed at the wrong section.
    for (const c of body.children) {
      if (c.type !== 'element') continue;
      if (c.localName() === 'p') {
        const pPr = child(c, 'pPr');
        if (!pPr) continue;
        for (const cc of pPr.children) {
          if (cc.type === 'element' && cc.localName() === 'sectPr') out.push(cc);
        }
      } else if (c.localName() === 'sectPr') {
        out.push(c);
      }
    }
    return out;
  }

  sectionInfo() {
    return this.sections().map((sectPr, i) => {
      const pgSz = child(sectPr, 'pgSz');
      const pgMar = child(sectPr, 'pgMar');
      const numType = child(sectPr, 'pgNumType');
      const headerRefs = childrenNamed(sectPr, 'headerReference').map((r) => ({
        type: r.attr('w:type'),
        id: r.attr('r:id'),
        target: this.resolveTarget('word/document.xml', this.rels('word/document.xml').get(r.attr('r:id'))?.target ?? ''),
      }));
      const footerRefs = childrenNamed(sectPr, 'footerReference').map((r) => ({
        type: r.attr('w:type'),
        id: r.attr('r:id'),
        target: this.resolveTarget('word/document.xml', this.rels('word/document.xml').get(r.attr('r:id'))?.target ?? ''),
      }));
      return {
        index: i,
        el: sectPr,
        widthCm: pgSz ? round(twipToCm(Number(pgSz.attr('w:w')))) : null,
        heightCm: pgSz ? round(twipToCm(Number(pgSz.attr('w:h')))) : null,
        orientation: pgSz?.attr('w:orient') ?? 'portrait',
        marginsCm: pgMar
          ? {
              top: round(twipToCm(Number(pgMar.attr('w:top')))),
              bottom: round(twipToCm(Number(pgMar.attr('w:bottom')))),
              left: round(twipToCm(Number(pgMar.attr('w:left')))),
              right: round(twipToCm(Number(pgMar.attr('w:right')))),
            }
          : null,
        headerCm: pgMar ? round(twipToCm(Number(pgMar.attr('w:header')))) : null,
        footerCm: pgMar ? round(twipToCm(Number(pgMar.attr('w:footer')))) : null,
        gutterCm: pgMar ? round(twipToCm(Number(pgMar.attr('w:gutter') ?? 0))) : null,
        titlePg: !!child(sectPr, 'titlePg'),
        pgNumType: numType
          ? { fmt: numType.attr('w:fmt') ?? 'decimal', start: numType.attr('w:start') ?? null }
          : null,
        headerRefs,
        footerRefs,
      };
    });
  }

  /** Paragraph records for the whole body, with effective formatting resolved. */
  paragraphs() {
    const body = this.body();
    if (!body) return [];
    const styles = this.styles();
    const out = [];
    let index = 0;
    let tableDepth = 0;
    const walk = (el) => {
      for (const c of el.children) {
        if (c.type !== 'element') continue;
        const ln = c.localName();
        if (ln === 'p') {
          out.push(this.describeParagraph(c, index++, tableDepth, styles));
        } else if (ln === 'tbl') {
          tableDepth++;
          walk(c);
          tableDepth--;
        } else if (ln === 'sdt') {
          const content = child(c, 'sdtContent');
          if (content) walk(content);
        } else if (ln === 'tr' || ln === 'tc') {
          walk(c);
        }
      }
    };
    walk(body);
    return out;
  }

  describeParagraph(p, index, tableDepth, styles) {
    const pPr = child(p, 'pPr');
    const pStyle = pPr ? child(pPr, 'pStyle')?.attr('w:val') : null;
    const eff = styles.effectivePPr(pStyle, pPr);
    const styleDef = styles.get(pStyle);
    const runs = childrenNamed(p, 'r').map((r) => {
      const rPr = child(r, 'rPr');
      const rStyle = rPr ? child(rPr, 'rStyle')?.attr('w:val') : null;
      const effR = styles.effectiveRPr(rStyle ?? pStyle, rPr);
      return {
        el: r,
        rPr,
        styleId: rStyle,
        text: r.textContent,
        ...effR,
      };
    });
    // Hyperlinks / fields hold runs one level deeper; flatten their text too.
    const deeperText = findAllByLocal(p, 't')
      .map((t) => t.textContent)
      .join('');
    const text = p.textContent || deeperText;
    return {
      index,
      el: p,
      pPr,
      styleId: pStyle,
      styleName: styleDef?.name ?? (pStyle ? pStyle : 'Normal'),
      outlineLevel: eff.outlineLevel ?? styles.outlineLevel(pStyle),
      text: text.replace(/\s+/g, ' ').trim(),
      rawText: text,
      tableDepth,
      inTable: tableDepth > 0,
      hasPageBreak: findAllByLocal(p, 'br').some((b) => b.attr('w:type') === 'page'),
      hasLastRenderedPageBreak: findAllByLocal(p, 'lastRenderedPageBreak').length > 0,
      hasDrawing: !!findFirstByLocal(p, 'drawing') || !!findFirstByLocal(p, 'pict'),
      hasFieldPage: p.textContent === '' && this.paragraphHasPageField(p),
      runs,
      formatting: eff,
    };
  }

  paragraphHasPageField(p) {
    return findAllByLocal(p, 'instrText').some((t) => /PAGE/i.test(t.textContent));
  }

  /**
   * Embedded pictures with the facts the figure rules can judge.
   *
   * A picture lives in two places: the drawing markup says how big it is on the
   * page (`wp:extent`, in EMU) and the media part holds the pixels. Neither half
   * alone can answer "is this at least 150 dpi", so both are resolved here and
   * the arithmetic stays in image-size.js.
   */
  figures() {
    if (this._figures) return this._figures;
    const rels = this.rels('word/document.xml');
    const out = [];
    for (const p of this.paragraphs()) {
      if (!p.hasDrawing) continue;
      const drawings = findAllByLocal(p.el, 'drawing');
      drawings.push(...findAllByLocal(p.el, 'pict'));
      for (const drawing of drawings) {
        // `wp:extent` is the placed size; `a:ext` is the same numbers for the
        // shape. Word writes wp:extent for both inline and anchored pictures.
        const extentEl = findFirstByLocal(drawing, 'extent');
        const cx = Number(extentEl?.attr('cx') ?? NaN);
        const cy = Number(extentEl?.attr('cy') ?? NaN);
        const extent = Number.isFinite(cx) && Number.isFinite(cy) && cx > 0 && cy > 0 ? { cx, cy } : null;
        const blip = findFirstByLocal(drawing, 'blip');
        const embedId = blip?.attr('r:embed') ?? blip?.attr('r:link') ?? null;
        // Legacy VML pictures carry the id in `v:imagedata@r:id`.
        const legacy = findFirstByLocal(drawing, 'imagedata');
        const rId = embedId ?? legacy?.attr('r:id') ?? null;
        const rel = rId ? rels.get(rId) : null;
        const target = rel ? this.resolveTarget('word/document.xml', rel.target ?? '') : null;
        const bytes = target && this.has(target) ? this.zip.get(target) : null;
        const pixel = bytes ? imagePixelSize(bytes) : null;
        // `wp:inline` vs `wp:anchor` is the actual word for "embedded" vs
        // "floating". Inferring it from the presence of wp:extent would be
        // wrong: an anchored picture has an extent too.
        const anchor = findFirstByLocal(drawing, 'anchor');
        out.push({
          paragraph: p,
          part: target,
          external: rel ? rel.mode === 'External' : false,
          inline: anchor == null,
          pixel,
          extent,
          dpi: effectiveDpi(pixel, extent),
        });
      }
    }
    this._figures = out;
    return out;
  }

  styles() {
    if (this._styles) return this._styles;
    this._styles = new StyleSheet(this.doc('word/styles.xml'), this.theme());
    return this._styles;
  }

  theme() {
    if (this._theme) return this._theme;
    const part = this.has('word/theme/theme1.xml') ? 'word/theme/theme1.xml' : null;
    this._theme = loadThemeFonts(part ? this.doc(part) : null);
    return this._theme;
  }

  tables() {
    const body = this.body();
    if (!body) return [];
    const styles = this.styles();
    const out = [];
    const tbls = findAllByLocal(body, 'tbl');
    tbls.forEach((tbl, i) => {
      const tblPr = child(tbl, 'tblPr');
      const styleId = tblPr ? child(tblPr, 'tblStyle')?.attr('w:val') : null;
      // Word stores table styles under opaque numeric ids; the human name
      // ("Light Grid Accent 1") only exists on the style definition.
      const styleDef = styleId ? styles.get(styleId) : null;
      const rows = childrenNamed(tbl, 'tr');
      const rowInfo = rows.map((tr) => {
        const trPr = child(tr, 'trPr');
        const cells = childrenNamed(tr, 'tc').map((tc) => {
          const tcPr = child(tc, 'tcPr');
          const paras = childrenNamed(tc, 'p');
          return {
            el: tc,
            tcPr,
            paragraphs: paras.map((p) => this.describeParagraph(p, -1, 1, styles)),
            text: paras.map((p) => p.textContent).join(' ').replace(/\s+/g, ' ').trim(),
          };
        });
        return {
          el: tr,
          cells,
          isHeader: trPr ? !!child(trPr, 'tblHeader') : false,
          heightTwip: trPr ? Number(child(trPr, 'trHeight')?.attr('w:val') ?? 0) || null : null,
        };
      });
      out.push({
        index: i,
        el: tbl,
        styleId,
        styleName: styleDef?.name ?? styleId ?? '(无表格样式)',
        rows: rowInfo,
        rowCount: rowInfo.length,
        colCount: Math.max(0, ...rowInfo.map((r) => r.cells.length)),
        text: rowInfo.map((r) => r.cells.map((c) => c.text).join(' | ')).join('\n'),
        headerRowBold: rowInfo[0]
          ? rowInfo[0].cells.every((c) =>
              c.paragraphs.every((p) => {
                // Blank runs are padding — python-docx leaves one behind after
                // `cell.text = ''`, and a run with no text cannot be "not bold"
                // in any sense a reader would notice. Counting them reported
                // every properly bold header as unbolded.
                const runs = p.runs.filter((r) => r.text && r.text.trim());
                return runs.length > 0 && runs.every((r) => r.bold === true);
              }),
            )
          : null,
      });
    });
    return out;
  }

  /** Header/footer parts actually referenced by any section. */
  headerFooterParts() {
    const out = [];
    for (const s of this.sectionInfo()) {
      for (const h of s.headerRefs) out.push({ kind: 'header', section: s.index, ...h });
      for (const f of s.footerRefs) out.push({ kind: 'footer', section: s.index, ...f });
    }
    return out.filter((p) => p.target && this.has(p.target));
  }

  /** Effective run facts for a header/footer part's first paragraph. */
  describePartParagraphs(part) {
    const root = this.root(part);
    if (!root) return [];
    const styles = this.styles();
    return findAllByLocal(root, 'p').map((p) => this.describeParagraph(p, -1, 0, styles));
  }

  /** Serialize dirty parts and return the container buffer. */
  toBuffer() {
    for (const part of this.dirty) {
      const doc = this.docs.get(part);
      if (!doc) continue;
      const src = this.text(part) ?? '';
      const decl = /^\uFEFF?\s*<\?xml[^>]*\?>/.exec(src)?.[0] ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
      const text = `${decl}\r\n${serializeDocument(doc, { standalone: false }).replace(/^\uFEFF?\s*<\?xml[^>]*\?>\r?\n?/, '')}`;
      this.zip.set(part, Buffer.from(text, 'utf8'));
    }
    const pairs = this.order.filter((n) => this.zip.has(n)).map((n) => [n, this.zip.get(n)]);
    return writeZip(pairs);
  }

  save(path) {
    fs.writeFileSync(path, this.toBuffer());
  }
}

/* ------------------------------------------------------------------ */
/* Opening                                                             */
/* ------------------------------------------------------------------ */

export function openDocx(filePath) {
  const buf = fs.readFileSync(filePath);
  const { entries, order } = readZip(buf);
  if (!entries.has('word/document.xml')) {
    throw new Error('不是 Word 文档：缺少 word/document.xml（可能是 .doc 或加密文件）');
  }
  const docx = new Docx(entries, order);
  docx.path = filePath;
  docx.sizeBytes = buf.length;
  return docx;
}

/** Main document text with one entry per block, for structure scanning. */
export function scanBlocks(docx) {
  const body = docx.body();
  if (!body) return [];
  const blocks = [];
  for (const c of body.children) {
    if (c.type !== 'element') continue;
    if (c.localName() === 'p') blocks.push({ kind: 'p', el: c });
    else if (c.localName() === 'tbl') blocks.push({ kind: 'tbl', el: c });
    else if (c.localName() === 'sectPr') blocks.push({ kind: 'sectPr', el: c });
    else if (c.localName() === 'sdt') {
      const content = child(c, 'sdtContent');
      if (content) {
        for (const cc of content.children) {
          if (cc.type !== 'element') continue;
          if (cc.localName() === 'p') blocks.push({ kind: 'p', el: cc });
          else if (cc.localName() === 'tbl') blocks.push({ kind: 'tbl', el: cc });
        }
      }
    }
  }
  return blocks;
}

export { StyleSheet, mergeRPr, mergePPr, NS_REL };
