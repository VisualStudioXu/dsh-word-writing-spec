/**
 * The fix engine: apply the rule table to a document.
 *
 * Two hard safety rails from §8.2 and §8.6.3:
 *   1. Never write in place before a backup exists — `backupPathFor()` builds
 *      "原文件名_backup_日期.docx" next to the original, and `applyFixes` refuses
 *      to save without one.
 *   2. Heading normalization is plan-then-apply. `planHeadings` only reads; it
 *      returns the candidate list for the user to confirm, exactly as §8.6.3
 *      requires. `applyHeadings` runs only on the confirmed subset.
 *
 * Every fixer is keyed by the same rule prefix the auditor emits, so a caller
 * can say "fix rule X only" and get precisely the fixer for X.
 */

import fs from 'node:fs';
import path from 'node:path';
import { child, childrenNamed, findAllByLocal, makeElement } from './xml.js';
import { cmToTwip, ptToTwip, round } from './docx.js';
import { captionPattern, lineSpacingTarget, numberingLevel, resolveStyleId } from './spec.js';
import {
  classify,
  decideLevel,
  dominantBodySize,
  formatEvidence,
} from './outline.js';
import { buildScope } from './scope.js';

/* ------------------------------------------------------------------ */
/* Schema-ordered child insertion                                      */
/* ------------------------------------------------------------------ */

/** CT_PPr child order (ECMA-376 §17.3.1.26). */
const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
  'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
  'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
  'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
  'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
  'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle',
  'rPr', 'sectPr', 'pPrChange',
];

/** CT_RPr child order (ECMA-376 §17.3.2.28). */
const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
  'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
  'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz',
  'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign',
  'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath',
];

const localOf = (el) => (el && typeof el.localName === 'function' ? el.localName() : null);

/**
 * Find or create `localName` inside `parent`, inserting it at the position the
 * schema demands so Word does not reject the file.
 */
function ensure(parent, localName, order) {
  const existing = parent.children.find((c) => c.type === 'element' && c.localName() === localName);
  if (existing) return existing;
  const el = makeElement(localName);
  const rank = (n) => {
    const i = order.indexOf(n);
    return i < 0 ? order.length : i;
  };
  const mine = rank(localName);
  const at = parent.children.findIndex((c) => c.type === 'element' && rank(c.localName()) > mine);
  if (at < 0) parent.append(el);
  else parent.insertBefore(el, parent.children[at]);
  return el;
}

function ensurePPr(p) {
  return ensure(p, 'pPr', PPR_ORDER);
}

function ensureRPr(r) {
  return ensure(r, 'rPr', RPR_ORDER);
}

function setAttr(el, name, value) {
  el.setAttr(name, String(value));
}

/** Drop theme indirection so an explicit font actually takes effect. */
function clearThemeFonts(rFonts) {
  for (const name of ['w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme', 'w:cstheme']) {
    rFonts.removeAttr(name);
  }
}

function applyRunFont(rPr, { ascii, eastAsia, sizePt, bold, color }) {
  const rFonts = ensure(rPr, 'rFonts', RPR_ORDER);
  clearThemeFonts(rFonts);
  if (ascii) {
    setAttr(rFonts, 'w:ascii', ascii);
    setAttr(rFonts, 'w:hAnsi', ascii);
  }
  if (eastAsia) setAttr(rFonts, 'w:eastAsia', eastAsia);
  if (sizePt != null) {
    const half = Math.round(sizePt * 2);
    setAttr(ensure(rPr, 'sz', RPR_ORDER), 'w:val', half);
    setAttr(ensure(rPr, 'szCs', RPR_ORDER), 'w:val', half);
  }
  if (bold != null) {
    const b = ensure(rPr, 'b', RPR_ORDER);
    setAttr(b, 'w:val', bold ? '1' : '0');
    const bCs = ensure(rPr, 'bCs', RPR_ORDER);
    setAttr(bCs, 'w:val', bold ? '1' : '0');
  }
  // Written only when the rule table names a colour. An absent key means the
  // document keeps whatever colour it already had, rather than being silently
  // repainted black.
  if (typeof color === 'string' && color.trim() !== '') {
    setAttr(ensure(rPr, 'color', RPR_ORDER), 'w:val', color.trim().toUpperCase());
  }
}

function applyLineSpacing(pPr, spacing0) {
  const spacing = ensure(pPr, 'spacing', PPR_ORDER);
  // One place decides twips-vs-multiple: `lineSpacingTarget` in spec.js. The
  // writer used to branch on the mode name itself, so a mode it did not know
  // fell through to the multiplier branch and wrote a plausible-looking but
  // wrong `w:line`.
  const target = lineSpacingTarget(spacing0);
  if (target.unit === 'pt') {
    setAttr(spacing, 'w:line', ptToTwip(target.value));
    setAttr(spacing, 'w:lineRule', target.rule);
  } else {
    setAttr(spacing, 'w:line', Math.round(target.value * 240));
    setAttr(spacing, 'w:lineRule', 'auto');
  }
  return spacing;
}

function runParagraphs(docx) {
  return docx.paragraphs();
}

/** The same body-paragraph scope the auditor uses, so fixes match findings. */
function bodyParagraphs(docx, spec) {
  const scope = buildScope(docx, spec, { resolveStyleId });
  return scope.paragraphs.filter(scope.isBody);
}

function headingParagraphs(docx, level) {
  const styles = docx.styles();
  const id = resolveStyleId(styles, `heading${level}`);
  if (!id) return [];
  return docx.paragraphs().filter((p) => p.styleId === id && p.text);
}

/* ------------------------------------------------------------------ */
/* Fixers                                                              */
/* ------------------------------------------------------------------ */

function fixPageSetup(docx, spec) {
  let count = 0;
  for (const sect of docx.sections()) {
    const pgSz = ensure(sect, 'pgSz', ['footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']);
    setAttr(pgSz, 'w:w', cmToTwip(spec.page.widthCm));
    setAttr(pgSz, 'w:h', cmToTwip(spec.page.heightCm));
    pgSz.removeAttr('w:orient');
    const pgMar = ensure(sect, 'pgMar', ['footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']);
    setAttr(pgMar, 'w:top', cmToTwip(spec.page.marginsCm.top));
    setAttr(pgMar, 'w:bottom', cmToTwip(spec.page.marginsCm.bottom));
    setAttr(pgMar, 'w:left', cmToTwip(spec.page.marginsCm.left));
    setAttr(pgMar, 'w:right', cmToTwip(spec.page.marginsCm.right));
    setAttr(pgMar, 'w:header', cmToTwip(spec.page.headerDistanceCm));
    setAttr(pgMar, 'w:footer', cmToTwip(spec.page.footerDistanceCm));
    count++;
  }
  return count;
}

function fixBodyFormat(docx, spec) {
  const body = spec.body;
  // The rule table stores the indent as EMU (266700 = 0.74cm); OOXML wants
  // twips. Derive rather than read a second, drift-prone key.
  const firstLineTwip = Math.round(body.firstLineIndentEmu / 635);
  const firstLineChars = Math.round(body.firstLineIndentChars * 100);
  let count = 0;
  for (const p of bodyParagraphs(docx, spec)) {
    const pPr = ensurePPr(p.el);
    const spacing = applyLineSpacing(pPr, body.lineSpacing);
    setAttr(spacing, 'w:before', ptToTwip(body.spaceBeforePt ?? 0));
    setAttr(spacing, 'w:after', ptToTwip(body.spaceAfterPt ?? 0));
    const ind = ensure(pPr, 'ind', PPR_ORDER);
    ind.removeAttr('w:hanging');
    ind.removeAttr('w:hangingChars');
    setAttr(ind, 'w:firstLine', firstLineTwip);
    setAttr(ind, 'w:firstLineChars', firstLineChars);
    setAttr(ensure(pPr, 'jc', PPR_ORDER), 'w:val', body.alignment);
    for (const r of childrenNamed(p.el, 'r')) {
      if (!r.textContent.trim()) continue;
      applyRunFont(ensureRPr(r), {
        ascii: body.fontEN,
        eastAsia: body.fontCN,
        sizePt: body.sizePt,
        bold: null, // body emphasis is intentional; never flatten it
        color: body.color,
      });
    }
    count++;
  }
  return count;
}

function applyStyleDefinition(docx, def) {
  const styles = docx.styles();
  const id = resolveStyleId(styles, `heading${def.level}`);
  const style = styles.get(id);
  if (!style) return false;
  const sEl = style.el;
  // <w:style> children: name, aliases, basedOn, next, link, autoRedefine,
  // hidden, uiPriority, semiHidden, unhideWhenUsed, qFormat, locked, personal,
  // personalCompose, personalReply, rsid, pPr, rPr, tblPr, trPr, tcPr.
  const sOrder = ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority', 'semiHidden', 'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose', 'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr'];
  const rPr = ensure(sEl, 'rPr', sOrder);
  applyRunFont(rPr, {
    ascii: def.fontEN,
    eastAsia: def.fontCN,
    sizePt: def.sizePt,
    bold: def.bold,
    color: def.color,
  });
  const pPr = ensure(sEl, 'pPr', sOrder);
  setAttr(ensure(pPr, 'jc', PPR_ORDER), 'w:val', def.alignment);
  setAttr(ensure(pPr, 'outlineLvl', PPR_ORDER), 'w:val', def.level - 1);
  const spacing = applyLineSpacing(pPr, { type: 'exact', valuePt: def.level === 1 ? 26 : def.level === 2 ? 24 : 22 });
  setAttr(spacing, 'w:before', ptToTwip(def.spaceBeforePt ?? 0));
  setAttr(spacing, 'w:after', ptToTwip(def.spaceAfterPt ?? 0));
  const ind = ensure(pPr, 'ind', PPR_ORDER);
  ind.removeAttr('w:firstLine');
  ind.removeAttr('w:firstLineChars');
  setAttr(ind, 'w:firstLine', 0);
  setAttr(ensure(pPr, 'keepNext', PPR_ORDER), 'w:val', '1');
  if (def.pageBreakBefore) setAttr(ensure(pPr, 'pageBreakBefore', PPR_ORDER), 'w:val', '1');
  return true;
}

/**
 * Label shapes the writer can normalise, in match order. The first capture is
 * the label; whatever spaces follow it are replaced with the required gap.
 * Ordered strictest-first so "1.1" is not read as the label "1.".
 */
const HEADING_LABELS = [
  /^(\s*附\s*录\s*[A-Z])/,
  /^(\s*第[一二三四五六七八九十百零〇\d]+[章节])/,
  /^(\s*[一二三四五六七八九十百零〇]+、)/,
  /^(\s*[（(][一二三四五六七八九十百零〇\d]+[）)])/,
  /^(\s*[①②③④⑤⑥⑦⑧⑨⑩])/,
  /^(\s*\d+(?:\.\d+)+)/,
  /^(\s*\d+\.(?!\d))/,
  /^(\s*\d+)/,
];

/**
 * Normalise the spaces between a heading's label and its title.
 *
 * `def` is optional. When given, a scheme level whose `gap` is null (「一、标题」
 * carries no gap by convention) collapses the spaces to zero instead of the
 * table-wide `numberingGapSpaces`; without it the old numeric-only behaviour is
 * preserved for callers that only have the spec.
 */
function normalizeHeadingGap(p, spec, def) {
  const schemeId = def?.numberingScheme ?? spec?.headings?.numberingScheme;
  const lvl = def !== undefined && schemeId ? numberingLevel(schemeId, def.level) : null;
  const gap = lvl !== null && lvl.gap === null ? 0 : spec.headings.numberingGapSpaces;
  const runs = childrenNamed(p.el, 'r').filter((r) => r.textContent.trim());
  if (!runs.length) return false;
  const t = findAllByLocal(runs[0], 't')[0];
  if (!t) return false;
  const original = t.textContent;
  let text = original;
  for (const re of HEADING_LABELS) {
    const m = re.exec(original);
    if (!m) continue;
    const label = m[1];
    text = label + ' '.repeat(gap) + original.slice(label.length).replace(/^[ \t\u3000]+/, '');
    break;
  }
  if (text === original) return false;
  t.text = text;
  if (t.children.length) t.children = [];
  return true;
}

function fixHeadings(docx, spec) {
  let count = 0;
  for (const def of spec.headings.levels) {
    if (applyStyleDefinition(docx, def)) count++;
    for (const p of headingParagraphs(docx, def.level)) {
      const pPr = ensurePPr(p.el);
      setAttr(ensure(pPr, 'jc', PPR_ORDER), 'w:val', def.alignment);
      setAttr(ensure(pPr, 'outlineLvl', PPR_ORDER), 'w:val', def.level - 1);
      const ind = ensure(pPr, 'ind', PPR_ORDER);
      ind.removeAttr('w:firstLine');
      ind.removeAttr('w:firstLineChars');
      ind.removeAttr('w:hanging');
      setAttr(ind, 'w:firstLine', 0);
      const spacing = applyLineSpacing(pPr, { type: 'exact', valuePt: def.level === 1 ? 26 : def.level === 2 ? 24 : 22 });
      setAttr(spacing, 'w:before', ptToTwip(def.spaceBeforePt ?? 0));
      setAttr(spacing, 'w:after', ptToTwip(def.spaceAfterPt ?? 0));
      if (def.pageBreakBefore) setAttr(ensure(pPr, 'pageBreakBefore', PPR_ORDER), 'w:val', '1');
      else pPr.children.find((c) => c.type === 'element' && c.localName() === 'pageBreakBefore')?.removeAttr('w:val');
      for (const r of childrenNamed(p.el, 'r')) {
        if (!r.textContent.trim()) continue;
        applyRunFont(ensureRPr(r), {
          ascii: def.fontEN,
          eastAsia: def.fontCN,
          sizePt: def.sizePt,
          bold: def.bold,
        });
      }
      normalizeHeadingGap(p, spec, def);
      count++;
    }
  }
  return count;
}

function fixChapterPageBreaks(docx, spec) {
  const styles = docx.styles();
  const id = resolveStyleId(styles, 'heading1');
  if (!id) return 0;
  let count = 0;
  const chapters = docx.paragraphs().filter((p) => p.styleId === id && p.text);
  chapters.forEach((p, i) => {
    if (i === 0) return;
    if (p.hasPageBreak || p.formatting.pageBreakBefore === true) return;
    const pPr = ensurePPr(p.el);
    setAttr(ensure(pPr, 'pageBreakBefore', PPR_ORDER), 'w:val', '1');
    count++;
  });
  return count;
}

function fixTables(docx, spec) {
  const s = spec.table;
  let count = 0;
  for (const t of docx.tables()) {
    t.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell) => {
        cell.paragraphs.forEach((p) => {
          for (const r of childrenNamed(p.el, 'r')) {
            if (!r.textContent.trim()) continue;
            applyRunFont(ensureRPr(r), {
              ascii: s.bodyFontEN,
              eastAsia: rowIndex === 0 ? s.headerRowFontCN ?? s.bodyFontCN : s.bodyFontCN,
              sizePt: s.bodySizePt,
              bold: rowIndex === 0 ? !!s.headerRowBold : null,
            });
          }
        });
      });
    });
    count++;
  }
  return count;
}

function fixFigures(docx, spec) {
  const s = spec.figure;
  const paras = docx.paragraphs();
  const below = s.captionPosition !== 'above';
  const capPattern = captionPattern(s.numberPattern, /^图\s*\d+[-.]\d+/);
  let count = 0;
  for (const p of paras) {
    if (!p.hasDrawing) continue;
    const pPr = ensurePPr(p.el);
    // A fixed line height crops the image; "at least 1.5 lines" does not.
    const spacing = ensure(pPr, 'spacing', PPR_ORDER);
    setAttr(spacing, 'w:line', 360);
    setAttr(spacing, 'w:lineRule', 'atLeast');
    setAttr(spacing, 'w:before', ptToTwip(s.paragraphSpaceBeforePt));
    setAttr(spacing, 'w:after', ptToTwip(s.paragraphSpaceAfterPt));
    setAttr(ensure(pPr, 'jc', PPR_ORDER), 'w:val', s.paragraphAlignment);

    // The caption is reported as fixable, so something has to fix it. Only the
    // formatting is repaired; a missing caption is content, and inventing "图
    // x-x" text would number it wrong on the next insertion.
    const step = below ? 1 : -1;
    let caption = null;
    for (let k = p.index + step; k >= 0 && k < paras.length; k += step) {
      if (paras[k].hasDrawing) break;
      if (paras[k].text.trim()) {
        caption = paras[k];
        break;
      }
    }
    if (caption && capPattern.test(caption.text)) {
      const cPr = ensurePPr(caption.el);
      setAttr(ensure(cPr, 'jc', PPR_ORDER), 'w:val', s.captionAlignment);
      for (const r of childrenNamed(caption.el, 'r')) {
        if (!r.textContent.trim()) continue;
        applyRunFont(ensureRPr(r), {
          ascii: s.captionFontEN,
          eastAsia: s.captionFontCN,
          sizePt: s.captionSizePt,
          bold: s.captionBold === true ? true : null,
          color: s.color,
        });
      }
    }
    count++;
  }
  return count;
}

function fixCodeBlocks(docx, spec) {
  const s = spec.code;
  let count = 0;
  for (const p of docx.paragraphs()) {
    if (!p.text.trim() || p.inTable) continue;
    const runs = childrenNamed(p.el, 'r').filter((r) => r.textContent.trim());
    if (!runs.length) continue;
    const courier = runs.every((r) => /courier/i.test(
      child(child(r, 'rPr') ?? makeElement('rPr'), 'rFonts')?.attr('w:ascii') ?? '',
    ));
    if (!courier) continue;
    const pPr = ensurePPr(p.el);
    const ind = ensure(pPr, 'ind', PPR_ORDER);
    setAttr(ind, 'w:left', cmToTwip(s.indentLeftCm));
    setAttr(ind, 'w:firstLine', 0);
    ind.removeAttr('w:firstLineChars');
    ind.removeAttr('w:hanging');
    const spacing = applyLineSpacing(pPr, s.lineSpacing);
    setAttr(spacing, 'w:before', 0);
    setAttr(spacing, 'w:after', 0);
    setAttr(ensure(pPr, 'jc', PPR_ORDER), 'w:val', 'left');
    for (const r of runs) {
      applyRunFont(ensureRPr(r), { ascii: s.font, sizePt: s.sizePtRange[0] + 0.5, bold: false });
    }
    count++;
  }
  return count;
}

function fixHeaderFooter(docx, spec) {
  const s = spec.headerFooter;
  let count = 0;
  for (const part of docx.headerFooterParts()) {
    const root = docx.root(part.target);
    if (!root) continue;
    const isHeader = part.kind === 'header';
    for (const p of findAllByLocal(root, 'p')) {
      const pPr = ensurePPr(p);
      setAttr(ensure(pPr, 'jc', PPR_ORDER), 'w:val', isHeader ? s.headerAlignment : s.pageNumberAlignment);
      for (const r of childrenNamed(p, 'r')) {
        const rPr = ensureRPr(r);
        applyRunFont(rPr, {
          ascii: s.headerFontEN,
          eastAsia: s.headerFontCN,
          sizePt: s.headerSizePt,
          bold: false,
        });
        if (isHeader && s.headerColor) setAttr(ensure(rPr, 'color', RPR_ORDER), 'w:val', s.headerColor);
      }
    }
    docx.markDirty(part.target);
    count++;
  }

  // The page number's numeral style is a property of the SECTION, not of the
  // footer text: writing "I" into the footer would freeze it, while
  // pgNumType lets Word recompute it on every page.
  const numeral = s.pageNumberNumeral ?? 'decimal';
  const sections = docx.sections();
  const bodySection = sections[sections.length - 1];
  if (bodySection) {
    const pgNumType = ensure(bodySection, 'pgNumType', ['footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']);
    setAttr(pgNumType, 'w:fmt', numeral);
  }

  return count;
}

function fixCoverAndToc(docx, spec) {
  const cover = spec.cover;
  const toc = spec.toc;
  if (!cover && !toc) return 0;
  const styles = docx.styles();
  const paras = docx.paragraphs();
  let count = 0;

  // --- cover title ---------------------------------------------------
  if (cover) {
    const h1 = resolveStyleId(styles, 'heading1');
    const firstHeadingIdx = paras.findIndex((p) => p.styleId === h1);
    const coverEnd = firstHeadingIdx > 0 ? firstHeadingIdx : Math.min(paras.length, 12);
    const coverParas = paras.slice(0, coverEnd);

    // Same selection rule as the auditor: an explicit `cover.title` names the
    // paragraph to restyle; otherwise the largest type on the cover is the
    // title. Fixing a different paragraph than the one that was reported would
    // leave the finding standing after a "successful" fix.
    const explicitTitle = String(cover.title ?? '').trim();
    let titlePara = null;
    if (explicitTitle !== '') {
      titlePara = coverParas.find((p) => p.text.trim() === explicitTitle)
        ?? coverParas.find((p) => p.text.includes(explicitTitle))
        ?? null;
    } else {
      let best = null;
      for (const p of coverParas) {
        const size = Math.max(0, ...p.runs.map((r) => r.sizePt ?? 0));
        if (size > 0 && (best === null || size > best.size)) best = { p, size };
      }
      titlePara = best?.p ?? null;
    }
    if (titlePara) {
      const pPr = ensurePPr(titlePara.el);
      if (cover.alignment) setAttr(ensure(pPr, 'jc', PPR_ORDER), 'w:val', cover.alignment);
      if (cover.spaceBeforePt != null) {
        setAttr(ensure(pPr, 'spacing', PPR_ORDER), 'w:before', ptToTwip(cover.spaceBeforePt));
      }
      for (const r of childrenNamed(titlePara.el, 'r')) {
        if (!r.textContent.trim()) continue;
        applyRunFont(ensureRPr(r), {
          ascii: cover.titleFontEN ?? spec.body.fontEN,
          eastAsia: cover.titleFontCN ?? spec.body.fontCN,
          sizePt: cover.titleSizePt,
          bold: cover.titleBold,
          color: cover.color,
        });
      }
      count++;
    }
  }

  // --- TOC heading ---------------------------------------------------
  //
  // The field itself is not synthesised: a TOC field's cached entries are pages
  // of result runs, and writing a placeholder would make Word show stale page
  // numbers until someone pressed F9. Restyling the heading is safe; inserting
  // the field is the writer's job, which the checklist says out loud.
  const tocTitle = String(toc?.title ?? '').trim();
  if (toc?.enabled && tocTitle !== '') {
    const p = paras.find((q) => q.text.trim() === tocTitle);
    if (p) {
      const pPr = ensurePPr(p.el);
      setAttr(ensure(pPr, 'jc', PPR_ORDER), 'w:val', toc?.alignment ?? 'left');
      for (const r of childrenNamed(p.el, 'r')) {
        if (!r.textContent.trim()) continue;
        applyRunFont(ensureRPr(r), {
          ascii: spec.body.fontEN,
          eastAsia: toc.titleFontCN ?? spec.body.fontCN,
          sizePt: toc.titleSizePt,
          bold: toc.titleBold,
          color: toc.titleColor,
        });
      }
      count++;
    }
  }
  return count;
}

function fixAppendixGap(docx, spec) {
  const styles = docx.styles();
  const id = resolveStyleId(styles, 'heading1');
  if (!id) return 0;
  let count = 0;
  for (const p of docx.paragraphs()) {
    if (p.styleId !== id) continue;
    if (!/^\s*附\s*录/.test(p.text)) continue;
    if (normalizeHeadingGap(p, spec)) count++;
  }
  return count;
}

/* ------------------------------------------------------------------ */
/* Heading normalization (§8.6) — plan, then apply                     */
/* ------------------------------------------------------------------ */

/**
 * Phase 1 of §8.6.3: scan and identify. Never modifies the document.
 * @returns {{candidates: object[], bodySizePt: number}}
 */
export function planHeadings(docx, spec) {
  const paras = docx.paragraphs();
  const styles = docx.styles();
  const bodySize = dominantBodySize(paras) ?? spec.body.sizePt;
  const headingIds = [1, 2, 3, 4].map((l) => resolveStyleId(styles, `heading${l}`));
  const candidates = [];
  for (const p of paras) {
    if (p.styleId && headingIds.includes(p.styleId)) continue;
    const c = classify(p, { bodySizePt: bodySize, maxHeadingChars: 40 });
    if (!c) continue;
    const ev = formatEvidence(p, bodySize);
    const decided = decideLevel(c, ev);
    candidates.push({
      paragraph: p.index + 1,
      text: p.text.slice(0, 60),
      currentStyle: p.styleId ?? 'Normal',
      reason: c.reason,
      suggestedLevel: decided.level,
      confidence: c.confidence,
      evidence: {
        sizePt: ev.sizePt,
        sizeDelta: round(ev.sizeDelta, 1),
        bold: ev.bold,
        centered: ev.centered,
        outdented: ev.outdented,
      },
      requiresConfirmation: c.confidence !== 'high' || decided.level == null,
    });
  }
  return { candidates, bodySizePt: bodySize };
}

/**
 * Phase 2 of §8.6.3: convert the confirmed subset. Callers must pass only the
 * paragraphs the user accepted. A backup must already exist.
 * @param {object[]} accepted rows from planHeadings().candidates
 */
export function applyHeadings(docx, spec, accepted) {
  const styles = docx.styles();
  const paras = docx.paragraphs();
  const applied = [];
  for (const row of accepted) {
    const p = paras.find((x) => x.index === row.paragraph - 1);
    if (!p) continue;
    const level = row.suggestedLevel ?? row.level;
    if (!level) continue;
    const id = resolveStyleId(styles, `heading${level}`);
    if (!id) continue;
    const pPr = ensurePPr(p.el);
    setAttr(ensure(pPr, 'pStyle', PPR_ORDER), 'w:val', id);
    setAttr(ensure(pPr, 'outlineLvl', PPR_ORDER), 'w:val', level - 1);
    const ind = ensure(pPr, 'ind', PPR_ORDER);
    ind.removeAttr('w:firstLine');
    ind.removeAttr('w:firstLineChars');
    setAttr(ind, 'w:firstLine', 0);
    const def = spec.headings.levels.find((l) => l.level === level);
    if (def) {
      const spacing = applyLineSpacing(pPr, { type: 'exact', valuePt: level === 1 ? 26 : level === 2 ? 24 : 22 });
      setAttr(spacing, 'w:before', ptToTwip(def.spaceBeforePt ?? 0));
      setAttr(spacing, 'w:after', ptToTwip(def.spaceAfterPt ?? 0));
      for (const r of childrenNamed(p.el, 'r')) {
        if (!r.textContent.trim()) continue;
        applyRunFont(ensureRPr(r), {
          ascii: def.fontEN,
          eastAsia: def.fontCN,
          sizePt: def.sizePt,
          bold: def.bold,
        });
      }
      if (def.pageBreakBefore) setAttr(ensure(pPr, 'pageBreakBefore', PPR_ORDER), 'w:val', '1');
    }
    normalizeHeadingGap(p, spec);
    applied.push({ paragraph: row.paragraph, text: p.text.slice(0, 40), level });
  }
  return applied;
}

/* ------------------------------------------------------------------ */
/* Driver                                                              */
/* ------------------------------------------------------------------ */

const FIXERS = [
  { rule: 'page', label: '页面设置（纸张、页边距、页眉页脚距离）', apply: fixPageSetup },
  { rule: 'body', label: '正文格式（字体、字号、行距、首行缩进、对齐）', apply: fixBodyFormat },
  { rule: 'headings', label: '标题格式与标题样式定义', apply: fixHeadings },
  { rule: 'sectionBreak', label: '一级标题另起新页', apply: fixChapterPageBreaks },
  { rule: 'appendix', label: '附录标题编号空格', apply: fixAppendixGap },
  { rule: 'table', label: '表格字体与字号', apply: fixTables },
  { rule: 'figure', label: '插图段落行距与对齐', apply: fixFigures },
  { rule: 'code', label: '代码块字体、缩进、行距', apply: fixCodeBlocks },
  { rule: 'headerFooter', label: '页眉页脚字体与对齐', apply: fixHeaderFooter },
  { rule: 'cover', label: '封面与目录标题格式', apply: fixCoverAndToc },
];

/**
 * Apply fixes for the given rule prefixes.
 * @param {import('./docx.js').Docx} docx
 * @param {object} spec
 * @param {{rules?: string[], skip?: string[]}} [options]
 */
export function applyFixes(docx, spec, options = {}) {
  const wanted = options.rules;
  const skip = options.skip ?? [];
  const applied = [];
  const skipped = [];
  for (const fixer of FIXERS) {
    if (wanted && !wanted.some((w) => fixer.rule === w || fixer.rule.startsWith(w) || w.startsWith(fixer.rule))) continue;
    if (skip.some((s) => fixer.rule === s || fixer.rule.startsWith(s))) {
      skipped.push({ rule: fixer.rule, label: fixer.label, reason: '按调用方要求跳过' });
      continue;
    }
    let count = 0;
    try {
      count = fixer.apply(docx, spec);
    } catch (err) {
      skipped.push({ rule: fixer.rule, label: fixer.label, reason: `执行失败：${err.message}` });
      continue;
    }
    if (count > 0) applied.push({ rule: fixer.rule, label: fixer.label, count });
    // A fixer that changed something must mark its part, or saveWithBackup
    // writes the untouched original back and the "applied" list lies. Keying
    // this off a hard-coded rule list meant every new fixer silently became a
    // no-op on disk (cover and toc did exactly that).
    if (count > 0) {
      docx.markDirty('word/document.xml');
      docx.markDirty('word/styles.xml');
    }
  }
  return { applied, skipped };
}

/** "原文件名_backup_20261017.docx" next to the original (§8.2). */
export function backupPathFor(filePath, date = new Date()) {
  const dir = path.dirname(filePath);
  const ext = path.extname(filePath);
  const base = path.basename(filePath, ext);
  const stamp = `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
  let candidate = path.join(dir, `${base}_backup_${stamp}${ext}`);
  let n = 1;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${base}_backup_${stamp}_${n++}${ext}`);
  }
  return candidate;
}

/**
 * Save a modified document, creating the §8.2 backup first.
 * @returns {{backup: string, output: string, bytes: number}}
 */
export function saveWithBackup(docx, outputPath, options = {}) {
  const source = docx.path;
  if (!source) throw new Error('文档没有来源路径，无法建立备份');
  if (!fs.existsSync(source)) throw new Error(`原文件不存在：${source}`);
  const backup = options.backupPath ?? backupPathFor(source, options.date ?? new Date());
  if (!fs.existsSync(backup)) fs.copyFileSync(source, backup);
  const buf = docx.toBuffer();
  fs.writeFileSync(outputPath, buf);
  return {
    backup,
    output: outputPath,
    bytes: buf.length,
    verified: fs.statSync(backup).size === fs.statSync(source).size,
  };
}

export { FIXERS, ensure, ensurePPr, ensureRPr, applyRunFont, normalizeHeadingGap };
