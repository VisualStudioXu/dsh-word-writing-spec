/**
 * Document scope: which paragraphs are "body text" in the sense §3 and §8.6
 * mean.
 *
 * This exists because getting scope wrong produces confident nonsense. In the
 * reference document the cover title is 26pt, the TOC entries are `toc 1`/`toc
 * 2` styles holding `HYPERLINK \l _Toc...` field instructions, and the abstract
 * and 目录 are Heading 1 — auditing all of that as "body text" reports dozens of
 * font and size errors that are not errors at all. Both the auditor and the
 * fixer call into this module, so a finding and its fix always agree on scope.
 *
 * Scope is derived from structure, never from a hard-coded paragraph index:
 *   - the body starts at the first styled Heading 1 (cover and abstract sit
 *     above it), or at the first paragraph matching the first chapter's
 *     numbering when nothing is styled;
 *   - TOC entries are field-driven: a paragraph carrying an `instrText` that
 *     starts with `TOC` or `HYPERLINK \l` is a directory entry;
 *   - captions, drawings, table cells and code lines are their own scopes.
 */

import { findAllByLocal } from './xml.js';

const TOC_STYLE = /^(toc|目录|目录\d)/i;

/** A TOC field code or a TOC hyperlink: this paragraph is a directory entry. */
export function isTocEntry(para) {
  if (!para?.el) return false;
  if (para.styleId && TOC_STYLE.test(String(para.styleId))) return true;
  if (para.styleName && TOC_STYLE.test(String(para.styleName))) return true;
  const fields = findAllByLocal(para.el, 'instrText').map((t) => t.textContent.trim());
  return fields.some((f) => /^(TOC\b|HYPERLINK\s+\\l\b)/i.test(f));
}

/**
 * Locate the body of the document: cover + front matter excluded.
 * @returns {{start:number, reason:string}}
 */
export function bodyStart(paragraphs, styles, resolveStyleId) {
  const h1 = resolveStyleId(styles, 'heading1');
  if (h1) {
    const firstChapter = paragraphs.find(
      (p) => p.styleId === h1 && !/^\s*附\s*录/.test(p.text) && /^\s*\d+/.test(p.text),
    );
    if (firstChapter) return { start: firstChapter.index, reason: `首个编号一级标题「${firstChapter.text.slice(0, 20)}」` };
    const firstHeading = paragraphs.find((p) => p.styleId === h1);
    if (firstHeading) return { start: firstHeading.index, reason: `首个一级标题「${firstHeading.text.slice(0, 20)}」` };
  }
  // No styled headings at all: trust the numbering, first "1" / "第1章" / "一、".
  const firstNumbered = paragraphs.find((p) => !p.inTable && /^\s*(1[\s.]|第[一1][章节]|[一]、)/.test(p.text));
  return firstNumbered
    ? { start: firstNumbered.index, reason: `首个章级编号段落「${firstNumbered.text.slice(0, 20)}」` }
    : { start: 0, reason: '未识别出封面/前置部分，全文按正文处理' };
}

/**
 * Build the scope once per document.
 * @returns {{bodyStart:number, bodyStartReason:string, isBody:(p)=>boolean}}
 */
export function buildScope(docx, spec, { resolveStyleId }) {
  const paragraphs = docx.paragraphs();
  const styles = docx.styles();
  const { start, reason } = bodyStart(paragraphs, styles, resolveStyleId);
  const headingIds = [1, 2, 3, 4].map((l) => resolveStyleId(styles, `heading${l}`));
  const isHeading = (p) => (p.styleId && headingIds.includes(p.styleId)) || (p.outlineLevel != null && p.outlineLevel < 9);

  // A code line is its own scope, exactly like a caption: Courier New at 8.5pt
  // with a 1.2cm left indent is correct for `code` and wrong for `body`, so
  // letting body checks see it reported every code block as a font/size/indent
  // triple error and buried the real findings.
  const isCodeLine = (p) => {
    const runs = p.runs.filter((r) => r.text && r.text.trim());
    if (!runs.length) return false;
    return runs.every((r) => /courier/i.test(r.fontAscii ?? '') || /courier/i.test(r.fontHAnsi ?? ''));
  };

  const isBody = (p) => {
    if (p.inTable || !p.text) return false;
    if (p.index < start) return false; // cover / abstract / TOC
    if (isHeading(p)) return false;
    if (isTocEntry(p)) return false;
    if (p.hasDrawing) return false;
    if (/^(表|图|续表)\s*\d+[-.]\d+/.test(p.text)) return false; // captions
    if (isCodeLine(p)) return false;
    return true;
  };

  return { bodyStart: start, bodyStartReason: reason, paragraphs, isBody, isHeading };
}
