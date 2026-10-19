/**
 * Heading-structure detection for the audit and for the normalizer.
 *
 * Two distinct jobs, deliberately kept apart:
 *  1. `classify` decides whether a paragraph LOOKS like a heading it should be
 *     (per §8.6.1's three-layer rule: format + text + position). This is the
 *     "could be converted" side, used only to *propose* candidates.
 *  2. `audit` decides whether the document CONFORMS: a styled heading must
 *     carry the level's font/size/bold, and chapter-level headings must start
 *     a new page.
 *
 * The count matters: the same numeric pattern `^1.2` is a heading in body text
 * but a subsection reference inside a table cell, and `表 2-3 ...` is a table
 * caption, not a heading. Both are excluded by position, not by regex, because
 * §8.6.5 lists them as known misjudgements.
 */

const CN_NUM = '一二三四五六七八九十百零〇';

/** Text-based numbering patterns, in strictest-first order. */
export function textHeadingCandidate(text) {
  const t = String(text ?? '').trim();
  if (!t) return null;
  if (t.length > 40) return null;

  // 第X章 / 第X节  -> level 1
  if (new RegExp(`^第[${CN_NUM}\\d]+[章节]`).test(t)) return { level: 1, kind: 'chapter' };
  // 一、 -> level 1 ; （一） -> level 2 ; 1. -> level 1
  if (new RegExp(`^[${CN_NUM}]+、`).test(t)) return { level: 1, kind: 'cn-dun' };
  if (new RegExp(`^（[${CN_NUM}]+）`).test(t)) return { level: 2, kind: 'cn-paren' };
  if (new RegExp(`^\\([${CN_NUM}]+\\)`).test(t)) return { level: 2, kind: 'cn-paren-ascii' };
  // X.X.X -> level 3, X.X -> level 2, X. -> level 1 (checked deepest first)
  if (/^\d+\.\d+\.\d+(\s|$|[^\d.])/.test(t) || /^\d+\.\d+\.\d+\.\d+/.test(t)) return { level: 3, kind: 'dotted3' };
  if (/^\d+\.\d+(\s|$|[^\d.])/.test(t)) return { level: 2, kind: 'dotted2' };
  if (/^\d+\.(\s|$|[^\d.])/.test(t)) return { level: 1, kind: 'dotted1' };
  // Explicitly NOT headings per §8.6.1 (list items)
  if (/^\d+）/.test(t) || /^（\d+）/.test(t) || /^\(\d+\)/.test(t)) return null;
  return null;
}

/** Words that mark a table/figure caption (§8.6.5: captions are not headings). */
export function looksLikeCaption(text) {
  return /^(表|图|续表)\s*\d+[-.]\d+/.test(String(text ?? '').trim());
}

/** Abstract / TOC / keywords blocks are metadata, not chapter content. */
export function looksLikeFrontMatter(text) {
  return /^(摘\s*要|abstract|目\s*录|关键词|key\s*words?|前\s*言)\s*$/i.test(String(text ?? '').trim());
}

export function looksLikeReferenceHeading(text) {
  return /^(参考文献|致\s*谢|附\s*录)/.test(String(text ?? '').trim());
}

export function looksLikeCodeLine(text) {
  const t = String(text ?? '');
  return /[{};]\s*$/.test(t) || /^\s{2,}\S/.test(t) || /\bVAR\b|\bEND_VAR\b|\bIF\b.*\bTHEN\b/.test(t);
}

/**
 * Three-layer format evidence (§8.6.1 layer 1).
 * `bodySizePt` is the document's own dominant body size, not the spec's, so a
 * document written at 10.5pt is still judged against its own baseline.
 */
export function formatEvidence(para, bodySizePt) {
  const first = para.runs.find((r) => r.text && r.text.trim()) ?? para.runs[0] ?? null;
  const size = first?.sizePt ?? null;
  const bold = first?.bold === true;
  const centered = para.formatting.alignment === 'center';
  const outdented = para.formatting.firstLineIndentTwip == null && para.formatting.firstLineIndentChars == null;
  const sizeDelta = size != null && bodySizePt != null ? size - bodySizePt : null;
  return {
    sizePt: size,
    bold,
    centered,
    outdented,
    sizeDelta,
    // §8.6.1: "size larger than body by >= 2pt" AND at least two of
    // {bigger font, wholly bold, centered}.
    bigger: sizeDelta != null && sizeDelta >= 2,
    strong:
      (sizeDelta != null && sizeDelta >= 2) === true &&
      [bold, centered, outdented].filter(Boolean).length >= 2,
  };
}

/**
 * Classify a paragraph as a heading candidate.
 * @returns {{level:number, reason:string, confidence:'high'|'low'}|null}
 */
export function classify(para, ctx) {
  if (para.inTable) return null;
  const text = para.text;
  if (!text) return null;
  if (looksLikeCaption(text)) return null;
  if (looksLikeFrontMatter(text)) return null;
  if (para.hasDrawing) return null;
  if (looksLikeCodeLine(text) && !textHeadingCandidate(text)) return null;

  const byText = textHeadingCandidate(text);
  const ev = formatEvidence(para, ctx.bodySizePt);

  if (byText) {
    // Text numbering is authoritative (§8.6.2), but a metadata line or a
    // citation-like line still gets a lower confidence so the UI can hold it.
    const confidence = ev.bold || ev.bigger || ev.centered ? 'high' : 'low';
    return { level: byText.level, reason: `文字特征：${byText.kind}`, confidence };
  }

  // No numbering: only a strongly formatted short standalone paragraph set
  // inside an otherwise body-formatted document qualifies, and it is always
  // low confidence (§8.6.4 "不确定标待确认，宁可少转").
  if (ev.strong && text.length <= ctx.maxHeadingChars) {
    return { level: null, reason: `格式特征：字号 +${ev.sizeDelta}pt、${ev.bold ? '加粗' : ''}${ev.centered ? '居中' : ''}`, confidence: 'low' };
  }
  return null;
}

/**
 * Level decision per §8.6.2: text numbering wins; otherwise font-size delta
 * (>4pt level1, 2-4pt level2, 1-2pt level3).
 */
export function decideLevel(candidate, ev) {
  if (candidate.level != null) return { level: candidate.level, by: 'textNumbering' };
  const d = ev.sizeDelta;
  if (d == null) return { level: null, by: 'none' };
  if (d > 4) return { level: 1, by: 'fontSizeDelta' };
  if (d >= 2) return { level: 2, by: 'fontSizeDelta' };
  if (d >= 1) return { level: 3, by: 'fontSizeDelta' };
  return { level: null, by: 'none' };
}

/** Dominant body font size, used as the "is it bigger" baseline. */
export function dominantBodySize(paragraphs) {
  const counts = new Map();
  for (const p of paragraphs) {
    if (p.inTable || !p.text) continue;
    if (p.outlineLevel != null && p.outlineLevel < 9) continue;
    const r = p.runs.find((x) => x.text && x.text.trim());
    if (!r?.sizePt) continue;
    counts.set(r.sizePt, (counts.get(r.sizePt) ?? 0) + 1);
  }
  let best = null;
  let bestCount = -1;
  for (const [size, count] of counts) {
    if (count > bestCount) {
      best = size;
      bestCount = count;
    }
  }
  return best;
}
