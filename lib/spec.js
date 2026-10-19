/**
 * Rule-table loader.
 *
 * The JSON in rules/ is the SINGLE source of truth: the system-prompt
 * constraint text, the audit checks and the fix writer all read it. Nothing
 * about a font size or margin is hard-coded elsewhere, so changing the rules
 * file changes generation guidance and review verdicts together.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
export const RULES_PATH = path.join(HERE, '..', 'rules', 'docx-writing-spec.json');
export const SCHEMA_PATH = path.join(HERE, '..', 'rules', 'docx-writing-spec.schema.json');

let cached = null;
let cachedMtime = 0;

/** Load (and hot-reload on mtime change) the rule table. */
export function loadSpec() {
  const stat = fs.statSync(RULES_PATH);
  if (cached && stat.mtimeMs === cachedMtime) return cached;
  cached = JSON.parse(fs.readFileSync(RULES_PATH, 'utf8'));
  cachedMtime = stat.mtimeMs;
  return cached;
}

/**
 * Drop the cache. The settings page writes the rule table itself, and two
 * writes inside the same filesystem timestamp tick would otherwise leave the
 * mtime check looking unchanged and keep serving the old table until restart.
 */
export function invalidateSpec() {
  cached = null;
  cachedMtime = 0;
}

/**
 * Line-spacing modes, named the way Word and WPS show them.
 *
 * `unit: 'pt'` modes store twips in `w:line` and name themselves in
 * `w:lineRule`; `unit: 'lines'` modes store the multiple in `w:line` as
 * `multiple * 240` and always use `w:lineRule="auto"`. Word cannot tell
 * "1.5 倍" from "多倍行距 1.5" once it is written, so the distinction lives
 * here and only the factor travels into the file.
 *
 * `factor` is null when the mode reads its number from the rule table, which is
 * why `exact`/`atLeast`/`multiple` still need `valuePt` while `single`,
 * `oneAndHalf` and `double` do not.
 */
export const LINE_SPACING_MODES = {
  single: { label: '单倍行距', unit: 'lines', factor: 1 },
  oneAndHalf: { label: '1.5 倍行距', unit: 'lines', factor: 1.5 },
  double: { label: '2 倍行距', unit: 'lines', factor: 2 },
  multiple: { label: '多倍行距', unit: 'lines', factor: null },
  atLeast: { label: '最小值', unit: 'pt', factor: null },
  exact: { label: '固定值', unit: 'pt', factor: null },
};

/** Fall back to `exact` so an unknown mode degrades to a readable verdict. */
export function lineSpacingMode(type) {
  return LINE_SPACING_MODES[type] ?? LINE_SPACING_MODES.exact;
}

/**
 * What a compliant paragraph must carry, in the unit the mode speaks.
 *
 * Returning a unit instead of a number keeps every caller from re-deriving
 * "is this a multiplier or a point size?", which is the question that used to
 * be answered three different ways in the auditor, the writer and the prompt.
 */
export function lineSpacingTarget(spacing) {
  const mode = lineSpacingMode(spacing?.type);
  if (mode.unit === 'lines') {
    return { unit: 'lines', value: mode.factor ?? Number(spacing?.valuePt ?? 1) };
  }
  return { unit: 'pt', value: Number(spacing?.valuePt ?? 0), rule: mode === LINE_SPACING_MODES.atLeast ? 'atLeast' : 'exact' };
}

/** Human-readable form, used by the prompt, the findings and the checklist. */
export function lineSpacingText(spacing) {
  const mode = lineSpacingMode(spacing?.type);
  const target = lineSpacingTarget(spacing);
  if (target.unit === 'lines') {
    return mode.factor === null ? `多倍行距 ${target.value} 倍` : mode.label;
  }
  return `${mode.label} ${target.value} 磅`;
}

/** 12 → "12", 10.5 → "10.5", 3 → "3" (no trailing ".0"). */
function trimNum(value) {
  return typeof value === 'number' ? String(Number(value.toFixed(3))) : String(value);
}

/**
 * The page-number style is one editable string, not three.
 *
 * Splitting it into `pageNumberPrefix` / `pageNumberSuffix` plus a numeral enum
 * forced the owner to edit two boxes to change what readers see as one thing,
 * and made half-finished states possible ("— " with an empty suffix). The
 * format is now the whole visual: `数字` marks where the counter goes, exactly
 * as the rule table always wrote it ("— 数字 —").
 *
 * Both the prompt text and the auditor read this, so there is one answer to
 * "what should the footer look like".
 */
export const PAGE_NUMBER_TOKEN = '数字';

/**
 * @param {string | undefined} format e.g. "— 数字 —" or "第 数字 页"
 * @returns {{ prefix: string, suffix: string, format: string, hasToken: boolean }}
 */
export function pageNumberParts(format) {
  const text = typeof format === 'string' && format !== '' ? format : `— ${PAGE_NUMBER_TOKEN} —`;
  const at = text.indexOf(PAGE_NUMBER_TOKEN);
  if (at < 0) {
    // A format with no token would drop the number entirely; keep the text as a
    // prefix and append the counter so the footer still shows a page number.
    return { prefix: text.trim() === '' ? '' : `${text} `, suffix: '', format: text, hasToken: false };
  }
  return {
    prefix: text.slice(0, at),
    suffix: text.slice(at + PAGE_NUMBER_TOKEN.length),
    format: text,
    hasToken: true,
  };
}

/**
 * Resolve a concrete numeral style into the visible decoration, used by the
 * findings so a reader sees "— 阿拉伯数字 1, 2, 3 —" rather than "— 数字 —".
 */
export function pageNumberText(spec, numeralText) {
  const parts = pageNumberParts(spec?.headerFooter?.pageNumberFormat);
  return `${parts.prefix}${numeralText}${parts.suffix}`;
}

const ALIGN_TEXT = { left: '左对齐', center: '居中', right: '右对齐', both: '两端对齐' };

const POSITION_TEXT = { above: '位于表格上方', below: '位于图片下方' };

/**
 * Resolve `{path}` / `{path:formatter}` tokens in a checklist label.
 *
 * The self-check labels used to hard-code the numbers they describe
 * ("行距（固定值 22 磅）"), so changing the rule left the label advertising the
 * old value — exactly the drift this plugin exists to prevent, reproduced
 * inside the plugin itself. Labels are now templates: the wording stays
 * hand-written, every number comes from the table it describes.
 *
 * Unknown paths resolve to the original token so a typo shows up as `{body.fnt}`
 * in the rendered list instead of silently blanking the sentence.
 */
export function resolveLabel(template, spec) {
  if (typeof template !== 'string') return template;
  return template.replace(/\{([A-Za-z0-9_.]+)(?::([A-Za-z]+))?\}/g, (whole, path, fmt) => {
    const value = path.split('.').reduce((cur, seg) => {
      if (cur == null) return undefined;
      if (Array.isArray(cur)) return cur[Number(seg)];
      return cur[seg];
    }, spec);
    if (value === undefined) return whole;
    switch (fmt) {
      case 'spacing':
        return lineSpacingText(value);
      case 'bold':
        return value ? '加粗' : '不加粗';
      case 'pageBreak':
        return value ? '，另起页' : '';
      case 'range':
        return `${trimNum(value[0])}~${trimNum(value[1])}pt`;
      case 'align':
        return ALIGN_TEXT[value] ?? value;
      case 'position':
        return POSITION_TEXT[value] ?? value;
      case 'color':
        return String(value).toLowerCase() === '808080' ? '灰色' : `#${value}`;
      case 'pageNumber':
        return pageNumberParts(value).format;
      default:
        return trimNum(value);
    }
  });
}

/** A checklist item's label with every value read from the current table. */
export function checklistLabel(item, spec) {
  return resolveLabel(item?.label ?? '', spec);
}

/**
 * Compile a caption's `numberPattern` from the rule table.
 *
 * The pattern is an anchored prefix like `^图\s*\d+-\d+`, so it can be tested
 * against the whole caption text. A table that omits it falls back to the
 * literal default rather than matching nothing — a rule table with a missing key
 * must not silently turn every caption into "missing".
 */
export function captionPattern(numberPattern, fallback) {
  if (typeof numberPattern === 'string' && numberPattern.trim() !== '') {
    try {
      return new RegExp(numberPattern);
    } catch {
      // An invalid pattern is a rule-table bug; the fallback keeps the audit
      // running so the rest of the findings still surface.
      return fallback;
    }
  }
  return fallback;
}

/**
 * Heading numbering schemes.
 *
 * A scheme is a table of per-level labels, not a format string: the auditor has
 * to recognise "第一章" as a label and measure the gap after it, and the writer
 * has to produce the same token. Sharing the table keeps the two from drifting,
 * which is why `numbering`/`numberingExample` in the rule table are *resolved*
 * values written by the settings page rather than the source of truth.
 *
 * `gap` says whether the label takes `numberingGapSpaces` spaces before the
 * title. "一、标题" and "（一）标题" carry no gap by convention, so they are null.
 */
export const NUMBERING_SCHEMES = {
  arabicDotted: {
    label: '1 / 1.1 / 1.1.1',
    note: '科技文档、报告',
    levels: [
      { example: '1', pattern: '^\\d+$', gap: 'spaces' },
      { example: '1.1', pattern: '^\\d+\\.\\d+$', gap: 'spaces' },
      { example: '1.1.1', pattern: '^\\d+\\.\\d+\\.\\d+$', gap: 'spaces' },
    ],
  },
  chineseChapter: {
    label: '第一章 / 第一节 / 一、',
    note: '论文、书稿',
    levels: [
      { example: '第一章', pattern: '^第[一二三四五六七八九十百零〇\\d]+章$', gap: 'spaces' },
      { example: '第一节', pattern: '^第[一二三四五六七八九十百零〇\\d]+节$', gap: 'spaces' },
      { example: '一、', pattern: '^[一二三四五六七八九十百零〇]+、', gap: null },
    ],
  },
  chineseOrdinal: {
    label: '一、 / （一） / 1.',
    note: '公文、规范',
    levels: [
      { example: '一、', pattern: '^[一二三四五六七八九十百零〇]+、', gap: null },
      { example: '（一）', pattern: '^（[一二三四五六七八九十百零〇]+）', gap: null },
      { example: '1.', pattern: '^\\d+\\.(?!\\d)', gap: 'spaces' },
    ],
  },
  arabicParen: {
    label: '1 / （1） / ①',
    note: '与 deeperLevelRepresentation 一致',
    levels: [
      { example: '1', pattern: '^\\d+$', gap: 'spaces' },
      { example: '（1）', pattern: '^（\\d+）', gap: null },
      { example: '①', pattern: '^[①②③④⑤⑥⑦⑧⑨⑩]', gap: null },
    ],
  },
};

/** Fall back to arabicDotted so an unknown id still yields working labels. */
export function numberingScheme(id) {
  return NUMBERING_SCHEMES[id] ?? NUMBERING_SCHEMES.arabicDotted;
}

/** The label spec for one level of a scheme; null past the last defined level. */
export function numberingLevel(schemeId, level) {
  return numberingScheme(schemeId).levels[level - 1] ?? null;
}

/** Font aliases: a document may legitimately carry the fallback family. */
export function fontMatches(actual, expected, fallback) {
  if (!actual) return false;
  const norm = (s) => String(s).trim().toLowerCase().replace(/\s+/g, ' ');
  const a = norm(actual);
  if (a === norm(expected)) return true;
  if (fallback && a === norm(fallback)) return true;
  return false;
}

/**
 * Word stores built-in style names in the document's language ("heading 1"
 * rather than "Heading 1") and style IDs vary by producer ("Heading1" vs "3").
 * So every style lookup goes through this canonical form.
 */
export function canonicalStyleName(name) {
  return String(name ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, '');
}

export const STYLE_ALIASES = {
  heading1: ['heading1', '标题1', '标题一', 'h1'],
  heading2: ['heading2', '标题2', '标题二', 'h2'],
  heading3: ['heading3', '标题3', '标题三', 'h3'],
  heading4: ['heading4', '标题4'],
  normal: ['normal', '正文', '默认段落字体'],
  caption: ['caption', '题注'],
  header: ['header', '页眉'],
  footer: ['footer', '页脚'],
  toc: ['tocheading', 'toc1', 'toc2', 'toc3', '目录标题'],
};

/**
 * Find the actual style id for a canonical role, by name first then by
 * well-known id, tolerating numeric ids (a real Word export used "3" for
 * heading 1).
 */
export function resolveStyleId(styles, role) {
  const wanted = STYLE_ALIASES[role] ?? [canonicalStyleName(role)];
  for (const [, s] of styles.byId) {
    if (s.type !== 'paragraph' && s.type !== 'character') continue;
    if (wanted.includes(canonicalStyleName(s.name))) return s.id;
  }
  for (const [id, s] of styles.byId) {
    if (wanted.includes(canonicalStyleName(id))) return id;
    if (s.name && wanted.includes(canonicalStyleName(s.name))) return id;
  }
  return null;
}

/** Is this paragraph a real heading, by style role or outline level? */
export function headingLevelOf(structure, styles, para) {
  const spec = structure;
  const roleIds = ['heading1', 'heading2', 'heading3', 'heading4'].map((r) => resolveStyleId(styles, r));
  if (para.styleId && roleIds.includes(para.styleId)) {
    return roleIds.indexOf(para.styleId) + 1;
  }
  if (para.outlineLevel != null && para.outlineLevel >= 0 && para.outlineLevel <= 8) {
    return para.outlineLevel + 1;
  }
  void spec;
  return null;
}
