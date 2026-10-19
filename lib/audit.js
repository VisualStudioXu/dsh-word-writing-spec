/**
 * The audit engine: turn a document plus the rule table into findings.
 *
 * Every check is a pure function of (model facts, spec). Nothing is reported
 * that was not measured, and every finding carries the rule id, the clause it
 * implements, where the offending thing lives (paragraph index / table index /
 * part name) and the expected-vs-actual pair — so a report can be acted on
 * without re-reading the document.
 *
 * Severity vocabulary (stable; used by the fix engine and the UI):
 *   error   — violates a hard rule from the spec, must be fixed before delivery
 *   warning — deviates from a rule in a way that may be intentional
 *   info    — observation worth knowing, no verdict
 */

import {
  round,
  twipToCm,
  TWIP_PER_CM,
} from './docx.js';
import {
  resolveStyleId,
  canonicalStyleName,
  captionPattern,
  checklistLabel,
  fontMatches,
  lineSpacingTarget,
  lineSpacingText,
  numberingLevel,
  pageNumberParts,
} from './spec.js';
import {
  classify,
  decideLevel,
  dominantBodySize,
  formatEvidence,
  looksLikeFrontMatter,
  textHeadingCandidate,
} from './outline.js';
import { buildScope, isTocEntry } from './scope.js';

const CM = (v) => (v == null ? null : round(v, 2));

const withinTolerance = (actual, expected, tolerance) =>
  actual != null && expected != null && Math.abs(actual - expected) <= tolerance;

/**
 * How many spaces belong between the numbering label and the title text.
 * Returns null when the heading carries no label at all ("摘  要") or when the
 * level's scheme puts no gap there ("一、标题"), because §3 only governs
 * 编号与文字之间 and 一、 already ends with a delimiter.
 */
function gapFor(para, spec, def) {
  const raw = String(para.rawText ?? '').replace(/^\s+/, '');
  // 附录 is outside every numbering scheme and always takes the gap.
  if (/^附\s*录\s*[A-Z](\s|$)/.test(raw)) return spec.headings.numberingGapSpaces;
  const schemeId = def?.numberingScheme;
  if (schemeId) {
    const lvl = numberingLevel(schemeId, def.level);
    if (lvl !== null) return lvl.gap === 'spaces' ? spec.headings.numberingGapSpaces : null;
  }
  if (/^\d+(\.\d+)*(\s|$)/.test(raw)) return spec.headings.numberingGapSpaces;
  if (/^第[一二三四五六七八九十百零〇\d]+[章节](\s|$)/.test(raw)) return spec.headings.numberingGapSpaces;
  if (/^[一二三四五六七八九十百零〇]+、/.test(raw)) return null; // "一、标题" has no gap convention
  return null;
}

/**
 * The regex a level's label must satisfy: the rule table's stored pattern wins
 * (the settings page resolves it from the scheme), otherwise derive it from the
 * scheme. Null means "no opinion" — an empty or unparseable pattern must not
 * turn every heading into a finding.
 */
function labelRegexFor(def) {
  if (typeof def.numbering === 'string' && def.numbering.trim() !== '') {
    try {
      return new RegExp(def.numbering);
    } catch {
      // fall through to the scheme
    }
  }
  const lvl = def.numberingScheme ? numberingLevel(def.numberingScheme, def.level) : null;
  if (lvl === null) return null;
  try {
    return new RegExp(lvl.pattern);
  } catch {
    return null;
  }
}

/**
 * Spaces between the heading label and its title, measured on raw text.
 * The label itself may contain spaces ("附录 A"), so splitting on the first
 * whitespace would report the space inside the label as the gap.
 */
function headingGap(raw) {
  const s = String(raw ?? '').replace(/^\s+/, '');
  const patterns = [
    /^(附\s*录\s*[A-Z])([ \t\u3000]*)/,
    /^(第[一二三四五六七八九十百零〇\d]+[章节])([ \t\u3000]*)/,
    /^(\d+(?:\.\d+)*)([ \t\u3000]*)/,
    /^([一二三四五六七八九十百零〇]+、)([ \t\u3000]*)/,
  ];
  for (const re of patterns) {
    const m = re.exec(s);
    if (m) return { label: m[1], gap: m[2].length };
  }
  return null;
}

/* ------------------------------------------------------------------ */
/* Checks                                                              */
/* ------------------------------------------------------------------ */

function checkPageSetup(docx, spec) {
  const findings = [];
  const sections = docx.sectionInfo();
  const tol = spec.page.toleranceCm;
  // The paper name is whatever the rule table says; hard-coding "A4" here made the
  // finding lie as soon as the setting page offered another size.
  const paperName =
    typeof spec.page.size === 'string' && spec.page.size.trim() !== '' ? spec.page.size.trim() : 'A4';
  if (!sections.length) {
    findings.push({
      rule: 'page',
      severity: 'error',
      title: '文档没有节设置（sectPr）',
      expected: `${paperName} 页面 + 页边距`,
      actual: '缺失',
      location: 'word/document.xml',
      fixable: false,
    });
    return findings;
  }
  sections.forEach((s) => {
    const where = `第 ${s.index + 1} 节`;
    if (!withinTolerance(s.widthCm, spec.page.widthCm, tol) || !withinTolerance(s.heightCm, spec.page.heightCm, tol)) {
      findings.push({
        rule: 'page.size',
        severity: 'error',
        title: `${where}纸张尺寸不是 ${paperName}`,
        expected: `${paperName} ${spec.page.widthCm}×${spec.page.heightCm} cm`,
        actual: `${s.widthCm}×${s.heightCm} cm`,
        location: `sectPr[${s.index}]`,
        fixable: true,
      });
    }
    const m = s.marginsCm ?? {};
    const expectedMargins = spec.page.marginsCm;
    for (const key of ['top', 'bottom', 'left', 'right']) {
      const allowedLeft = key === 'left' ? [expectedMargins.left, spec.page.leftMarginNoBindingCm] : [expectedMargins[key]];
      const ok = allowedLeft.some((v) => withinTolerance(m[key], v, tol));
      if (!ok) {
        findings.push({
          rule: `page.margins.${key}`,
          severity: 'error',
          title: `${where}${key === 'top' ? '上' : key === 'bottom' ? '下' : key === 'left' ? '左' : '右'}页边距不符`,
          expected: allowedLeft.map((v) => `${v}cm`).join(' 或 '),
          actual: m[key] == null ? '未设置' : `${m[key]}cm`,
          location: `sectPr[${s.index}]/pgMar`,
          fixable: true,
        });
      }
    }
    if (!withinTolerance(s.headerCm, spec.page.headerDistanceCm, tol)) {
      findings.push({
        rule: 'page.headerDistance',
        severity: 'error',
        title: `${where}页眉距上边界不符`,
        expected: `${spec.page.headerDistanceCm}cm`,
        actual: s.headerCm == null ? '未设置' : `${s.headerCm}cm`,
        location: `sectPr[${s.index}]/pgMar`,
        fixable: true,
      });
    }
    if (!withinTolerance(s.footerCm, spec.page.footerDistanceCm, tol)) {
      findings.push({
        rule: 'page.footerDistance',
        severity: 'error',
        title: `${where}页脚距下边界不符`,
        expected: `${spec.page.footerDistanceCm}cm`,
        actual: s.footerCm == null ? '未设置' : `${s.footerCm}cm`,
        location: `sectPr[${s.index}]/pgMar`,
        fixable: true,
      });
    }
  });
  return findings;
}

function describeRunFont(run) {
  return {
    cn: run.fontEastAsia ?? null,
    en: run.fontAscii ?? null,
    size: run.sizePt ?? null,
    bold: run.bold ?? null,
  };
}

function checkBodyFormat(docx, spec) {
  const findings = [];
  const body = spec.body;
  const styles = docx.styles();
  const scope = buildScope(docx, spec, { resolveStyleId });

  const bodyParas = scope.paragraphs.filter(scope.isBody);

  const fontIssues = [];
  const sizeIssues = [];
  const spacingIssues = [];
  const indentIssues = [];
  const alignIssues = [];
  const colorIssues = [];
  const wantColor = typeof body.color === 'string' && body.color.trim() !== '' ? body.color.trim().toLowerCase() : null;

  for (const p of bodyParas) {
    const runs = p.runs.filter((r) => r.text && r.text.trim());
    if (!runs.length) continue;
    for (const r of runs) {
      const hasCN = /[\u4e00-\u9fff]/.test(r.text);
      const hasLatin = /[A-Za-z0-9]/.test(r.text);
      if (hasCN && !fontMatches(r.fontEastAsia, body.fontCN, body.fontFallbackCN)) {
        fontIssues.push({ p, run: r, kind: 'cn', actual: r.fontEastAsia, from: r.source.fontEastAsia });
      }
      if (hasLatin && !fontMatches(r.fontAscii, body.fontEN, body.fontFallbackEN)) {
        fontIssues.push({ p, run: r, kind: 'en', actual: r.fontAscii, from: r.source.fontAscii });
      }
      if (r.sizePt != null && Math.abs(r.sizePt - body.sizePt) > 0.5) {
        sizeIssues.push({ p, run: r, actual: r.sizePt });
      }
      if (wantColor !== null) {
        // A run with no `w:color` inherits — Word renders it black, which IS
        // 000000. Telling the user to set the default colour explicitly on
        // every run would add noise without changing a single pixel, so the
        // implicit case only fails when the rule asks for something other
        // than black.
        const actual = r.color == null ? null : String(r.color).toLowerCase();
        const ok = actual === null ? wantColor === '000000' : actual === wantColor;
        if (!ok) colorIssues.push({ p, run: r, actual: actual ?? '继承（默认黑）' });
      }
    }

    const f = p.formatting;
    const want = lineSpacingTarget(body.lineSpacing);
    const ls = f.lineSpacing;
    if (ls) {
      // Compare in the mode's own unit. Checking "not exact means 1 line" made
      // 1.5 倍 and 多倍行距 indistinguishable from 单倍, so a rule table asking
      // for 1.5 倍 reported every compliant paragraph as wrong.
      const ok = want.unit === 'pt'
        ? ls.type === want.rule && Math.abs((ls.pt ?? 0) - want.value) <= 0.5
        : ls.type === 'auto' && Math.abs((ls.lines ?? 0) - want.value) <= 0.05;
      if (!ok) {
        const actual = ls.type === 'auto'
          ? `${round(ls.lines ?? 0, 2)} 倍行距`
          : `${ls.type === 'atLeast' ? '最小值' : '固定值'} ${round(ls.pt, 1)} 磅`;
        spacingIssues.push({ p, actual });
      }
    } else {
      spacingIssues.push({ p, actual: '未设置行距（随样式）' });
    }

    // Measure from the twip value only: `firstLineChars` and `firstLine` are
    // mutually exclusive in OOXML, and reading whichever one is set twice
    // produces mismatched units.
    const indentTwip = f.firstLineIndentTwip;
    const indentChars = indentTwip != null && indentTwip > 0
      ? round(indentTwip / (body.sizePt * 20), 2)
      : null;
    const okIndent =
      indentTwip != null &&
      Math.abs(indentTwip - body.firstLineIndentEmu / 635) <= 30;
    if (!okIndent) {
      indentIssues.push({
        p,
        actual: indentTwip == null
          ? '无首行缩进'
          : indentTwip === 0
            ? '首行缩进 0'
            : `${round(twipToCm(indentTwip), 2)}cm（约 ${indentChars ?? '?'} 字符）`,
      });
    }

    if (f.alignment && f.alignment !== body.alignment) {
      alignIssues.push({ p, actual: f.alignment });
    }
  }

  const sample = (arr, n = 6) => arr.slice(0, n);
  const paraRef = (p) => `段落 #${p.index + 1}「${p.text.slice(0, 24)}」`;

  if (fontIssues.length) {
    const cnCount = fontIssues.filter((i) => i.kind === 'cn').length;
    const enCount = fontIssues.length - cnCount;
    const first = fontIssues[0];
    findings.push({
      rule: 'body.font',
      severity: 'error',
      title: '正文字体不符合规范',
      expected: `中文 ${body.fontCN}，英文/数字 ${body.fontEN}`,
      actual: `中文不符 ${cnCount} 处、英文不符 ${enCount} 处；首例 ${paraRef(first.p)} 中文=${first.kind === 'cn' ? first.actual ?? '(继承主题)' : 'ok'}`,
      detail: {
        sample: sample(fontIssues).map((i) => ({
          paragraph: i.p.index + 1,
          text: i.p.text.slice(0, 30),
          script: i.kind,
          font: i.actual ?? '(未显式设置，继承自 ' + i.from + ')',
          inheritedFrom: i.from,
        })),
      },
      location: 'word/document.xml 正文 run rPr',
      fixable: true,
    });
  }

  if (sizeIssues.length) {
    findings.push({
      rule: 'body.size',
      severity: 'error',
      title: '正文字号不符合规范',
      expected: `${body.sizeNameCN} ${body.sizePt}pt`,
      actual: `${sizeIssues.length} 处不符；首例 ${paraRef(sizeIssues[0].p)} 为 ${sizeIssues[0].actual}pt`,
      detail: { sample: sample(sizeIssues).map((i) => ({ paragraph: i.p.index + 1, sizePt: i.actual })) },
      location: 'word/document.xml 正文 run rPr/sz',
      fixable: true,
    });
  }

  if (spacingIssues.length) {
    findings.push({
      rule: 'body.lineSpacing',
      severity: 'error',
      title: '正文行距不符合规范',
      expected: lineSpacingText(body.lineSpacing),
      actual: `${spacingIssues.length} 处不符；首例 ${paraRef(spacingIssues[0].p)} 为 ${spacingIssues[0].actual}`,
      detail: { sample: sample(spacingIssues).map((i) => ({ paragraph: i.p.index + 1, actual: i.actual })) },
      location: 'word/document.xml 段落 pPr/spacing',
      fixable: true,
    });
  }

  if (indentIssues.length) {
    findings.push({
      rule: 'body.firstLineIndent',
      severity: 'error',
      title: '正文首行缩进不符合规范',
      expected: `首行缩进 ${body.firstLineIndentChars} 字符`,
      actual: `${indentIssues.length} 处不符；首例 ${paraRef(indentIssues[0].p)} 为 ${indentIssues[0].actual}`,
      detail: { sample: sample(indentIssues).map((i) => ({ paragraph: i.p.index + 1, actual: i.actual })) },
      location: 'word/document.xml 段落 pPr/ind',
      fixable: true,
    });
  }

  if (alignIssues.length) {
    findings.push({
      rule: 'body.alignment',
      severity: 'warning',
      title: '正文对齐方式不是两端对齐',
      expected: 'both（两端对齐）',
      actual: `${alignIssues.length} 处；首例 ${paraRef(alignIssues[0].p)} 为 ${alignIssues[0].actual}`,
      detail: { sample: sample(alignIssues).map((i) => ({ paragraph: i.p.index + 1, actual: i.actual })) },
      location: 'word/document.xml 段落 pPr/jc',
      fixable: true,
    });
  }

  if (colorIssues.length) {
    findings.push({
      rule: 'body.color',
      severity: 'warning',
      title: '正文字体颜色不符合规范',
      expected: `#${wantColor}`,
      actual: `${colorIssues.length} 处；首例 ${paraRef(colorIssues[0].p)} 为 ${colorIssues[0].actual}`,
      detail: { sample: sample(colorIssues).map((i) => ({ paragraph: i.p.index + 1, actual: i.actual })) },
      location: 'word/document.xml 正文 run rPr/color',
      fixable: true,
    });
  }

  // Widow/orphan control is a document-level + style-level setting.
  const normalId = resolveStyleId(styles, 'normal');
  const normal = styles.get(normalId);
  const wc = normal?.pPr ? normal.pPr.children.find((c) => c.name === 'w:widowControl') : null;
  if (body.widowControl && wc && /^(0|false|off)$/.test(wc.attr('w:val') ?? '')) {
    findings.push({
      rule: 'body.widowControl',
      severity: 'warning',
      title: '正文样式关闭了孤行控制',
      expected: '开启（避免孤行）',
      actual: 'widowControl=0',
      location: `styles.xml style[${normalId}]`,
      fixable: true,
    });
  }

  return findings;
}

function checkHeadings(docx, spec) {
  const findings = [];
  const paras = docx.paragraphs();
  const styles = docx.styles();
  const bodySize = dominantBodySize(paras) ?? spec.body.sizePt;

  const levelIds = [1, 2, 3, 4].map((l) => resolveStyleId(styles, `heading${l}`));
  const styled = paras.filter((p) => p.styleId && levelIds.includes(p.styleId));
  const styleLevel = (id) => levelIds.indexOf(id) + 1;

  // 1. Per-level formatting conformance on styled headings.
  for (const rawDef of spec.headings.levels) {
    // A scheme id on the level wins, so a document may mix (e.g. 第一章 for
    // level 1 and 一、 for level 3) without duplicating it in every level.
    const def = { ...rawDef, numberingScheme: rawDef.numberingScheme ?? spec.headings.numberingScheme };
    const id = resolveStyleId(styles, `heading${def.level}`);
    if (!id) {
      findings.push({
        rule: `headings.${def.level}.style`,
        severity: 'error',
        title: `文档缺少 ${def.level} 级标题样式（Heading ${def.level}）`,
        expected: def.styleId,
        actual: '未定义',
        location: 'styles.xml',
        fixable: false,
      });
      continue;
    }
    const bad = [];
    for (const p of styled.filter((x) => x.styleId === id)) {
      const run = p.runs.find((r) => r.text && r.text.trim()) ?? p.runs[0];
      if (!run) continue;
      const problems = [];
      if (run.sizePt == null || Math.abs(run.sizePt - def.sizePt) > 0.5) problems.push(`字号 ${run.sizePt ?? '继承'}pt（应 ${def.sizePt}pt）`);
      if (def.bold !== (run.bold === true)) problems.push(`加粗 ${run.bold === true ? '是' : '否'}（应 ${def.bold ? '是' : '否'}）`);
      if (!fontMatches(run.fontEastAsia, def.fontCN, null)) problems.push(`中文字体 ${run.fontEastAsia ?? '继承主题'}（应 ${def.fontCN}）`);
      if (typeof def.color === 'string' && def.color.trim() !== '') {
        const want = def.color.trim().toLowerCase();
        const actual = run.color == null ? null : String(run.color).toLowerCase();
        const ok = actual === null ? want === '000000' : actual === want;
        if (!ok) problems.push(`颜色 ${actual ?? '继承（默认黑）'}（应 ${want}）`);
      }
      if (def.alignment === 'left' && p.formatting.alignment && p.formatting.alignment !== 'left') {
        problems.push(`对齐 ${p.formatting.alignment}（应左对齐）`);
      }
      if (p.formatting.firstLineIndentTwip != null && p.formatting.firstLineIndentTwip > 0) problems.push('存在首行缩进（应顶格）');
      // Measure the gap on the raw text: `text` collapses runs of spaces, so a
      // compliant two-space gap reads back as one space.
      const raw = p.rawText.replace(/^\s+/, '');
      const measured = headingGap(raw);
      const gap = measured?.gap ?? 0;
      const expectedGap = gapFor(p, spec, def);
      if (expectedGap != null && measured && gap !== expectedGap) problems.push(`编号后空格 ${gap} 个（应 ${expectedGap} 个）`);
      // The label itself must match the level's numbering scheme. Without this
      // check a document numbered "1.1" at level 1 passes every font test and
      // only looks wrong to a human reader. 附录 A is outside every scheme —
      // it is a fixed H1 label — so it is exempt.
      const labelPattern = labelRegexFor(def);
      const isAppendixLabel = /^附\s*录/.test(measured?.label ?? '');
      if (measured && !isAppendixLabel && labelPattern !== null && !labelPattern.test(measured.label)) {
        problems.push(`编号「${measured.label}」不符合本级形态（${def.numberingExample ?? def.numbering}）`);
      }
      if (problems.length) bad.push({ p, problems });
    }
    if (bad.length) {
      findings.push({
        rule: `headings.${def.level}`,
        severity: 'error',
        title: `${def.level} 级标题格式不符合规范`,
        expected: `${def.fontCN} ${def.sizePt}pt${def.bold ? ' 加粗' : ' 不加粗'}，左对齐，编号后 ${spec.headings.numberingGapSpaces} 个空格`,
        actual: `${bad.length} 个标题有问题；首例「${bad[0].p.text.slice(0, 20)}」：${bad[0].problems.join('；')}`,
        detail: { sample: bad.slice(0, 8).map((b) => ({ paragraph: b.p.index + 1, text: b.p.text.slice(0, 30), problems: b.problems })) },
        location: `styles.xml / heading ${def.level} 段落`,
        fixable: true,
      });
    }
  }

  // 2. Styles must carry the level's own definition, not rely on direct formatting.
  for (const def of spec.headings.levels) {
    const id = resolveStyleId(styles, `heading${def.level}`);
    const s = styles.get(id);
    if (!s) continue;
    const rPr = s.rPr;
    const szEl = rPr ? rPr.children.find((c) => c.name === 'w:sz') : null;
    const styleSize = szEl ? Number(szEl.attr('w:val')) / 2 : null;
    const chainRPr = styles.effectiveRPr(id, null);
    if (chainRPr.sizePt != null && Math.abs(chainRPr.sizePt - def.sizePt) > 0.5) {
      findings.push({
        rule: `headings.${def.level}.styleDef`,
        severity: 'warning',
        title: `Heading ${def.level} 样式本身的字号与规范不一致（文档靠直接格式补救）`,
        expected: `${def.sizePt}pt`,
        actual: `${chainRPr.sizePt}pt（样式层 ${styleSize ?? '未设置'}）`,
        location: `styles.xml style[${id}]`,
        fixable: true,
      });
    }
    if (chainRPr.bold !== def.bold) {
      findings.push({
        rule: `headings.${def.level}.styleBold`,
        severity: 'warning',
        title: `Heading ${def.level} 样式的加粗设置与规范不一致`,
        expected: def.bold ? '加粗' : '不加粗',
        actual: chainRPr.bold ? '加粗' : '不加粗',
        location: `styles.xml style[${id}]`,
        fixable: true,
      });
    }
  }

  // 3. Deeper-than-3 levels used as headings.
  const tooDeep = paras.filter((p) => p.outlineLevel != null && p.outlineLevel >= spec.headings.maxLevel);
  if (tooDeep.length) {
    findings.push({
      rule: 'headings.maxLevel',
      severity: 'warning',
      title: `使用了 ${spec.headings.maxLevel + 1} 级及更深的标题`,
      expected: `全文不超过 ${spec.headings.maxLevel} 级，更深层次用 ${spec.headings.deeperLevelRepresentation.join(' / ')} 在正文内表示`,
      actual: `${tooDeep.length} 处；首例「${tooDeep[0].text.slice(0, 24)}」`,
      detail: { sample: tooDeep.slice(0, 8).map((p) => ({ paragraph: p.index + 1, text: p.text.slice(0, 30), level: p.outlineLevel + 1 })) },
      location: 'word/document.xml',
      fixable: false,
    });
  }

  // 4. Chapter headings must start a new page (§6.3).
  if (spec.sectionBreak.newPagePerLevel1Heading) {
    const h1Id = resolveStyleId(styles, 'heading1');
    const chapters = paras.filter((p) => p.styleId === h1Id && p.text);
    const missing = [];
    for (let i = 0; i < chapters.length; i++) {
      const p = chapters[i];
      if (i === 0) continue; // the first chapter follows the front matter
      const stylePPr = styles.effectivePPr(p.styleId, null);
      const explicitBreak = p.hasPageBreak || p.formatting.pageBreakBefore === true || stylePPr.pageBreakBefore === true;
      if (!explicitBreak) missing.push(p);
    }
    if (missing.length) {
      findings.push({
        rule: 'sectionBreak.newPagePerLevel1Heading',
        severity: 'error',
        title: '一级标题未另起新页',
        expected: '每个一级标题（章）前使用分页符另起一页',
        actual: `${missing.length} 个一级标题前没有分页符；首例「${missing[0].text.slice(0, 24)}」`,
        detail: { sample: missing.slice(0, 8).map((p) => ({ paragraph: p.index + 1, text: p.text.slice(0, 30) })) },
        location: 'word/document.xml 段落分页设置',
        fixable: true,
      });
    }
  }

  // 5. Manually simulated headings: bold, big, standalone, not styled.
  //
  // Cover and front matter are out of scope: a cover title is *supposed* to be
  // large, bold and centred without a Heading style, so classifying it would
  // report every correctly built cover as a manually faked heading.
  const { bodyStart } = buildScope(docx, spec, { resolveStyleId });
  const candidates = [];
  for (const p of paras) {
    if (p.index < bodyStart) continue;
    if (p.styleId && levelIds.includes(p.styleId)) continue;
    const c = classify(p, { bodySizePt: bodySize, maxHeadingChars: 40 });
    if (!c) continue;
    const ev = formatEvidence(p, bodySize);
    const lvl = decideLevel(c, ev);
    candidates.push({ p, candidate: c, ev, decided: lvl });
  }
  if (candidates.length) {
    const high = candidates.filter((c) => c.candidate.confidence === 'high');
    findings.push({
      rule: 'headings.mustUseBuiltinStyles',
      severity: high.length ? 'error' : 'warning',
      title: '存在疑似「手工模拟」的标题（未使用内置 Heading 样式）',
      expected: '标题必须使用 Word 内置 Heading 1/2/3 样式，否则目录无法自动更新',
      actual: `${candidates.length} 段疑似标题未使用 Heading 样式（高置信 ${high.length} 段）`,
      detail: {
        sample: candidates.slice(0, 12).map((c) => ({
          paragraph: c.p.index + 1,
          text: c.p.text.slice(0, 40),
          reason: c.candidate.reason,
          suggestedLevel: c.decided.level,
          confidence: c.candidate.confidence,
          currentStyle: c.p.styleId ?? 'Normal',
        })),
        note: '按 §8.6.4，转换前必须先向用户展示候选列表并确认；不确定的标为待确认。',
      },
      location: 'word/document.xml 正文段落',
      fixable: false,
    });
  }

  return findings;
}

function checkTables(docx, spec) {
  const findings = [];
  const tables = docx.tables();
  const s = spec.table;
  if (!tables.length) return findings;
  // Like the figure captions, the table caption shape is a rule-table value
  // (`table.numberPattern`); the literal default is only a safety net.
  const tblCapPattern = captionPattern(s.numberPattern, /^(表|续表)\s*\d+\s*[-–.]\s*\d+/);

  const badNumbering = [];
  const badCaptionFont = [];
  const badBodyFont = [];
  const badHeaderFont = [];
  const badSize = [];
  const badHeaderBold = [];
  const badStyle = [];

  tables.forEach((t) => {
    // Body cell fonts + sizes. The header row is exempt: it is supposed to be
    // a different font (黑体) and is validated separately below.
    t.rows.forEach((row, rowIndex) => {
      const isHeader = rowIndex === 0;
      for (const cell of row.cells) {
        for (const p of cell.paragraphs) {
          for (const r of p.runs) {
            if (!r.text || !r.text.trim()) continue;
            if (isHeader) {
              if (/[\u4e00-\u9fff]/.test(r.text) && !fontMatches(r.fontEastAsia, s.headerRowFontCN, s.bodyFontCN)) {
                badHeaderFont.push({ table: t, actual: r.fontEastAsia });
              }
              continue;
            }
            if (r.sizePt != null && Math.abs(r.sizePt - s.bodySizePt) > 0.5) badSize.push({ table: t, actual: r.sizePt });
            if (/[\u4e00-\u9fff]/.test(r.text) && !fontMatches(r.fontEastAsia, s.bodyFontCN, null)) {
              badBodyFont.push({ table: t, actual: r.fontEastAsia, expected: s.bodyFontCN });
            }
          }
        }
      }
    });
    if (s.headerRowBold && t.rows[0] && t.headerRowBold === false) badHeaderBold.push(t);
    if (s.style && t.styleId && canonicalStyleName(t.styleName) !== canonicalStyleName(s.style)) badStyle.push(t);

    // Caption: the paragraph immediately above the table (in document order).
  });

  // Caption check needs document order between paragraphs and tables.
  const body = docx.body();
  if (body) {
    const kids = body.children.filter((c) => c.type === 'element' && (c.localName() === 'p' || c.localName() === 'tbl'));
    let tableSeq = 0;
    for (let i = 0; i < kids.length; i++) {
      if (kids[i].localName() !== 'tbl') continue;
      const tbl = tables[tableSeq++];
      // Walk back over empty spacer paragraphs to find the caption line.
      let j = i - 1;
      while (j >= 0 && kids[j].localName() === 'p' && !kids[j].textContent.trim()) j--;
      const capEl = j >= 0 && kids[j].localName() === 'p' ? kids[j] : null;
      const capText = capEl ? capEl.textContent.replace(/\s+/g, ' ').trim() : '';
      const pModel = capEl ? docx.describeParagraph(capEl, -1, 0, docx.styles()) : null;
      // Look a little further up for a heading: a table often follows a short
      // lead-in sentence ("下表列出了…") placed under the section heading.
      const capStyles = docx.styles();
      const capHeadingIds = [1, 2, 3, 4].map((l) => resolveStyleId(capStyles, `heading${l}`));
      const isHeadingModel = (m) =>
        !!m && ((m.styleId && capHeadingIds.includes(m.styleId)) || (m.outlineLevel != null && m.outlineLevel <= 8));
      let precededByHeading = isHeadingModel(pModel);
      for (let k = j - 1, seen = 0; !precededByHeading && k >= 0 && seen < 3; k--) {
        const sib = kids[k];
        if (sib.localName() !== 'p') break;
        if (!sib.textContent.trim()) continue;
        seen++;
        precededByHeading = isHeadingModel(docx.describeParagraph(sib, -1, 0, capStyles));
      }
      if (!tblCapPattern.test(capText)) {
        badNumbering.push({ table: tbl, caption: capText, precededByHeading });
        // No caption line at all: the paragraph above is ordinary body text or
        // a heading, so judging its font/size as a caption would be noise.
        continue;
      }
      const r = pModel?.runs.find((x) => x.text && x.text.trim());
      if (r) {
        if (r.sizePt != null && Math.abs(r.sizePt - s.captionSizePt) > 0.5) badCaptionFont.push({ table: tbl, actual: `${r.sizePt}pt` });
        if (!fontMatches(r.fontEastAsia, s.captionFontCN, null)) badCaptionFont.push({ table: tbl, actual: `字体 ${r.fontEastAsia ?? '继承主题'}` });
        if (pModel.formatting.alignment && pModel.formatting.alignment !== s.captionAlignment) {
          badCaptionFont.push({ table: tbl, actual: `对齐 ${pModel.formatting.alignment}` });
        }
      }
    }
  }

  const headingTitled = badNumbering.filter((b) => b.precededByHeading);
  const trueMissing = badNumbering.filter((b) => !b.precededByHeading);
  if (headingTitled.length) {
    findings.push({
      rule: 'table.caption',
      severity: 'warning',
      title: '表格紧跟在标题下方、以章节标题代替表题',
      expected: `表题在表上方，格式「${s.captionFormat}」，左对齐`,
      actual: `${headingTitled.length} 个表格未单独编号，沿用上方标题`,
      detail: { sample: headingTitled.slice(0, 6).map((b) => ({ table: b.table.index + 1, precedingText: b.caption.slice(0, 30) || '(空)' })) },
      location: 'word/document.xml 表格前置段落',
      fixable: false,
    });
  }
  if (trueMissing.length) {
    findings.push({
      rule: 'table.caption',
      severity: 'error',
      title: '表格缺少合规表题（表 x-x）',
      expected: `表题在表上方，格式「${s.captionFormat}」，左对齐`,
      actual: `${trueMissing.length} 个表格上方没有可识别的表题`,
      detail: { sample: trueMissing.slice(0, 6).map((b) => ({ table: b.table.index + 1, precedingText: b.caption.slice(0, 30) || '(空)' })) },
      location: 'word/document.xml 表格前置段落',
      fixable: false,
    });
  }
  if (badCaptionFont.length) {
    findings.push({
      rule: 'table.captionFormat',
      severity: 'error',
      title: '表题格式不符合规范',
      expected: `${s.captionFontCN} ${s.captionSizePt}pt，左对齐`,
      actual: `${badCaptionFont.length} 处不符；首例：${badCaptionFont[0].actual}`,
      detail: { sample: badCaptionFont.slice(0, 6).map((b) => ({ table: b.table.index + 1, problem: b.actual })) },
      location: 'word/document.xml 表题段落',
      fixable: true,
    });
  }
  if (badSize.length) {
    findings.push({
      rule: 'table.bodySize',
      severity: 'error',
      title: '表格内容字号不符合规范',
      expected: `${s.bodySizePt}pt`,
      actual: `${badSize.length} 处不符；首例 ${badSize[0].actual}pt`,
      detail: { sample: badSize.slice(0, 6).map((b) => ({ table: b.table.index + 1, sizePt: b.actual })) },
      location: 'word/document.xml 表格单元格',
      fixable: true,
    });
  }
  if (badBodyFont.length) {
    findings.push({
      rule: 'table.bodyFont',
      severity: 'error',
      title: '表格中文字体不符合规范',
      expected: s.bodyFontCN,
      actual: `${badBodyFont.length} 处不符；首例 ${badBodyFont[0].actual ?? '继承主题'}`,
      detail: { sample: badBodyFont.slice(0, 6).map((b) => ({ table: b.table.index + 1, font: b.actual ?? '继承主题' })) },
      location: 'word/document.xml 表格单元格',
      fixable: true,
    });
  }
  if (badHeaderFont.length) {
    findings.push({
      rule: 'table.headerRowFont',
      severity: 'warning',
      title: '表格首行（表头）字体与规范不一致',
      expected: s.headerRowFontCN,
      actual: `${badHeaderFont.length} 处不符；首例 ${badHeaderFont[0].actual ?? '继承主题'}`,
      detail: { sample: badHeaderFont.slice(0, 6).map((b) => ({ table: b.table.index + 1, font: b.actual ?? '继承主题' })) },
      location: 'word/document.xml 表格首行',
      fixable: true,
    });
  }
  if (badHeaderBold.length) {
    findings.push({
      rule: 'table.headerRowBold',
      severity: 'error',
      title: '表格首行（表头）未加粗',
      expected: '表头行黑体加粗',
      actual: `${badHeaderBold.length} 个表格首行未加粗`,
      location: 'word/document.xml 表格首行',
      fixable: true,
    });
  }
  if (badStyle.length) {
    findings.push({
      rule: 'table.style',
      severity: 'warning',
      title: '表格样式与规范建议不一致',
      expected: `${s.style} 或同等线框样式`,
      actual: `${badStyle.length} 个表格样式不符：${badStyle.map((t) => t.styleName).join('、')}`,
      location: 'word/document.xml tblStyle',
      fixable: true,
    });
  }
  return findings;
}

function checkFigures(docx, spec) {
  const findings = [];
  const paras = docx.paragraphs();
  const s = spec.figure;
  const withDrawing = paras.filter((p) => p.hasDrawing);
  if (!withDrawing.length) return findings;

  const noCaption = [];
  const exactSpacing = [];
  const notCentered = [];
  const captionBad = [];
  const lowDpi = [];
  const floating = [];
  const below = s.captionPosition !== 'above';
  // The rule table owns what a caption looks like; the literal 「图 1-1」 fallback
  // only covers a table that left numberPattern out.
  const capPattern = captionPattern(s.numberPattern, /^图\s*\d+[-.]\d+/);

  const captionOf = (p) => {
    // The caption belongs on the side the rule table names. Scanning only
    // downward made `captionPosition: "above"` impossible to satisfy: a
    // perfectly captioned figure would be reported as uncaptioned forever.
    const step = below ? 1 : -1;
    for (let k = p.index + step; k >= 0 && k < paras.length; k += step) {
      if (paras[k].hasDrawing) return null; // another image: nothing in between
      if (paras[k].text.trim()) return paras[k];
    }
    return null;
  };

  for (const p of withDrawing) {
    const next = captionOf(p);
    const capText = next?.text ?? '';
    if (!next || !capPattern.test(capText)) noCaption.push(p);
    else {
      const r = next.runs.find((x) => x.text && x.text.trim());
      if (r) {
        if (r.sizePt != null && Math.abs(r.sizePt - s.captionSizePt) > 0.5) captionBad.push({ p: next, actual: `${r.sizePt}pt` });
        if (!fontMatches(r.fontEastAsia, s.captionFontCN, null)) captionBad.push({ p: next, actual: `字体 ${r.fontEastAsia ?? '继承主题'}` });
        if (next.formatting.alignment !== s.captionAlignment) captionBad.push({ p: next, actual: `对齐 ${next.formatting.alignment ?? '默认'}` });
        // A run with no w:color inherits and renders black, so only a non-black
        // rule can be violated by absence.
        const wantCapColor = typeof s.color === 'string' ? s.color.trim().toUpperCase() : '';
        const capColor = typeof r.color === 'string' ? r.color.trim().toUpperCase() : '';
        if (wantCapColor !== '' && capColor !== wantCapColor && !(capColor === '' && wantCapColor === '000000')) {
          captionBad.push({ p: next, actual: `颜色 ${capColor === '' ? '继承主题' : '#' + capColor}` });
        }
      }
    }
    const ls = p.formatting.lineSpacing;
    if (s.paragraphLineSpacingMustNotBeExact && ls && ls.type === 'exact') {
      exactSpacing.push({ p, actual: `固定 ${round(ls.pt, 1)} 磅` });
    }
    if (p.formatting.alignment !== s.paragraphAlignment) notCentered.push({ p, actual: p.formatting.alignment ?? '默认' });
  }

  // Resolution and embedding are properties of the picture itself, not of its
  // paragraph, so they are read from the drawing markup and the media bytes.
  for (const fig of docx.figures()) {
    if (s.embedding === 'inline' && !fig.inline) {
      floating.push({ p: fig.paragraph, actual: fig.part ?? '无法解析图片' });
    }
    if (Number.isFinite(s.minDpi) && fig.dpi !== null) {
      const dpi = Math.min(fig.dpi.dpiX, fig.dpi.dpiY);
      if (dpi + 0.5 < s.minDpi) {
        lowDpi.push({
          p: fig.paragraph,
          actual: `${round(dpi, 0)} dpi（${fig.dpi.widthPx}×${fig.dpi.heightPx} px 显示为 ${round(fig.dpi.widthCm, 1)}×${round(fig.dpi.heightCm, 1)} cm）`,
        });
      }
    }
  }

  if (noCaption.length) {
    findings.push({
      rule: 'figure.caption',
      severity: 'error',
      title: '插图缺少图题（图 x-x）',
      expected: `图题在图${below ? '下' : '上'}方、居中、格式「图 x-x 标题」`,
      actual: `${noCaption.length} 张图${below ? '下' : '上'}方没有图题`,
      detail: { sample: noCaption.slice(0, 6).map((p) => ({ paragraph: p.index + 1 })) },
      location: 'word/document.xml 图片段落',
      fixable: false,
    });
  }
  if (captionBad.length) {
    findings.push({
      rule: 'figure.captionFormat',
      severity: 'error',
      title: '图题格式不符合规范',
      expected: `${s.captionFontCN} ${s.captionSizePt}pt，居中`,
      actual: `${captionBad.length} 处不符；首例：${captionBad[0].actual}`,
      detail: { sample: captionBad.slice(0, 6).map((b) => ({ paragraph: b.p.index + 1, problem: b.actual })) },
      location: 'word/document.xml 图题段落',
      fixable: true,
    });
  }
  if (exactSpacing.length) {
    findings.push({
      rule: 'figure.lineSpacing',
      severity: 'error',
      title: '图片段落使用了固定行距（会裁切图片）',
      expected: '至少 1.5 倍行距或单倍行距',
      actual: `${exactSpacing.length} 处使用固定行距；首例 ${exactSpacing[0].actual}`,
      detail: { sample: exactSpacing.slice(0, 6).map((e) => ({ paragraph: e.p.index + 1, actual: e.actual })) },
      location: 'word/document.xml 图片段落 pPr/spacing',
      fixable: true,
    });
  }
  if (notCentered.length) {
    findings.push({
      rule: 'figure.alignment',
      severity: 'warning',
      title: '图片段落未居中',
      expected: '居中',
      actual: `${notCentered.length} 处未居中`,
      location: 'word/document.xml 图片段落 pPr/jc',
      fixable: true,
    });
  }
  if (lowDpi.length) {
    findings.push({
      rule: 'figure.minDpi',
      severity: 'warning',
      title: '插图分辨率低于规范要求',
      expected: `≥ ${s.minDpi} dpi（建议 ${(s.recommendedDpiRange ?? []).join('–')}）`,
      actual: `${lowDpi.length} 张图低于下限；首例 ${lowDpi[0].actual}`,
      detail: { sample: lowDpi.slice(0, 6).map((d) => ({ paragraph: d.p.index + 1, actual: d.actual })) },
      location: 'word/media/* 与 word/document.xml wp:extent',
      fixable: false,
    });
  }
  if (floating.length) {
    findings.push({
      rule: 'figure.embedding',
      severity: 'warning',
      title: '插图不是嵌入式（inline）',
      expected: '嵌入方式 inline',
      actual: `${floating.length} 张图为浮动放置（wp:anchor）`,
      detail: { sample: floating.slice(0, 6).map((f) => ({ paragraph: f.p.index + 1 })) },
      location: 'word/document.xml wp:anchor',
      fixable: false,
    });
  }
  return findings;
}

function checkCodeBlocks(docx, spec) {
  const findings = [];
  const paras = docx.paragraphs();
  const s = spec.code;
  // Heuristic: a paragraph whose every run is Courier New, or that looks like
  // ST/structured code and is left-indented with no first-line indent.
  const codeParas = paras.filter((p) => {
    if (!p.text.trim() || p.inTable) return false;
    const runs = p.runs.filter((r) => r.text && r.text.trim());
    if (!runs.length) return false;
    const courier = runs.every((r) => /courier/i.test(r.fontAscii ?? '') || /courier/i.test(r.fontHAnsi ?? ''));
    if (courier) return true;
    return false;
  });
  if (!codeParas.length) return findings;

  const badFont = [];
  const badSize = [];
  const badIndent = [];
  const badSpacing = [];
  for (const p of codeParas) {
    for (const r of p.runs) {
      if (!r.text.trim()) continue;
      if (!/^courier new$/i.test(String(r.fontAscii ?? r.fontHAnsi ?? '').trim())) badFont.push({ p, actual: r.fontAscii ?? r.fontHAnsi });
      if (r.sizePt != null && (r.sizePt < s.sizePtRange[0] || r.sizePt > s.sizePtRange[1])) badSize.push({ p, actual: r.sizePt });
    }
    const left = p.formatting.leftIndentTwip != null ? round(p.formatting.leftIndentTwip / TWIP_PER_CM, 2) : null;
    if (left == null || Math.abs(left - s.indentLeftCm) > 0.15) badIndent.push({ p, actual: left == null ? '无左缩进' : `${left}cm` });
    const ls = p.formatting.lineSpacing;
    const wantCode = lineSpacingTarget(s.lineSpacing);
    const okSpacing = ls != null && (wantCode.unit === 'pt'
      ? ls.type === wantCode.rule && Math.abs((ls.pt ?? 0) - wantCode.value) <= 0.5
      : ls.type === 'auto' && Math.abs((ls.lines ?? 0) - wantCode.value) <= 0.05);
    if (!okSpacing) {
      badSpacing.push({ p, actual: ls ? (ls.type === 'auto' ? `${round(ls.lines ?? 0, 2)} 倍` : `${ls.type} ${round(ls.pt, 1)}磅`) : '未设置' });
    }
  }
  const add = (rule, title, expected, bad, key) => {
    if (!bad.length) return;
    findings.push({
      rule,
      severity: 'error',
      title,
      expected,
      actual: `${bad.length} 处不符；首例「${bad[0].p.text.slice(0, 24)}」：${bad[0][key] ?? bad[0].actual}`,
      detail: { sample: bad.slice(0, 6).map((b) => ({ paragraph: b.p.index + 1, text: b.p.text.slice(0, 30), actual: b[key] ?? b.actual })) },
      location: 'word/document.xml 代码段落',
      fixable: true,
    });
  };
  add('code.font', '代码块字体不符合规范', s.font, badFont, 'actual');
  add('code.size', '代码块字号不符合规范', `${s.sizePtRange[0]}~${s.sizePtRange[1]}pt`, badSize, 'actual');
  add('code.indent', '代码块左缩进不符合规范', `左缩进 ${s.indentLeftCm}cm、无首行缩进`, badIndent, 'actual');
  add('code.lineSpacing', '代码块行距不符合规范', lineSpacingText(s.lineSpacing), badSpacing, 'actual');
  return findings;
}

const NUMERAL_TEXT = {
  decimal: '阿拉伯数字 1, 2, 3',
  lowerRoman: '罗马小写 i, ii, iii',
  upperRoman: '罗马大写 I, II, III',
  lowerLetter: '字母小写 a, b, c',
  upperLetter: '字母大写 A, B, C',
};

const ALIGN_CN = { left: '左对齐', center: '居中', right: '右对齐', both: '两端对齐' };

function pageNumberText(s) {
  const numeral = NUMERAL_TEXT[s.pageNumberNumeral ?? 'decimal'] ?? s.pageNumberNumeral ?? '阿拉伯数字';
  const parts = pageNumberParts(s.pageNumberFormat);
  return `${parts.prefix}${numeral}${parts.suffix}`;
}

function checkHeaderFooter(docx, spec) {
  const findings = [];
  const s = spec.headerFooter;
  const parts = docx.headerFooterParts();

  const headers = parts.filter((p) => p.kind === 'header');
  const footers = parts.filter((p) => p.kind === 'footer');

  for (const h of headers) {
    const ps = docx.describePartParagraphs(h.target);
    const r = ps.flatMap((p) => p.runs).find((x) => x.text && x.text.trim());
    const p0 = ps[0];
    if (r) {
      const problems = [];
      if (r.sizePt != null && Math.abs(r.sizePt - s.headerSizePt) > 0.5) problems.push(`字号 ${r.sizePt}pt（应 ${s.headerSizePt}pt）`);
      if (!fontMatches(r.fontEastAsia, s.headerFontCN, null)) problems.push(`中文字体 ${r.fontEastAsia ?? '继承主题'}（应 ${s.headerFontCN}）`);
      if (p0 && p0.formatting.alignment !== s.headerAlignment) problems.push(`对齐 ${p0.formatting.alignment ?? '默认'}（应居中）`);
      if (s.headerColor && r.color && String(r.color).toLowerCase() !== s.headerColor.toLowerCase()) problems.push(`颜色 ${r.color}（应 ${s.headerColor}）`);
      if (problems.length) {
        findings.push({
          rule: 'headerFooter.header',
          severity: 'error',
          title: '页眉格式不符合规范',
          expected: `${s.headerFontCN} ${s.headerSizePt}pt、灰色、居中`,
          actual: `${h.target}：${problems.join('；')}`,
          location: h.target,
          fixable: true,
        });
      }
    }
  }

  if (!footers.length) {
    findings.push({
      rule: 'headerFooter.pageNumber',
      severity: 'error',
      title: '文档没有页脚（缺页码）',
      expected: `页码格式「${pageNumberText(s)}」${ALIGN_CN[s.pageNumberAlignment] ?? s.pageNumberAlignment}`,
      actual: '未找到页脚部件',
      location: 'word/document.xml sectPr/footerReference',
      fixable: false,
    });
  } else {
    const noPageField = [];
    const partsToCheck = s.pageNumberPosition === 'header' ? headers : footers;
    // One string describes the whole decorated number ("— 数字 —"); splitting
    // it into prefix/suffix meant fixing the same visual in two boxes and made
    // half-done states ("— " with no closing dash) look saved.
    const { prefix, suffix } = pageNumberParts(s.pageNumberFormat);
    for (const f of partsToCheck) {
      const ps = docx.describePartParagraphs(f.target);
      const hasPage = docx.root(f.target)?.findAll((e) => e.localName() === 'instrText').some((t) => /PAGE/i.test(t.textContent)) ?? false;
      if (!hasPage) noPageField.push(f.target);
      const p0 = ps[0];
      if (p0 && p0.formatting.alignment !== s.pageNumberAlignment) {
        findings.push({
          rule: 'headerFooter.pageNumberAlign',
          severity: 'warning',
          title: `页码未${ALIGN_CN[s.pageNumberAlignment] ?? '居中'}`,
          expected: ALIGN_CN[s.pageNumberAlignment] ?? s.pageNumberAlignment,
          actual: `${f.target} 对齐为 ${p0.formatting.alignment ?? '默认'}`,
          location: f.target,
          fixable: true,
        });
      }
      // Prefix/suffix are what the reader actually sees around the number;
      // a rule asking for "第 1 页" but a footer rendering "— 1 —" is wrong
      // even though the PAGE field itself is fine.
      if (hasPage && p0) {
        const text = ps.map((x) => x.rawText ?? x.text ?? '').join('');
        const decorated = `${prefix}${suffix}` !== '' || text !== '';
        if (decorated && prefix && !text.includes(prefix)) {
          findings.push({
            rule: 'headerFooter.pageNumberDecoration',
            severity: 'info',
            title: '页码缺少前缀',
            expected: `页码前应有「${prefix}」`,
            actual: text.slice(0, 20) || '(空)',
            location: f.target,
            fixable: false,
          });
        }
        if (decorated && suffix && !text.includes(suffix)) {
          findings.push({
            rule: 'headerFooter.pageNumberDecoration',
            severity: 'info',
            title: '页码缺少后缀',
            expected: `页码后应有「${suffix}」`,
            actual: text.slice(0, 20) || '(空)',
            location: f.target,
            fixable: false,
          });
        }
      }
    }
    if (noPageField.length) {
      findings.push({
        rule: 'headerFooter.pageNumber',
        severity: 'error',
        title: `没有找到 PAGE 域（页码不会自动更新）`,
        expected: `页码使用 PAGE 域，形如「${pageNumberText(s)}」`,
        actual: noPageField.join('、'),
        location: noPageField.join('、'),
        fixable: false,
      });
    }
  }

  // Numeral style of the page numbers actually written by Word (`pgNumType`).
  //
  // The body section is the one that matters: a document may legitimately use
  // lowerRoman for the abstract and then decimal for the body, so testing
  // "some section uses the wanted numeral" would pass a document whose body
  // pages came out as i, ii, iii while the rule asks for 1, 2, 3.
  {
    const sections = docx.sectionInfo();
    const bodySection = sections.length > 1 ? sections[sections.length - 1] : sections[0];
    const actual = bodySection?.pgNumType?.fmt ?? 'decimal';
    const want = s.pageNumberNumeral ?? 'decimal';
    if (bodySection && actual !== want) {
      findings.push({
        rule: 'headerFooter.pageNumberNumeral',
        severity: 'warning',
        title: '正文页码数字样式不符合规范',
        expected: `${NUMERAL_TEXT[want] ?? want}（pgNumType=${want}）`,
        actual: `最后一节为 ${NUMERAL_TEXT[actual] ?? actual}（pgNumType=${actual}）`,
        location: 'sectPr/pgNumType',
        fixable: true,
      });
    }
  }

  // Header suppression on cover/front matter: needs titlePg on the first
  // section, or a separate section for the front matter.
  if (s.hideHeaderOnCoverAndAbstract) {
    const sections = docx.sectionInfo();
    const first = sections[0];
    const hasHeaderOnFirst = first && first.headerRefs.some((r) => r.type === 'default');
    if (first && hasHeaderOnFirst && !first.titlePg) {
      findings.push({
        rule: 'headerFooter.hideHeaderOnCoverAndAbstract',
        severity: 'warning',
        title: '封面/摘要页仍显示页眉',
        expected: '封面与摘要页不显示页眉（首节启用「首页不同」或单独分节）',
        actual: `第 1 节有页眉引用且 titlePg=${first.titlePg}`,
        location: 'sectPr[0]',
        fixable: true,
      });
    }
  }

  // Front matter pagination: roman numerals for abstract/TOC section.
  const sections = docx.sectionInfo();
  const hasRoman = sections.some((x) => x.pgNumType && /roman|upperRoman|lowerRoman/i.test(x.pgNumType.fmt));
  if (!hasRoman && sections.length > 1) {
    findings.push({
      rule: 'headerFooter.frontMatterNumeral',
      severity: 'warning',
      title: '摘要/目录未使用罗马数字单独编页',
      expected: '封面不编页码，摘要与目录用罗马数字，正文从 1 起',
      actual: sections.map((x) => x.pgNumType?.fmt ?? 'decimal').join(' / '),
      location: 'sectPr/pgNumType',
      fixable: true,
    });
  }
  return findings;
}

function checkReferences(docx, spec) {
  const findings = [];
  const paras = docx.paragraphs();
  const s = spec.reference;
  const styles = docx.styles();
  const headingIds = [1, 2, 3].map((l) => resolveStyleId(styles, `heading${l}`));
  const refIdx = paras.findIndex((p) => /^\s*参考文献\s*$/.test(p.text) && headingIds.includes(p.styleId));
  if (refIdx < 0) return findings;

  const list = [];
  for (let i = refIdx + 1; i < paras.length; i++) {
    const p = paras[i];
    if (p.styleId && headingIds.includes(p.styleId)) break;
    if (!p.text) continue;
    if (/^\s*附录|^\s*致\s*谢/.test(p.text)) break;
    list.push(p);
  }
  if (!list.length) {
    findings.push({
      rule: 'reference.list',
      severity: 'error',
      title: '参考文献章节没有条目',
      expected: '按 GB/T 7714-2015 列出文献',
      actual: '空',
      location: 'word/document.xml',
      fixable: false,
    });
    return findings;
  }

  const bad = [];
  for (const p of list) {
    const t = p.text;
    const typeTag = /\[(J|M|S|EB\/OL|C|D|N|P|R|G)\]/i.exec(t)?.[1]?.toUpperCase();
    if (!/^\[\d+\]/.test(t)) {
      bad.push({ p, problem: '缺少「[序号]」起始' });
      continue;
    }
    if (!typeTag) bad.push({ p, problem: '缺少文献类型标识（如 [J]/[M]/[S]/[EB/OL]）' });
    // Font + size per spec.
    const r = p.runs.find((x) => x.text && x.text.trim());
    if (r) {
      if (r.sizePt != null && Math.abs(r.sizePt - s.sizePt) > 0.5) bad.push({ p, problem: `字号 ${r.sizePt}pt（应 ${s.sizePt}pt）` });
      if (!fontMatches(r.fontEastAsia, s.fontCN, null)) bad.push({ p, problem: `中文字体 ${r.fontEastAsia ?? '继承主题'}（应 ${s.fontCN}）` });
    }
  }
  if (bad.length) {
    findings.push({
      rule: 'reference.format',
      severity: 'error',
      title: '参考文献条目格式不符合规范',
      expected: `按 ${s.listStandard}：期刊/专著/标准/电子文献各有模板；仿宋_GB2312 ${s.sizePt}pt`,
      actual: `${bad.length} 处不符；首例「${bad[0].p.text.slice(0, 40)}」：${bad[0].problem}`,
      detail: { sample: bad.slice(0, 10).map((b) => ({ paragraph: b.p.index + 1, text: b.p.text.slice(0, 60), problem: b.problem })) },
      location: 'word/document.xml 参考文献列表',
      fixable: false,
    });
  }

  // In-text superscript citations.
  const bodyText = paras.filter((p) => !p.inTable && p.text && p.index < refIdx && !(p.styleId && headingIds.includes(p.styleId)));
  const citations = [];
  for (const p of bodyText) {
    for (const r of p.runs) {
      if (/\[\d+(?:[,-\s]\s*\d+)*\]/.test(r.text)) {
        citations.push({ p, run: r, superscript: p.text.includes('[') ? isSuperscript(docx, r) : false, text: r.text });
      }
    }
  }
  const notSuper = citations.filter((c) => !c.superscript);
  if (citations.length && notSuper.length) {
    findings.push({
      rule: 'reference.inTextCitation',
      severity: 'warning',
      title: '正文引用未使用上标',
      expected: '正文引用用方括号上标，如 [1]、[2,3]、[4-6]',
      actual: `${notSuper.length}/${citations.length} 处引用不是上标；首例「${notSuper[0].text.slice(0, 30)}」`,
      detail: { sample: notSuper.slice(0, 6).map((c) => ({ paragraph: c.p.index + 1, text: c.text.slice(0, 40) })) },
      location: 'word/document.xml 正文 run',
      fixable: false,
    });
  }
  return findings;
}

function isSuperscript(docx, run) {
  if (!run.rPr) return false;
  const va = run.rPr.children.find((c) => c.localName() === 'vertAlign');
  const v = va?.attr('w:val');
  return v === 'superscript';
}

function checkAppendixNumbering(docx, spec) {
  const findings = [];
  const paras = docx.paragraphs();
  const styles = docx.styles();
  const h1 = resolveStyleId(styles, 'heading1');
  const gap = spec.headings.numberingGapSpaces;
  const bad = [];
  for (const p of paras) {
    if (p.styleId !== h1) continue;
    const raw = String(p.rawText ?? '').replace(/^\s+/, '');
    if (!/^附\s*录/.test(raw)) continue;
    if (!/^附\s*录\s*[A-Z]/.test(raw)) {
      bad.push({ p, problem: '附录编号缺失或不规范' });
      continue;
    }
    const spaces = /^附\s*录\s*[A-Z](\s*)/.exec(raw)?.[1]?.length ?? 0;
    if (spaces !== gap) bad.push({ p, problem: `附录编号后空格 ${spaces} 个（应 ${gap} 个）` });
  }
  if (bad.length) {
    findings.push({
      rule: 'headings.numberingGap',
      severity: 'error',
      title: '附录标题编号空格数量不符',
      expected: `编号与文字之间空 ${gap} 个半角空格`,
      actual: `${bad.length} 处不符；首例「${bad[0].p.text.slice(0, 24)}」：${bad[0].problem}`,
      detail: { sample: bad.slice(0, 6).map((b) => ({ paragraph: b.p.index + 1, text: b.p.text.slice(0, 30), problem: b.problem })) },
      location: 'word/document.xml 附录标题',
      fixable: true,
    });
  }
  return findings;
}

function checkTocPollution(docx, spec) {
  const findings = [];
  const paras = docx.paragraphs();
  const styles = docx.styles();
  const h1 = resolveStyleId(styles, 'heading1');
  if (!h1) return findings;
  const polluted = paras.filter((p) => p.styleId === h1 && (looksLikeFrontMatter(p.text) || !p.text.trim()));
  if (polluted.length) {
    findings.push({
      rule: 'headings.tocPollution',
      severity: 'warning',
      title: '前置元数据使用了 Heading 1 样式（会污染目录）',
      expected: '摘要、目录、关键词等不应使用 Heading 样式，或应排除在目录之外',
      actual: `${polluted.length} 段：${polluted.map((p) => p.text || '(空段落)').slice(0, 6).join('、')}`,
      detail: { sample: polluted.slice(0, 8).map((p) => ({ paragraph: p.index + 1, text: p.text || '(空)', style: p.styleId })) },
      location: 'word/document.xml',
      fixable: false,
    });
  }
  return findings;
}

function checkFileNaming(docx, spec) {
  const findings = [];
  const s = spec.aiSupplement;
  const path = docx.path ?? '';
  const base = path.split(/[\\/]/).pop() ?? '';
  if (!base) return findings;
  if (s.fileNamingForbidPureChinese && /^[\u4e00-\u9fff\s（）()、，。]+\.docx$/i.test(base)) {
    findings.push({
      rule: 'aiSupplement.fileNaming',
      severity: 'warning',
      title: '文件名是纯中文，不符合 AI 补充规定',
      expected: s.fileNamingPattern.replace('{文档类型}', '文档类型').replace('{日期}', '日期') + `（示例 ${s.fileNamingExample}）`,
      actual: base,
      location: path,
      fixable: false,
    });
  }
  return findings;
}

/* ------------------------------------------------------------------ */
/* Orchestration                                                       */
/* ------------------------------------------------------------------ */

/**
 * Cover and TOC structure.
 *
 * These are not "is the font right" checks — they catch the cases where the
 * settings page promises something the document does not deliver: a cover the
 * rule table asked for that has no title at all, or a TOC setting turned on
 * while the document has no TOC field (so the reader sees nothing, no matter
 * what the page says).
 */
function checkCoverAndToc(docx, spec) {
  const findings = [];
  const paras = docx.paragraphs();
  const styles = docx.styles();
  const cover = spec.cover ?? {};
  const toc = spec.toc ?? {};

  // --- cover ---------------------------------------------------------
  {
    const explicitTitle = String(cover.title ?? '').trim();
    // The cover is everything before the first level-1 heading, or the first
    // 12 paragraphs when the document has no headings at all.
    const firstHeadingIdx = paras.findIndex((p) => p.styleId === resolveStyleId(styles, 'heading1'));
    const coverEnd = firstHeadingIdx > 0 ? firstHeadingIdx : Math.min(paras.length, 12);
    const coverParas = paras.slice(0, coverEnd);

    // Which paragraph is "the title"?
    //
    // Only an explicit `cover.title` names a title to look for; falling back
    // to `meta.title` would compare every document against the NAME OF THE
    // SPECIFICATION, and `includes` then matched a line like
    // 「（遵照《AI 辅助开发文档格式规范》生成）」 — a 12pt reference, not the
    // 26pt title — and reported the real title as badly formatted. When no
    // title is configured the largest type on the cover is the title by
    // definition.
    let titlePara = null;
    if (explicitTitle !== '') {
      titlePara = coverParas.find((p) => p.text.trim() === explicitTitle)
        ?? coverParas.find((p) => p.text.includes(explicitTitle));
      if (!titlePara) {
        findings.push({
          rule: 'cover.title',
          severity: 'warning',
          title: '封面没有规则表指定的标题',
          expected: `封面应出现「${explicitTitle}」`,
          actual: `前 ${coverEnd} 段中未找到该文本`,
          location: 'word/document.xml 开头段落',
          fixable: false,
        });
      }
    } else {
      let best = null;
      for (const p of coverParas) {
        const size = Math.max(0, ...p.runs.map((r) => r.sizePt ?? 0));
        if (size > 0 && (best === null || size > best.size)) best = { p, size };
      }
      titlePara = best?.p ?? null;
    }

    if (cover.showProject && spec.meta?.project) {
      const hasProject = coverParas.some((p) => p.text.includes(String(spec.meta.project)));
      if (!hasProject) {
        findings.push({
          rule: 'cover.project',
          severity: 'info',
          title: '封面没有系统/项目名',
          expected: `封面应出现「${spec.meta.project}」`,
          actual: '未找到',
          location: 'word/document.xml 开头段落',
          fixable: false,
        });
      }
    }
    if (titlePara) {
      const run = titlePara.runs.find((r) => r.text && r.text.trim());
      if (run) {
        const problems = [];
        if (cover.titleSizePt != null && run.sizePt != null && Math.abs(run.sizePt - cover.titleSizePt) > 0.5) {
          problems.push(`字号 ${run.sizePt}pt（应 ${cover.titleSizePt}pt）`);
        }
        if (cover.titleBold != null && cover.titleBold !== (run.bold === true)) {
          problems.push(`加粗 ${run.bold === true ? '是' : '否'}（应 ${cover.titleBold ? '是' : '否'}）`);
        }
        if (cover.titleFontCN && !fontMatches(run.fontEastAsia, cover.titleFontCN, null)) {
          problems.push(`中文字体 ${run.fontEastAsia ?? '继承主题'}（应 ${cover.titleFontCN}）`);
        }
        if (cover.alignment && titlePara.formatting.alignment && titlePara.formatting.alignment !== cover.alignment) {
          problems.push(`对齐 ${titlePara.formatting.alignment}（应 ${ALIGN_CN[cover.alignment] ?? cover.alignment}）`);
        }
        // An absent w:color renders as black, so only a configured non-black
        // colour can be violated by inheritance.
        const wantCoverColor = String(cover.color ?? '').trim();
        if (/^[0-9A-Fa-f]{6}$/.test(wantCoverColor)) {
          const actualColor = (run.color ?? '000000').replace(/^#/, '').toUpperCase();
          if (actualColor !== wantCoverColor.toUpperCase()) {
            problems.push(`颜色 ${actualColor}（应 ${wantCoverColor.toUpperCase()}）`);
          }
        }
        if (problems.length) {
          findings.push({
            rule: 'cover.titleFormat',
            severity: 'warning',
            title: '封面标题格式不符合设置',
            expected: `${cover.titleFontCN ?? ''} ${cover.titleSizePt ?? ''}pt${cover.titleBold ? ' 加粗' : ''}`,
            actual: problems.join('；'),
            location: `段落 #${titlePara.index + 1}`,
            fixable: true,
          });
        }
      }
    }
  }

  // --- TOC -----------------------------------------------------------
  if (toc.enabled) {
    const hasTocField = docx.root('word/document.xml')?.findAll((e) => e.localName() === 'instrText').some((t) => /TOC/i.test(t.textContent)) ?? false;
    const tocTitle = String(toc.title ?? '').trim();
    const hasTocHeading = tocTitle === '' ? false : paras.some((p) => p.text.trim() === tocTitle);
    if (!hasTocField) {
      findings.push({
        rule: 'toc.field',
        severity: 'warning',
        title: '规则表启用了目录，但文档里没有 TOC 域',
        expected: '插入 TOC 域（打开 Word 后按 F9 更新）',
        actual: '未找到 instrText 含 TOC 的域',
        location: 'word/document.xml',
        fixable: false,
      });
    }
    if (!hasTocHeading) {
      findings.push({
        rule: 'toc.title',
        severity: 'info',
        title: '文档中没有目录标题',
        expected: `应有「${tocTitle || '目录'}」标题段`,
        actual: '未找到',
        location: 'word/document.xml',
        fixable: false,
      });
    } else {
      // Style of the TOC heading: a title the writer forgot to format is the
      // common case, and it is cheap to point at.
      const heading = paras.find((p) => p.text.trim() === tocTitle);
      const run = heading?.runs.find((r) => r.text && r.text.trim());
      if (run) {
        const problems = [];
        if (toc.titleSizePt != null && run.sizePt != null && Math.abs(run.sizePt - toc.titleSizePt) > 0.5) {
          problems.push(`字号 ${run.sizePt}pt（应 ${toc.titleSizePt}pt）`);
        }
        if (toc.titleBold != null && toc.titleBold !== (run.bold === true)) {
          problems.push(`加粗 ${run.bold === true ? '是' : '否'}（应 ${toc.titleBold ? '是' : '否'}）`);
        }
        if (toc.titleFontCN && !fontMatches(run.fontEastAsia, toc.titleFontCN, null)) {
          problems.push(`中文字体 ${run.fontEastAsia ?? '继承主题'}（应 ${toc.titleFontCN}）`);
        }
        const wantTocColor = String(toc.titleColor ?? '').trim();
        if (/^[0-9A-Fa-f]{6}$/.test(wantTocColor)) {
          const actualColor = (run.color ?? '000000').replace(/^#/, '').toUpperCase();
          if (actualColor !== wantTocColor.toUpperCase()) {
            problems.push(`颜色 ${actualColor}（应 ${wantTocColor.toUpperCase()}）`);
          }
        }
        if (problems.length) {
          findings.push({
            rule: 'toc.titleFormat',
            severity: 'warning',
            title: '目录标题格式不符合设置',
            expected: `${toc.titleFontCN ?? ''} ${toc.titleSizePt ?? ''}pt${toc.titleBold ? ' 加粗' : ''}`,
            actual: problems.join('；'),
            location: `段落 #${heading.index + 1}`,
            fixable: true,
          });
        }
      }
    }
  }

  return findings;
}

const CHECK_GROUPS = [
  ['page', checkPageSetup],
  ['body', checkBodyFormat],
  ['headings', checkHeadings],
  ['table', checkTables],
  ['figure', checkFigures],
  ['code', checkCodeBlocks],
  ['headerFooter', checkHeaderFooter],
  ['reference', checkReferences],
  ['appendix', checkAppendixNumbering],
  ['cover', checkCoverAndToc],
  ['toc', checkTocPollution],
  ['naming', checkFileNaming],
];

/**
 * Run the audit.
 * @param {import('./docx.js').Docx} docx
 * @param {object} spec
 * @param {{only?: string[]}} [options] restrict to rule prefixes
 * @returns {{findings: object[], summary: object, stats: object}}
 */
export function audit(docx, spec, options = {}) {
  const findings = [];
  const errors = [];
  for (const [group, fn] of CHECK_GROUPS) {
    if (options.only && !options.only.some((o) => group.startsWith(o) || o.startsWith(group))) continue;
    try {
      findings.push(...fn(docx, spec));
    } catch (err) {
      errors.push({ group, message: err.message, stack: err.stack });
    }
  }

  const bySeverity = { error: 0, warning: 0, info: 0 };
  for (const f of findings) bySeverity[f.severity] = (bySeverity[f.severity] ?? 0) + 1;

  const paras = docx.paragraphs();
  const tables = docx.tables();
  const styles = docx.styles();
  const headingIds = [1, 2, 3].map((l) => resolveStyleId(styles, `heading${l}`));
  const headingCounts = { 1: 0, 2: 0, 3: 0 };
  for (const p of paras) {
    const i = headingIds.indexOf(p.styleId);
    if (i >= 0) headingCounts[i + 1]++;
  }

  const stats = {
    paragraphs: paras.length,
    nonEmptyParagraphs: paras.filter((p) => p.text).length,
    tables: tables.length,
    figures: paras.filter((p) => p.hasDrawing).length,
    sections: docx.sectionInfo().length,
    headings: headingCounts,
    styles: styles.byId.size,
    parts: docx.order.length,
    bodySizePt: dominantBodySize(paras),
  };

  const checklist = (spec.checklist ?? []).map((item) => {
    const related = findings.filter((f) => f.rule === item.rule || f.rule.startsWith(item.rule + '.'));
    const failed = related.filter((f) => f.severity === 'error');
    const warned = related.filter((f) => f.severity === 'warning');
    return {
      id: item.id,
      label: checklistLabel(item, spec),
      rule: item.rule,
      status: failed.length ? 'fail' : warned.length ? 'warn' : 'pass',
      findings: related.map((f) => f.title),
    };
  });

  return {
    findings,
    errors,
    stats,
    summary: {
      ...bySeverity,
      total: findings.length,
      passed: checklist.filter((c) => c.status === 'pass').length,
      warned: checklist.filter((c) => c.status === 'warn').length,
      failed: checklist.filter((c) => c.status === 'fail').length,
      verdict: bySeverity.error ? '不通过' : bySeverity.warning ? '基本通过（有警告）' : '通过',
    },
    checklist,
  };
}

export { textHeadingCandidate, formatEvidence };
