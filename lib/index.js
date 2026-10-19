/**
 * dsh-word-writing-spec — Host half.
 *
 * Two jobs, both driven by the same rule table (`rules/docx-writing-spec.json`):
 *
 * 1. Constrain generation. One ordered `systemPrompt` section carries the
 *    normative rules (page setup, fonts, heading hierarchy, tables, figures,
 *    code blocks, header/footer, references, appendix, naming, backup) so a
 *    document produced in this session is written to spec from the first
 *    character instead of being repaired afterwards.
 *
 * 2. Review and repair. Tools that open a real .docx, audit it against the same
 *    rule table, and optionally write a fixed copy — always after backing up
 *    the original (§ AI 补充约定: backupRequiredBeforeEdit).
 *
 * Why the prompt section and the audit read one file: two copies of a font
 * size drift apart the first time the rule changes, and then the guidance says
 * one thing while the verdict says another. `loadSpec()` hot-reloads on mtime,
 * so editing the JSON changes both halves without a restart.
 *
 * Lifecycle: every side effect belongs to this plugin's fiber via `ctx.effect`,
 * so unmounting the row removes the prompt section and every tool.
 */

import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { checklistLabel, lineSpacingText, loadSpec, numberingScheme, pageNumberParts, RULES_PATH } from './spec.js';
import { openDocx } from './docx.js';
import { audit } from './audit.js';
import { applyFixes, applyHeadings, planHeadings, saveWithBackup } from './fix.js';
import { registerSettingsRoute } from './settings-route.js';

export const name = 'word-writing-spec';

/** `systemPrompt` and `tools` are hard dependencies: without them there is
 * nothing to inject into and nothing to register. */
export const inject = ['systemPrompt', 'tools'];

const HERE = path.dirname(fileURLToPath(import.meta.url));
export const SPEC_DIR = path.join(HERE, '..');
export const SECTION_NAME = 'word-writing-spec';
/** Late in the assembly: this is reference material, not identity or policy. */
const SECTION_ORDER = 900;
/** Same-origin API the settings page (lib/client.js) talks to. */
export const API_PREFIX = '/word-writing-spec/api';

/* ------------------------------------------------------------------ */
/* Prompt section                                                      */
/* ------------------------------------------------------------------ */

const pct = (v) => `${v}pt`;

const ALIGN_CN = { left: '左对齐', center: '居中', right: '右对齐', both: '两端对齐' };

const NUMERAL_CN = {
  decimal: '阿拉伯数字 1, 2, 3',
  lowerRoman: '罗马小写 i, ii, iii',
  upperRoman: '罗马大写 I, II, III',
  lowerLetter: '字母小写 a, b, c',
  upperLetter: '字母大写 A, B, C',
  roman: '罗马数字',
};

const cm = (v) => `${v}cm`;

/**
 * Render the rule table as prompt text.
 *
 * Deliberately flat and table-like: the reader is a model that must be able to
 * pull one value ("how many spaces after 附录 A") without parsing prose. Every
 * number here is read from the JSON, never typed in.
 */
export function renderSpecSection(spec) {
  const p = spec.page;
  const b = spec.body;
  const h = spec.headings;
  const t = spec.table;
  const f = spec.figure;
  const c = spec.code;
  const hf = spec.headerFooter;
  const sb = spec.sectionBreak;
  const r = spec.reference;
  const ai = spec.aiSupplement;
  const cover = spec.cover;
  const toc = spec.toc;

  const lines = [];
  const push = (...xs) => lines.push(...xs);
  const blank = () => lines.push('');

  push(`## ${spec.meta.title}${spec.meta.subtitle ?? ''}`);
  blank();
  push(
    '本会话产出的任何 .docx 都必须遵守下列规则。它们与 word_spec_audit / word_spec_fix 检查的是',
    '同一张表，按规则写出来的文档无需事后修复即可通过审查。规则未覆盖的细节，优先沿用文档中',
    '已有的一致写法，不要另创一套。',
  );
  blank();

  push('### 页面设置');
  push(`- 纸张 ${p.size}：${cm(p.widthCm)} × ${cm(p.heightCm)}`);
  push(
    `- 页边距：上 ${cm(p.marginsCm.top)}、下 ${cm(p.marginsCm.bottom)}、左 ${cm(p.marginsCm.left)}、右 ${cm(p.marginsCm.right)}`,
  );
  push(`- 页眉距边界 ${cm(p.headerDistanceCm)}，页脚距边界 ${cm(p.footerDistanceCm)}；装订线 ${cm(p.gutterCm)}`);
  blank();

  push('### 正文');
  push(`- 中文 ${b.fontCN}，英文与数字 ${b.fontEN}，${pct(b.sizePt)}（${b.sizeNameCN}）`);
  push(`- 行距${lineSpacingText(b.lineSpacing)}；首行缩进 ${b.firstLineIndentChars} 字符；对齐两端对齐`);
  push(`- 段前段后 ${pct(b.spaceBeforePt)}，不加额外空行来制造间距`);
  blank();

  push('### 标题层级');
  push(`- 最多 ${h.maxLevel} 级，更深层级用 ${h.deeperLevelRepresentation}`);
  push(`- 必须使用 Word 内置标题样式（Heading 1/2/3），不要用手工加粗加大的普通段落冒充标题`);
  if (h.numberingScheme) {
    push(`- 编号方案：${numberingScheme(h.numberingScheme).label}（${numberingScheme(h.numberingScheme).note}），不要混用其他编号形态`);
  }
  push(`- 编号与标题文字之间空 ${h.numberingGapSpaces} 个半角空格（编号自带「、」或「）」时不留空格）`);
  for (const lv of h.levels) {
    const bits = [
      `${lv.level} 级：${lv.fontCN} ${pct(lv.sizePt)}（${lv.sizeNameCN}）${lv.bold ? ' 加粗' : ' 不加粗'}`,
      `编号形如 ${lv.numberingExample}`,
      `${lv.alignment === 'left' ? '左对齐' : lv.alignment}`,
      `段前 ${pct(lv.spaceBeforePt)} 段后 ${pct(lv.spaceAfterPt)}`,
    ];
    if (lv.pageBreakBefore) bits.push('另起新页');
    push(`  - ${bits.join('，')}`);
  }
  blank();

  push('### 表格');
  push(`- 表题在表格上方，形如「${t.captionFormat}」（${t.numberingScheme.replace('{chapter}', '章节号').replace('{seq}', '序号')}），${t.captionFontCN} ${pct(t.captionSizePt)}，${t.captionAlignment === 'left' ? '左对齐' : t.captionAlignment}`);
  push(`- 表内文字 ${t.bodyFontCN}/${t.bodyFontEN} ${pct(t.bodySizePt)}；表头行 ${t.headerRowFontCN}${t.headerRowBold ? ' 加粗' : ''}`);
  push(`- 表格样式 ${t.style}；跨页时重复表头行`);
  push(`- 正文引用写「${t.referenceInText}」`);
  blank();

  push('### 插图');
  const figureSide = f.captionPosition === 'above' ? '上方' : '下方';
  push(`- 图题在图片${figureSide}，形如「${f.numberPattern.replace(/^\^/, '').replace(/\$$/, '')}」，${f.captionFontCN} ${pct(f.captionSizePt)}，${ALIGN_CN[f.captionAlignment] ?? f.captionAlignment}`);
  push(`- 图片段落行距不得使用固定值（会裁切图片）；用 ${f.paragraphLineSpacingAllowed.join(' 或 ')}`);
  push(`- 图片段落${ALIGN_CN[f.paragraphAlignment] ?? f.paragraphAlignment}，段前段后各 ${pct(f.paragraphSpaceBeforePt)}；分辨率不低于 ${f.minDpi} dpi（建议 ${f.recommendedDpiRange.join('–')}）`);
  push(`- 嵌入方式 ${f.embedding === 'inline' ? '嵌入式（inline）' : '浮动式（floating）'}；正文引用写「${f.referenceInText}」`);
  blank();

  push('### 代码块');
  push(`- ${c.font} ${c.sizePtRange[0]}–${c.sizePtRange[1]}pt，左缩进 ${cm(c.indentLeftCm)}，不设首行缩进`);
  push(`- 行距${lineSpacingText(c.lineSpacing)}；行内代码同样用 ${c.inlineCodeFont}`);
  push(`- 注释解释为什么而不是做了什么；标识符用${c.stNamingConvention}`);
  blank();

  push('### 页眉页脚');
  push(`- 页眉内容「${hf.headerContent}」，${hf.headerFontCN}/${hf.headerFontEN} ${pct(hf.headerSizePt)}（${hf.headerSizeNameCN}），灰色 #${hf.headerColor}，居中`);
  push(
    `- 页码位于${ALIGN_CN[hf.pageNumberAlignment] ?? '居中'}的${hf.pageNumberPosition === 'header' ? '页眉' : '页脚'}，样式「${pageNumberParts(hf.pageNumberFormat).format}」，数字用${NUMERAL_CN[hf.pageNumberNumeral ?? 'decimal'] ?? '阿拉伯数字'}；封面不编页码`,
  );
  push(
    `- 摘要与目录用${NUMERAL_CN[hf.frontMatterNumeral] ?? hf.frontMatterNumeral}单独编页，正文从第 ${hf.bodyStartNumber} 页重新起算`,
  );
  if (hf.hideHeaderOnCoverAndAbstract) push('- 封面与摘要页不显示页眉');
  blank();

  push('### 封面与目录');
  if (cover) {
    const coverTitle = String(cover.title ?? '').trim() || spec.meta?.title || '（未设置）';
    push(`- 封面标题「${coverTitle}」，${cover.titleFontCN} ${pct(cover.titleSizePt)}${cover.titleBold ? ' 加粗' : ''}，${ALIGN_CN[cover.alignment] ?? cover.alignment}`);
    if (cover.subtitle || spec.meta?.subtitle) push(`- 封面副标题「${String(cover.subtitle ?? '').trim() || spec.meta.subtitle}」`);
    if (cover.showProject && spec.meta?.project) push(`- 封面显示系统/项目名「${spec.meta.project}」`);
    if (cover.showDate) push('- 封面显示日期');
  }
  if (toc) {
    if (toc.enabled) {
      push(`- 目录标题「${toc.title}」，${toc.titleFontCN} ${pct(toc.titleSizePt)}${toc.titleBold ? ' 加粗' : ''}；最多 ${toc.maxLevel} 级${toc.showPageNumbers ? '，含页码' : '，不含页码'}`);
      push('- 目录必须是 TOC 域（打开 Word 后按 F9 或右键「更新域」生成），不要手打条目');
    } else {
      push('- 本文档不生成目录');
    }
  }
  blank();

  push('### 分节与分页');
  push(
    sb.newPagePerLevel1Heading
      ? `- 每个一级标题另起新页（用 ${sb.method}，禁止用${sb.forbiddenMethod}）`
      : '- 一级标题不强制另起新页',
  );
  push(sb.level2And3NewPage ? '- 二三级标题也另起新页' : '- 二三级标题不另起新页');
  blank();

  push('### 参考文献');
  push(`- 著录标准 ${r.listStandard}；正文引用用${r.inTextCitation}，如 ${r.inTextPatterns.join('、')}`);
  push(`- 字体 ${r.fontCN}/${r.fontEN} ${pct(r.sizePt)}`);
  for (const [k, v] of Object.entries(r.templates ?? {})) push(`  - [${k}] ${v}`);
  blank();

  push('### AI 辅助开发补充约定');
  if (ai.backupRequiredBeforeEdit) {
    push(`- 修改任何已有文档前先备份：${ai.backupLocation}生成「${ai.backupNamePattern}」`);
  }
  push(`- 生成的文件按「${ai.fileNamingPattern}」命名，例如 ${ai.fileNamingExample}`);
  if (ai.fileNamingForbidPureChinese) push('- 文件名不要只用中文，需含日期或英文标识');
  if (Array.isArray(ai.requiredDirs) && ai.requiredDirs.length) {
    push(`- 交付目录至少包含：${ai.requiredDirs.join('、')}`);
  }
  const norm = ai.headingNormalize;
  if (norm) {
    push('- 把「看起来像标题但不是标题样式」的段落改成标题样式时，必须按这个顺序：');
    push(`  ${norm.steps.join(' → ')}`);
    if (norm.mustShowCandidatesFirst) push('  先展示候选列表等用户确认，不要直接改；不确定的标为待确认');
  }
  blank();

  push('### 生成前自检');
  push('提交文档前逐项核对（工具 word_spec_audit 会用同一张表复核）：');
  for (const item of spec.checklist ?? []) push(`- [ ] ${checklistLabel(item, spec)}`);

  return lines.join('\n');
}

/* ------------------------------------------------------------------ */
/* Tool helpers                                                        */
/* ------------------------------------------------------------------ */

/** Compact, model-readable audit rendering: verdict first, then each finding. */
export function renderAudit(report, { maxFindings = 40 } = {}) {
  const out = [];
  const s = report.summary;
  out.push(
    `verdict: ${s.verdict} — error ${s.error} / warning ${s.warning} / info ${s.info}（passed ${s.passed}，warned ${s.warned}，failed ${s.failed}）`,
  );
  const st = report.stats;
  out.push(
    `document: ${st.paragraphs} 段落（非空 ${st.nonEmptyParagraphs ?? '?'}）/ ${st.tables} 表格 / ${st.figures ?? 0} 插图 / ${st.sections} 节 / ${st.styles} 样式 / 正文 ${st.bodySizePt}pt`,
  );
  const shown = report.findings.slice(0, maxFindings);
  if (shown.length) {
    out.push('');
    for (const f of shown) {
      out.push(`[${f.severity}] ${f.rule} :: ${f.title}`);
      out.push(`    应: ${f.expected}`);
      out.push(`    实: ${f.actual}`);
      if (f.fixable) out.push('    可自动修复');
    }
    if (report.findings.length > shown.length) {
      out.push(`… 另有 ${report.findings.length - shown.length} 条未列出`);
    }
  }
  if (report.errors?.length) {
    out.push('', '检查器异常（不是文档问题，请报告）：');
    for (const e of report.errors) out.push(`  ${e.group}: ${e.message}`);
  }
  return out.join('\n');
}

/** Resolve a user-supplied path against the session cwd. */
function resolveInput(file, exec) {
  if (!file || typeof file !== 'string') throw new Error('file is required');
  const cwd = exec?.cwd ?? process.cwd();
  return path.isAbsolute(file) ? file : path.resolve(cwd, file);
}

const OUTPUT = {
  type: 'object',
  properties: {
    verdict: { type: 'string' },
    error: { type: 'integer' },
    warning: { type: 'integer' },
    info: { type: 'integer' },
    text: { type: 'string' },
  },
  required: ['verdict', 'text'],
  additionalProperties: false,
};

const renderOutput = (_args, value) => [{ type: 'text', text: value.text }];

/* ------------------------------------------------------------------ */
/* Plugin                                                              */
/* ------------------------------------------------------------------ */

export function apply(ctx) {
  const systemPrompt = ctx.systemPrompt;

  /** Render the section text; the returned disposer belongs to the fiber. */
  const sectionText = () => {
    try {
      return renderSpecSection(loadSpec());
    } catch (err) {
      // A broken rule table must not take the whole prompt down with it.
      console.error('[word-writing-spec] failed to render spec section:', err);
      return '';
    }
  };

  ctx.effect(
    () =>
      systemPrompt.section({
        name: SECTION_NAME,
        order: SECTION_ORDER,
        // A thunk, not a string: the host pulls it on every prompt assembly, so
        // `loadSpec()`'s mtime check runs each turn and an edited rule table
        // reaches the next prompt without a restart. Passing the rendered string
        // would freeze the guidance at apply() while the tools kept consulting
        // the live file — the exact drift this plugin exists to avoid.
        text: sectionText,
      }),
    'word-writing-spec: rule table section',
  );

  const register = (definition) => ctx.effect(() => ctx.tools.register(definition), `word-writing-spec: ${definition.name}`);

  /* -- audit ------------------------------------------------------- */
  register({
    name: 'word_spec_audit',
    description:
      '按《AI 辅助开发文档格式规范》审查一个 .docx：页面设置、正文字体字号行距、标题层级与编号、表格/插图/代码块格式、页眉页脚页码、参考文献、附录、目录污染、文件命名。只读，不修改文件。返回 verdict 与逐条 finding（含 expected/actual）。',
    parameters: {
      type: 'object',
      properties: {
        file: { type: 'string', description: '目标 .docx 路径；相对路径按会话工作目录解析' },
        only: {
          type: 'array',
          items: { type: 'string' },
          description: '只检查这些规则前缀，如 ["page","body","table"]；省略则全查',
        },
      },
      required: ['file'],
      additionalProperties: false,
    },
    output: { schema: OUTPUT, render: renderOutput },
    execute: async (args, exec) => {
      try {
        const file = resolveInput(args.file, exec);
        const spec = loadSpec();
        const docx = openDocx(file);
        const report = audit(docx, spec, { only: args.only });
        return {
          verdict: report.summary.verdict,
          error: report.summary.error,
          warning: report.summary.warning,
          info: report.summary.info,
          text: renderAudit(report),
        };
      } catch (err) {
        return { verdict: 'error', text: `审查失败：${err.message}` };
      }
    },
  });

  /* -- plan heading normalization (§8.6 two-step) ------------------ */
  register({
    name: 'word_spec_plan_headings',
    description:
      '扫描「看起来像标题但不是标题样式」的段落，返回候选列表（段落号、文本、当前样式、推断级别、置信度、证据）。只读。按 §8.6 要求，转换前必须先把这份列表给用户确认。',
    parameters: {
      type: 'object',
      properties: { file: { type: 'string', description: '目标 .docx 路径' } },
      required: ['file'],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          text: { type: 'string' },
          candidates: { type: 'array', items: { type: 'object' } },
        },
        required: ['text'],
        additionalProperties: false,
      },
      render: renderOutput,
    },
    execute: async (args, exec) => {
      try {
        const file = resolveInput(args.file, exec);
        const spec = loadSpec();
        const docx = openDocx(file);
        const { candidates, bodySizePt } = planHeadings(docx, spec);
        const lines = [`正文基准字号 ${bodySizePt}pt，发现 ${candidates.length} 个候选段落：`];
        for (const c of candidates) {
          lines.push(
            `- #${c.paragraph} 「${c.text}」 当前样式 ${c.currentStyle} → 建议 ${c.suggestedLevel ?? '?'} 级` +
              `（${c.confidence}${c.requiresConfirmation ? '，待确认' : ''}；${c.reason}）`,
          );
        }
        if (!candidates.length) lines.push('（无需转换）');
        return { text: lines.join('\n'), candidates };
      } catch (err) {
        return { text: `扫描失败：${err.message}` };
      }
    },
  });

  /* -- fix --------------------------------------------------------- */
  register({
    name: 'word_spec_fix',
    description:
      '按规范修复 .docx 并另存为新文件：页面设置、正文格式、标题格式与样式定义、一级标题分页、附录编号空格、表格、插图段落、代码块、页眉页脚。修改前自动备份原文件（同目录 _backup_日期.docx）。可选 acceptedHeadings 传入用户已确认的标题候选（来自 word_spec_plan_headings）。',
    parameters: {
      type: 'object',
      properties: {
        file: { type: 'string', description: '源 .docx 路径；不会被覆盖' },
        output: { type: 'string', description: '输出 .docx 路径；省略则写同目录 <名>_fixed.docx' },
        rules: {
          type: 'array',
          items: { type: 'string' },
          description: '只修这些规则前缀，如 ["page","body"]；省略则全修',
        },
        skip: { type: 'array', items: { type: 'string' }, description: '跳过这些规则前缀' },
        dryRun: { type: 'boolean', description: '只报告将会修什么，不写文件' },
        acceptedHeadings: {
          type: 'array',
          items: { type: 'object' },
          description: '已确认的标题候选行（word_spec_plan_headings 返回的 candidates 子集）',
        },
      },
      required: ['file'],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          text: { type: 'string' },
          output: { type: 'string' },
          backup: { type: 'string' },
          applied: { type: 'array', items: { type: 'object' } },
        },
        required: ['text'],
        additionalProperties: false,
      },
      render: renderOutput,
    },
    execute: async (args, exec) => {
      try {
        const file = resolveInput(args.file, exec);
        const spec = loadSpec();
        const docx = openDocx(file);

        // Dry run: report the fixers that would run without touching the file.
        if (args.dryRun) {
          const { applied, skipped } = applyFixes(docx, spec, { rules: args.rules, skip: args.skip });
          const lines = ['dry run，未写入文件：'];
          for (const a of applied) lines.push(`  将修：${a.rule} — ${a.label}（${a.count} 处）`);
          for (const s of skipped) lines.push(`  跳过：${s.rule} — ${s.reason}`);
          return { text: lines.join('\n'), applied };
        }

        const outPath =
          args.output != null
            ? resolveInput(args.output, exec)
            : path.join(path.dirname(file), `${path.basename(file, path.extname(file))}_fixed.docx`);

        const headingRows = Array.isArray(args.acceptedHeadings) ? args.acceptedHeadings : [];
        const applied = [];
        if (headingRows.length) {
          const rows = applyHeadings(docx, spec, headingRows);
          applied.push({ rule: 'headings.convert', label: '人工确认的标题转换', count: rows.length });
        }
        const fixResult = applyFixes(docx, spec, { rules: args.rules, skip: args.skip });
        applied.push(...fixResult.applied);

        const saved = saveWithBackup(docx, outPath, { date: new Date() });

        const lines = [`已保存：${saved.output}（${saved.bytes} 字节）`, `备份：${saved.backup}`];
        if (!saved.verified) lines.push('⚠ 写回校验未通过，请重新检查该文件');
        lines.push('', '已修复：');
        for (const a of applied) lines.push(`- ${a.rule} — ${a.label}（${a.count} 处）`);
        for (const s of fixResult.skipped) lines.push(`- 跳过 ${s.rule}：${s.reason}`);
        if (!applied.length) lines.push('（没有需要修的项目）');
        lines.push('', '改完请重新跑 word_spec_audit 复核；文档中的目录、页码等域需要在 Word 里更新（Ctrl+A 后 F9）。');
        return { text: lines.join('\n'), output: saved.output, backup: saved.backup, applied };
      } catch (err) {
        return { text: `修复失败：${err.message}` };
      }
    },
  });

  /* -- rules ------------------------------------------------------- */
  register({
    name: 'word_spec_rules',
    description:
      '读取规范规则表本身（唯一规则来源）。可按 section 只取一部分，如 page / body / headings / table / figure / code / headerFooter / reference。写文档前用它确认具体数值，不要凭记忆。',
    parameters: {
      type: 'object',
      properties: {
        section: { type: 'string', description: '只返回该顶层小节；省略返回全部' },
      },
      additionalProperties: false,
    },
    output: {
      schema: { type: 'object', properties: { text: { type: 'string' } }, required: ['text'], additionalProperties: false },
      render: renderOutput,
    },
    execute: async (args) => {
      try {
        const spec = loadSpec();
        const picked = args.section ? spec[args.section] : spec;
        if (picked === undefined) {
          return { text: `没有 section「${args.section}」；可用：${Object.keys(spec).join(', ')}` };
        }
        return { text: `规则来源：${RULES_PATH}\n\n${JSON.stringify(picked, null, 2)}` };
      } catch (err) {
        return { text: `读取规则失败：${err.message}` };
      }
    },
  });

  /* -- settings page API ------------------------------------------- */
  // Not a hard dependency: headless/tui sessions have no web server, and making
  // `webServer` part of `inject` would leave this whole row pending — taking the
  // prompt section and all four tools down with a page nobody can see anyway.
  registerSettingsRoute(ctx, { apiPrefix: API_PREFIX });
}
