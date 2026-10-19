/**
 * dsh-word-writing-spec — Client half (settings page).
 *
 * Hand-authored in the browser module-loader format the web shell expects
 * (`window.__ModuleLoader__.load` + `require`), so this package needs no build
 * step. Shipped as `exports['./client']` and declared under `dsh.client` in
 * package.json; the host ships it to the browser at startup.
 *
 * What the page is for: the rule table is the single source of truth for both
 * the generation constraints and the audit verdicts, and until now the only way
 * to change it was to hand-edit a JSON file inside the package. This page edits
 * that same file through the host's settings route (`/word-writing-spec/api`),
 * which validates against the schema, backs up the previous table, and writes
 * atomically.
 *
 * Deliberate: the page edits VALUES, not structure. It never adds or removes
 * keys (the form is generated from a fixed field list, and the raw-JSON tab is
 * validated against the same schema). A free-form editor that could silently
 * drop a required key would break the audit with no visible cause.
 */
window.__ModuleLoader__.load({
  id: 'dsh-word-writing-spec',
  factory: (require) => {
    var module = { exports: {} };
    var exports = module.exports;

    const React = require('react');
    const h = React.createElement;

    const API = '/word-writing-spec/api';
    const NS = 'word-writing-spec';

    /* ---------------------------------------------------------------- */
    /* Styles                                                            */
    /* ---------------------------------------------------------------- */

    const CSS = `
.wws-root{display:flex;flex-direction:column;gap:12px;padding:4px 2px 48px;max-width:1000px;font-size:13px;line-height:1.6;color:var(--dsw-alias-label-primary,#1f2329)}
.wws-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap}
.wws-title{font-size:16px;font-weight:600;margin:0}
.wws-sub{font-size:12px;color:var(--dsw-alias-label-secondary,#8a9099);margin:2px 0 0}
.wws-meta{font-size:11px;color:var(--dsw-alias-label-secondary,#8a9099);font-family:ui-monospace,Menlo,Consolas,monospace;word-break:break-all}
.wws-layout{display:grid;grid-template-columns:130px minmax(0,1fr);gap:16px;align-items:start}
.wws-nav{position:sticky;top:0;display:flex;flex-direction:column;gap:2px;padding:4px 0}
.wws-nav-item{font:inherit;font-size:12px;text-align:left;padding:5px 9px;border-radius:6px;border:1px solid transparent;background:transparent;color:var(--dsw-alias-label-secondary,#8a9099);cursor:pointer;display:flex;align-items:center;gap:6px}
.wws-nav-item:hover{background:var(--dsw-alias-bg-layer-2,rgba(0,0,0,.05))}
.wws-nav-item.wws-nav-on{background:var(--dsw-alias-bg-layer-2,rgba(0,0,0,.06));border-color:var(--dsw-alias-border-l1,rgba(0,0,0,.1));color:var(--dsw-alias-brand-primary,#3b6cff);font-weight:600}
.wws-nav-item .wws-dot{width:5px;height:5px;border-radius:50%;background:var(--dsw-alias-state-warn-primary,#d97706);flex:none}
.wws-body{display:flex;flex-direction:column;gap:12px;min-width:0}
.wws-card{border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.1));border-radius:10px;background:var(--dsw-alias-bg-layer-1,#fff);padding:12px 14px;display:flex;flex-direction:column;gap:8px;scroll-margin-top:8px}
.wws-card-head{display:flex;align-items:center;gap:8px;min-height:24px;flex-wrap:wrap}
.wws-hint{font-size:10px;color:var(--dsw-alias-label-secondary,#8a9099);font-family:ui-monospace,Menlo,Consolas,monospace;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.wws-hint.wws-hint-bad{color:var(--dsw-alias-state-error-primary,#d94838)}
.wws-inline{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--dsw-alias-label-secondary,#8a9099);cursor:pointer;white-space:nowrap}
.wws-inline .wws-select{width:auto;min-width:64px;height:24px;padding:1px 5px;font-size:12px}
.wws-inline.wws-dirty .wws-inline-label{color:var(--dsw-alias-brand-primary,#3b6cff);font-weight:600}
.wws-format{display:flex;flex-direction:column;gap:4px;min-width:0}
.wws-format .wws-select,.wws-format .wws-input{width:100%}
.wws-card h3{margin:0;font-size:13px;font-weight:600;display:flex;align-items:center;gap:8px}
.wws-card h3 .wws-count{font-weight:400;font-size:11px;color:var(--dsw-alias-label-secondary,#8a9099)}
.wws-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(132px,1fr));gap:8px 12px}
.wws-field{display:flex;flex-direction:column;gap:3px;min-width:0}
.wws-field label{font-size:11px;color:var(--dsw-alias-label-secondary,#8a9099);display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.wws-field.wws-dirty label{color:var(--dsw-alias-brand-primary,#3b6cff);font-weight:600}
.wws-input,.wws-select{width:100%;box-sizing:border-box;font:inherit;font-size:12px;height:28px;padding:3px 7px;border-radius:6px;border:1px solid var(--dsw-alias-border-l2,rgba(0,0,0,.18));background:var(--dsw-alias-bg-base,#fff);color:inherit}
.wws-input:focus,.wws-select:focus{outline:2px solid var(--dsw-alias-brand-primary,#3b6cff);outline-offset:-1px}
.wws-check{display:flex;align-items:center;gap:6px;font-size:12px;height:28px}
.wws-color{display:flex;align-items:center;gap:6px}
.wws-color-swatch{flex:none;width:28px;height:28px;padding:0;border:1px solid var(--dsw-alias-border-l2,rgba(0,0,0,.18));border-radius:6px;background:transparent;cursor:pointer}
.wws-color-hex{flex:1;min-width:0;font-family:ui-monospace,Menlo,Consolas,monospace;text-transform:lowercase}
.wws-check input{width:14px;height:14px;margin:0}
.wws-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.wws-btn{font:inherit;font-size:12px;padding:5px 12px;border-radius:7px;border:1px solid var(--dsw-alias-border-l2,rgba(0,0,0,.18));background:var(--dsw-alias-bg-layer-1,#fff);color:inherit;cursor:pointer}
.wws-btn:hover{background:var(--dsw-alias-bg-layer-2,rgba(0,0,0,.05))}
.wws-btn:disabled{opacity:.5;cursor:not-allowed}
.wws-btn-primary{background:var(--dsw-alias-brand-primary,#3b6cff);border-color:transparent;color:#fff}
.wws-btn-primary:hover{filter:brightness(1.08);background:var(--dsw-alias-brand-primary,#3b6cff)}
.wws-btn-sm{padding:3px 9px;font-size:11px}
.wws-spacer{flex:1}
.wws-note{font-size:12px;color:var(--dsw-alias-label-secondary,#8a9099)}
.wws-err{color:var(--dsw-alias-state-error-primary,#d94838);font-size:12px;white-space:pre-wrap}
.wws-ok{color:var(--dsw-alias-state-success-primary,#18a058);font-size:12px}
.wws-warn{color:var(--dsw-alias-state-warn-primary,#d97706);font-size:12px;white-space:pre-wrap}
.wws-list{display:flex;flex-direction:column;gap:6px;max-height:260px;overflow:auto}
.wws-item{display:flex;align-items:center;gap:10px;padding:6px 8px;border-radius:7px;background:var(--dsw-alias-bg-layer-2,rgba(0,0,0,.05))}
.wws-item .wws-name{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:11px;flex:1;word-break:break-all}
.wws-pre{margin:0;padding:10px 12px;border-radius:8px;background:var(--dsw-alias-bg-layer-2,rgba(0,0,0,.05));font-family:ui-monospace,Menlo,Consolas,monospace;font-size:11px;line-height:1.5;white-space:pre-wrap;word-break:break-word;max-height:340px;overflow:auto}
.wws-textarea{width:100%;box-sizing:border-box;min-height:200px;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:11px;line-height:1.5;padding:8px 10px;border-radius:8px;border:1px solid var(--dsw-alias-border-l2,rgba(0,0,0,.18));background:var(--dsw-alias-bg-base,#fff);color:inherit;resize:vertical}
.wws-tabs{display:flex;gap:4px}
.wws-tab{font:inherit;font-size:12px;padding:4px 12px;border-radius:7px 7px 0 0;border:1px solid transparent;border-bottom:none;background:transparent;color:var(--dsw-alias-label-secondary,#8a9099);cursor:pointer}
.wws-tab.wws-tab-on{background:var(--dsw-alias-bg-layer-1,#fff);border-color:var(--dsw-alias-border-l1,rgba(0,0,0,.1));color:var(--dsw-alias-label-primary,#1f2329);font-weight:600}
.wws-empty{padding:24px;text-align:center;color:var(--dsw-alias-label-secondary,#8a9099);font-size:12px}
`;

    /* ---------------------------------------------------------------- */
    /* Host transport                                                    */
    /* ---------------------------------------------------------------- */

    /**
     * POST to the plugin's own route.
     *
     * 15s abort: the route only touches one small file, so a slow answer means
     * something is wedged and the page should say so rather than spin forever.
     */
    async function callHost(method, args) {
      const controller = typeof AbortController === 'function' ? new AbortController() : null;
      const timer = controller !== null ? setTimeout(() => controller.abort(), 15000) : null;
      let response;
      try {
        response = await fetch(API, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ method, args: args === undefined ? null : args }),
          signal: controller !== null ? controller.signal : undefined,
        });
      } catch (err) {
        if (controller !== null && controller.signal.aborted) throw new Error('请求宿主超时（15 秒）');
        throw new Error('无法连接宿主：' + (err && err.message ? err.message : String(err)));
      } finally {
        if (timer !== null) clearTimeout(timer);
      }
      if (!response.ok) {
        // The route answers errors as 200 + ok:false, so a non-2xx here is a
        // transport-level refusal (403 origin guard, 415, 413). Try to surface
        // its message instead of a bare status.
        let detail = '';
        try {
          const body = await response.json();
          detail = body && body.message ? '：' + body.message : '';
        } catch {
          detail = '';
        }
        throw new Error(`宿主拒绝请求（HTTP ${response.status}）${detail}`);
      }
      return await response.json();
    }

    /* ---------------------------------------------------------------- */
    /* Path helpers                                                      */
    /* ---------------------------------------------------------------- */

    const getPath = (obj, dotted) => {
      let cur = obj;
      for (const seg of dotted.split('.')) {
        if (cur == null) return undefined;
        cur = cur[seg];
      }
      return cur;
    };

    const clone = (value) => JSON.parse(JSON.stringify(value));

    const sameValue = (a, b) => {
      if (a === b) return true;
      // A number typed into a text input arrives as a string; compare loosely so
      // "12" and 12 are not reported as a change.
      if (typeof a === 'number' && typeof b === 'string' && b.trim() !== '') return a === Number(b);
      if (typeof b === 'number' && typeof a === 'string' && a.trim() !== '') return b === Number(a);
      return false;
    };

    /* ---------------------------------------------------------------- */
    /* Field catalogue                                                   */
    /* ---------------------------------------------------------------- */

    /**
     * Which values the page exposes, and how to edit them.
     *
     * Curated rather than auto-generated from the schema: the schema also
     * carries regexes, file templates and prose, and a form built from it would
     * be mostly unusable text boxes. Everything not listed here stays editable
     * through the JSON tab, so the curated list is a shortcut, not a gate.
     */
    /**
     * Paper sizes the picker offers, in centimetres.
     *
     * The table lives here rather than in the rule JSON on purpose: the rule table
     * stores the *chosen* size plus its two dimensions, and the auditor compares
     * against those numbers. Keeping one authoritative copy of the table keeps the
     * picker from disagreeing with what actually gets written to the document.
     *
     * Adding a size means adding one row here and nothing else: the picker builds
     * its options from these keys, and the choice writes the two dimensions
     * alongside the name. Order below is display order.
     */
    const PAPER_SIZES = {
      A3: [29.7, 42],
      A4: [21, 29.7],
      A5: [14.8, 21],
      A6: [10.5, 14.8],
      B5: [17.6, 25],
      Letter: [21.59, 27.94],
    };

    /**
     * Common font families. The lists are suggestions, not a whitelist: a value
     * that is not listed (a document already using 方正小标宋简体, say) is shown
     * as an extra entry reading "<value>（当前值）" instead of being dropped, and
     * anything unusual can still be typed in the JSON tab.
     */
    const FONT_CN = [
      '宋体',
      '黑体',
      '楷体_GB2312',
      '仿宋_GB2312',
      '楷体',
      '仿宋',
      '微软雅黑',
      '等线',
      '方正小标宋简体',
      '华文中宋',
      '新宋体',
    ];
    const FONT_EN = [
      'Times New Roman',
      'Arial',
      'Calibri',
      'Cambria',
      'Georgia',
      'Verdana',
      'Segoe UI',
      'Tahoma',
    ];
    const FONT_MONO = ['Courier New', 'Consolas', 'Cascadia Mono', 'Lucida Console', 'Source Code Pro'];

    /**
     * Chinese type sizes. Picking a name writes BOTH the name and the point size,
     * because the rule table carries both and the auditor reads sizePt while the
     * prompt reads sizeNameCN — letting them disagree is how a document ends up
     * labelled 小四 but set in 12pt 三号.
     */
    const SIZE_NAMES = {
      初号: 42,
      小初: 36,
      一号: 26,
      小一: 24,
      二号: 22,
      小二: 18,
      三号: 16,
      小三: 15,
      四号: 14,
      小四: 12,
      五号: 10.5,
      小五: 9,
      六号: 7.5,
      小六: 6.5,
      七号: 5.5,
      八号: 5,
    };

    /**
     * Line-spacing modes, in the order Word shows them. Must stay in step with
     * LINE_SPACING_MODES in lib/spec.js — that copy decides what gets written to
     * the .docx and how the auditor compares it; this one only labels the picker.
     * `factor` is null for the modes whose number comes from the rule table.
     */
    const LINE_SPACING_MODES = {
      single: { label: '单倍行距', unit: 'lines', factor: 1 },
      oneAndHalf: { label: '1.5 倍行距', unit: 'lines', factor: 1.5 },
      double: { label: '2 倍行距', unit: 'lines', factor: 2 },
      multiple: { label: '多倍行距', unit: 'lines', factor: null },
      atLeast: { label: '最小值', unit: 'pt', factor: null },
      exact: { label: '固定值', unit: 'pt', factor: null },
    };

    /**
     * Heading numbering schemes, in the order the picker shows them. Must stay
     * in step with NUMBERING_SCHEMES in lib/spec.js — that copy decides which
     * label the auditor accepts and which examples the prompt quotes.
     */
    const NUMBERING_SCHEMES = {
      arabicDotted: {
        label: '1 / 1.1 / 1.1.1',
        note: '科技文档、报告',
        levels: [
          { example: '1', pattern: '^\\d+$' },
          { example: '1.1', pattern: '^\\d+\\.\\d+$' },
          { example: '1.1.1', pattern: '^\\d+\\.\\d+\\.\\d+$' },
        ],
      },
      chineseChapter: {
        label: '第一章 / 第一节 / 一、',
        note: '论文、书稿',
        levels: [
          { example: '第一章', pattern: '^第[一二三四五六七八九十百零〇\\d]+章$' },
          { example: '第一节', pattern: '^第[一二三四五六七八九十百零〇\\d]+节$' },
          { example: '一、', pattern: '^[一二三四五六七八九十百零〇]+、' },
        ],
      },
      chineseOrdinal: {
        label: '一、 / （一） / 1.',
        note: '公文、规范',
        levels: [
          { example: '一、', pattern: '^[一二三四五六七八九十百零〇]+、' },
          { example: '（一）', pattern: '^（[一二三四五六七八九十百零〇]+）' },
          { example: '1.', pattern: '^\\d+\\.(?!\\d)' },
        ],
      },
      arabicParen: {
        label: '1 / （1） / ①',
        note: '与更深层级表示法一致',
        levels: [
          { example: '1', pattern: '^\\d+$' },
          { example: '（1）', pattern: '^（\\d+）' },
          { example: '①', pattern: '^[①②③④⑤⑥⑦⑧⑨⑩]' },
        ],
      },
    };

    /**
     * Render the page-number format the way a reader will see it. A format with
     * no 「数字」 token is flagged rather than silently producing a footer with no
     * number: that is the one mistake this field can hide.
     */
    function pageNumberPreview(format) {
      const text = typeof format === 'string' && format !== '' ? format : '— 数字 —';
      if (!text.includes('数字')) return '缺少「数字」占位符，页码将无处放置';
      return '预览：' + text.split('数字').join('1');
    }

    /**
     * Ready-made page-number formats.
     *
     * The literal string is the label: it is exactly what lands in the footer, so
     * a paraphrase ("破折号包裹") would be a second spelling of the same fact and
     * would drift the first time someone edits a preset. Every entry carries one
     * 「数字」 token — a second one ("共 数字 页") needs a NUMPAGES field this
     * writer does not emit, and would print the word instead of a total.
     */
    const PAGE_NUMBER_PRESETS = [
      '— 数字 —',
      '第 数字 页',
      '- 数字 -',
      '(数字)',
      '数字',
      'Page 数字',
    ];

    const GROUPS = [
      {
        id: 'page',
        title: '页面设置',
        fields: [
          // `inline: true` lifts the control into the card header row instead of
          // the grid. A one-choice dropdown on its own grid row used to cost a
          // whole row of vertical space while saying nothing the title did not.
          // Changing it rewrites page.size, page.widthCm and page.heightCm
          // together, so the two dimension boxes below follow the choice.
          { path: 'page.size', label: '纸张大小', type: 'paper', inline: true },
          { path: 'page.widthCm', label: '纸张宽', type: 'number', step: 0.1 },
          { path: 'page.heightCm', label: '纸张高', type: 'number', step: 0.1 },
          { path: 'page.marginsCm.top', label: '上边距', type: 'number', step: 0.1 },
          { path: 'page.marginsCm.bottom', label: '下边距', type: 'number', step: 0.1 },
          { path: 'page.marginsCm.left', label: '左边距', type: 'number', step: 0.1 },
          { path: 'page.marginsCm.right', label: '右边距', type: 'number', step: 0.1 },
          { path: 'page.headerDistanceCm', label: '页眉距边界', type: 'number', step: 0.1 },
          { path: 'page.footerDistanceCm', label: '页脚距边界', type: 'number', step: 0.1 },
        ],
      },
      {
        id: 'body',
        title: '正文',
        fields: [
          { path: 'body.fontCN', label: '中文字体', type: 'font', options: FONT_CN },
          { path: 'body.fontEN', label: '西文字体', type: 'font', options: FONT_EN },
          { path: 'body.sizePt', label: '字号 (pt)', type: 'number', step: 0.5 },
          { path: 'body.sizeNameCN', label: '字号名称', type: 'sizeName' },
          { path: 'body.lineSpacing.valuePt', label: '行距数值', type: 'spacingValue', step: 0.5 },
          {
            path: 'body.lineSpacing.type',
            label: '行距方式',
            type: 'select',
            options: Object.keys(LINE_SPACING_MODES),
            labels: Object.fromEntries(Object.entries(LINE_SPACING_MODES).map(([k, v]) => [k, v.label])),
          },
          { path: 'body.color', label: '正文颜色 (hex)', type: 'color' },
          { path: 'body.firstLineIndentChars', label: '首行缩进 (字符)', type: 'number', step: 0.5 },
          {
            path: 'body.alignment',
            label: '对齐',
            type: 'select',
            options: ['both', 'left', 'center', 'right'],
          },
        ],
      },
      {
        id: 'headerFooter',
        title: '页眉页脚',
        fields: [
          { path: 'headerFooter.headerContent', label: '页眉内容', type: 'text' },
          { path: 'headerFooter.headerFontCN', label: '页眉中文字体', type: 'font', options: FONT_CN },
          { path: 'headerFooter.headerSizePt', label: '页眉字号 (pt)', type: 'number', step: 0.5 },
          { path: 'headerFooter.headerColor', label: '页眉颜色 (hex)', type: 'color' },
          {
            path: 'headerFooter.pageNumberNumeral',
            label: '页码数字样式',
            type: 'select',
            options: ['decimal', 'lowerRoman', 'upperRoman', 'lowerLetter', 'upperLetter'],
            labels: {
              decimal: '阿拉伯数字 1 2 3',
              lowerRoman: '罗马小写 i ii iii',
              upperRoman: '罗马大写 I II III',
              lowerLetter: '字母小写 a b c',
              upperLetter: '字母大写 A B C',
            },
          },
          {
            path: 'headerFooter.pageNumberPosition',
            label: '页码位置',
            type: 'select',
            options: ['footer', 'header'],
            labels: { footer: '页脚', header: '页眉' },
          },
          {
            path: 'headerFooter.pageNumberAlignment',
            label: '页码对齐',
            type: 'select',
            options: ['left', 'center', 'right'],
            labels: { left: '左对齐', center: '居中', right: '右对齐' },
          },
          {
            path: 'headerFooter.pageNumberFormat',
            label: '页码样式（数字为占位符）',
            // A dropdown of ready-made formats, plus 自定义 for anything else.
            // Typing "— 数字 —" by hand into an empty box was the only way
            // before, and the two most common mistakes (wrong dash, missing
            // placeholder) are exactly what presets remove.
            type: 'format',
            presets: PAGE_NUMBER_PRESETS,
            placeholder: '— 数字 —',
            preview: pageNumberPreview,
          },
          {
            path: 'headerFooter.frontMatterNumeral',
            label: '前置部分页码',
            type: 'select',
            options: ['roman', 'upperRoman', 'decimal'],
            labels: { roman: '罗马小写 i ii', upperRoman: '罗马大写 I II', decimal: '阿拉伯数字 1 2' },
          },
          { path: 'headerFooter.bodyStartNumber', label: '正文起始页码', type: 'number', step: 1 },
          { path: 'headerFooter.hideHeaderOnCoverAndAbstract', label: '封面/摘要不显示页眉', type: 'bool' },
          { path: 'headerFooter.coverNoPageNumber', label: '封面不编页码', type: 'bool' },
        ],
      },
      {
        id: 'cover',
        title: '封面',
        fields: [
          { path: 'cover.title', label: '封面标题（留空用 meta.title）', type: 'text' },
          { path: 'cover.subtitle', label: '封面副标题（留空用 meta.subtitle）', type: 'text' },
          { path: 'cover.titleFontCN', label: '标题字体', type: 'font', options: FONT_CN },
          { path: 'cover.titleSizePt', label: '标题字号 (pt)', type: 'number', step: 0.5 },
          { path: 'cover.titleBold', label: '标题加粗', type: 'bool' },
          { path: 'cover.color', label: '标题颜色 (hex)', type: 'color' },
          {
            path: 'cover.alignment',
            label: '标题对齐',
            type: 'select',
            options: ['center', 'left', 'right'],
            labels: { center: '居中', left: '左对齐', right: '右对齐' },
          },
          { path: 'cover.showProject', label: '显示项目名', type: 'bool' },
          { path: 'cover.showDate', label: '显示日期', type: 'bool' },
          { path: 'cover.spaceBeforePt', label: '标题上方留白 (pt)', type: 'number', step: 10 },
        ],
      },
      {
        id: 'toc',
        title: '目录',
        fields: [
          { path: 'toc.enabled', label: '生成目录', type: 'bool' },
          { path: 'toc.title', label: '目录标题', type: 'text' },
          { path: 'toc.maxLevel', label: '目录层级', type: 'number', step: 1 },
          { path: 'toc.titleFontCN', label: '目录标题字体', type: 'font', options: FONT_CN },
          { path: 'toc.titleSizePt', label: '目录标题字号 (pt)', type: 'number', step: 0.5 },
          { path: 'toc.titleBold', label: '目录标题加粗', type: 'bool' },
          { path: 'toc.titleColor', label: '目录标题颜色 (hex)', type: 'color' },
          { path: 'toc.entryFontCN', label: '条目字体', type: 'font', options: FONT_CN },
          { path: 'toc.entrySizePt', label: '条目字号 (pt)', type: 'number', step: 0.5 },
          { path: 'toc.entryColor', label: '条目颜色 (hex)', type: 'color' },
          { path: 'toc.showPageNumbers', label: '条目带页码', type: 'bool' },
        ],
      },
      {
        id: 'table',
        title: '表格',
        fields: [
          { path: 'table.captionFontCN', label: '表题字体', type: 'font', options: FONT_CN },
          { path: 'table.captionSizePt', label: '表题字号 (pt)', type: 'number', step: 0.5 },
          {
            path: 'table.captionPosition',
            label: '表题位置',
            type: 'select',
            options: ['above', 'below'],
          },
          {
            path: 'table.captionAlignment',
            label: '表题对齐',
            type: 'select',
            options: ['left', 'center', 'right'],
          },
          { path: 'table.bodyFontCN', label: '表内字体', type: 'font', options: FONT_CN },
          { path: 'table.bodySizePt', label: '表内字号 (pt)', type: 'number', step: 0.5 },
          { path: 'table.headerRowBold', label: '表头加粗', type: 'bool' },
          { path: 'table.repeatHeaderOnBreak', label: '跨页重复表头', type: 'bool' },
        ],
      },
      {
        id: 'figure',
        title: '插图',
        fields: [
          { path: 'figure.captionFontCN', label: '图题字体', type: 'font', options: FONT_CN },
          { path: 'figure.captionSizePt', label: '图题字号 (pt)', type: 'number', step: 0.5 },
          {
            path: 'figure.captionPosition',
            label: '图题位置',
            type: 'select',
            options: ['above', 'below'],
            labels: { above: '图片上方', below: '图片下方' },
          },
          {
            path: 'figure.captionAlignment',
            label: '图题对齐',
            type: 'select',
            options: ['left', 'center', 'right'],
            labels: { left: '左对齐', center: '居中', right: '右对齐' },
          },
          { path: 'figure.minDpi', label: '最低 dpi', type: 'number', step: 10, min: 72 },
          { path: 'figure.paragraphSpaceBeforePt', label: '图片段前 (pt)', type: 'number', step: 1 },
          { path: 'figure.paragraphSpaceAfterPt', label: '图片段后 (pt)', type: 'number', step: 1 },
          {
            path: 'figure.embedding',
            label: '嵌入方式',
            type: 'select',
            options: ['inline', 'floating'],
            labels: { inline: '嵌入式', floating: '浮动式' },
          },
        ],
      },
      {
        id: 'code',
        title: '代码块',
        fields: [
          { path: 'code.font', label: '等宽字体', type: 'font', options: FONT_MONO },
          { path: 'code.sizePtRange.0', label: '字号下限 (pt)', type: 'number', step: 0.5 },
          { path: 'code.sizePtRange.1', label: '字号上限 (pt)', type: 'number', step: 0.5 },
          { path: 'code.indentLeftCm', label: '左缩进 (cm)', type: 'number', step: 0.1 },
          { path: 'code.lineSpacing.valuePt', label: '代码行距数值', type: 'spacingValue', step: 0.5 },
          {
            path: 'code.lineSpacing.type',
            label: '代码行距方式',
            type: 'select',
            options: Object.keys(LINE_SPACING_MODES),
            labels: Object.fromEntries(Object.entries(LINE_SPACING_MODES).map(([k, v]) => [k, v.label])),
          },
        ],
      },
      {
        id: 'sectionBreak',
        title: '分节与分页',
        fields: [
          { path: 'sectionBreak.newPagePerLevel1Heading', label: '一级标题另起页', type: 'bool' },
          {
            path: 'sectionBreak.method',
            label: '分页方式',
            type: 'select',
            options: ['pageBreak', 'sectionBreak'],
          },
          { path: 'sectionBreak.level2And3NewPage', label: '二三级也另起页', type: 'bool' },
        ],
      },
      {
        id: 'meta',
        title: '文档标识',
        fields: [
          { path: 'meta.title', label: '规范名称', type: 'text' },
          { path: 'meta.subtitle', label: '副标题', type: 'text' },
        ],
      },
    ];

    /** Heading-level fields are per level, so they are built, not listed. */
    const levelFields = (index) => [
      { path: `headings.levels.${index}.fontCN`, label: `${index + 1} 级字体`, type: 'font', options: FONT_CN },
      { path: `headings.levels.${index}.sizePt`, label: `${index + 1} 级字号 (pt)`, type: 'number', step: 0.5 },
      { path: `headings.levels.${index}.bold`, label: `${index + 1} 级加粗`, type: 'bool' },
      { path: `headings.levels.${index}.pageBreakBefore`, label: `${index + 1} 级另起页`, type: 'bool' },
      { path: `headings.levels.${index}.spaceBeforePt`, label: `${index + 1} 级段前 (pt)`, type: 'number', step: 1 },
      { path: `headings.levels.${index}.spaceAfterPt`, label: `${index + 1} 级段后 (pt)`, type: 'number', step: 1 },
      { path: `headings.levels.${index}.color`, label: `${index + 1} 级颜色 (hex)`, type: 'color' },
    ];

    /**
     * Paths the numbering picker writes. They are not form fields (the picker
     * derives them), but a save has to carry them, or choosing a scheme would
     * update the page and leave the rule table saying "编号形如 1".
     */
    const schemePaths = (levelCount) => {
      const out = ['headings.numberingScheme'];
      for (let i = 0; i < levelCount; i++) {
        out.push(`headings.levels.${i}.numbering`, `headings.levels.${i}.numberingExample`);
      }
      return out;
    };

    /* ---------------------------------------------------------------- */
    /* Small components                                                  */
    /* ---------------------------------------------------------------- */

    function Field(props) {
      const { spec, path, label, type, options, labels, min, max, step, draft, onChange, inline, placeholder, preview, presets } = props;
      const value = getPath(draft, path);
      const dirty = !sameValue(value, getPath(spec, path));
      const id = 'wws-' + path.replace(/\./g, '-');

      // `font` and `sizeName` behave like selects but write differently, so they
      // are normalised to the same list shape here instead of duplicating the
      // "extra current value" logic three times.
      const isList = type === 'select' || type === 'paper' || type === 'font' || type === 'sizeName';
      const choices = type === 'paper'
        ? Object.keys(PAPER_SIZES)
        : type === 'sizeName'
          ? Object.keys(SIZE_NAMES)
          : options ?? [];

      const commit = (raw) => {
        if (type === 'bool') {
          onChange(path, raw);
          return;
        }
        if (type === 'paper') {
          // One choice, three writes: leaving the dimensions behind would make
          // the table self-contradictory (size says A5, dimensions still 21×29.7)
          // and the auditor would then flag a document the user just configured.
          const dims = PAPER_SIZES[raw];
          onChange('page.size', raw);
          if (dims !== undefined) {
            onChange('page.widthCm', dims[0]);
            onChange('page.heightCm', dims[1]);
          }
          return;
        }
        if (type === 'sizeName') {
          // Name and point size travel together: the prompt quotes the name while
          // the auditor measures sizePt, so writing only the name would leave the
          // document labelled 小四 and set in whatever the previous size was.
          onChange(path, raw);
          const pt = SIZE_NAMES[raw];
          if (pt !== undefined) {
            onChange('body.sizePt', pt);
          }
          return;
        }
        if (type === 'color') {
          // Both "#1f2329" and "1f2329" are accepted: the native picker emits
          // the hash while the rule table stores bare hex, and typing the hash
          // is the obvious mistake. Anything that is not six hex digits is
          // dropped rather than written — a malformed colour fails schema
          // validation and takes the whole save down with it.
          const hex = String(raw).trim().replace(/^#/, '');
          if (/^[0-9A-Fa-f]{6}$/.test(hex)) onChange(path, hex);
          return;
        }
        if (type === 'number' || type === 'spacingValue') {
          // An empty box means "not edited", not zero: writing 0 here would
          // silently zero a margin the user merely cleared.
          if (raw === '') return;
          const n = Number(raw);
          if (Number.isFinite(n)) onChange(path, n);
          return;
        }
        onChange(path, raw);
      };

      // The line-spacing number changes meaning with the mode: points for
      // 固定值/最小值, a multiple for 多倍行距, and nothing at all for the three
      // named multiples (their number is part of the mode). Reading the mode from
      // the draft — not from `spec` — makes the box follow the pending choice
      // before it is saved.
      const spacingBase = type === 'spacingValue' ? path.replace(/\.valuePt$/, '') : null;
      const spacingMode = spacingBase === null
        ? null
        : (LINE_SPACING_MODES[getPath(draft, spacingBase + '.type')] ?? LINE_SPACING_MODES.exact);
      const displayLabel = spacingMode === null
        ? label
        : spacingMode.unit === 'pt'
          ? label + '（磅）'
          : '行距（倍）';
      const fixedMultiple = spacingMode !== null && spacingMode.factor !== null;

      let control;
      if (type === 'bool') {
        control = h('input', {
          id,
          type: 'checkbox',
          checked: value === true,
          onChange: (e) => commit(e.target.checked),
        });
      } else if (isList) {
        // A value outside the list (hand-edited in the JSON tab, or a family this
        // picker does not offer) is shown as an extra read-only entry instead of
        // an empty box: silently displaying nothing would look like a bug, and
        // dropping it on save would silently rewrite the user's table.
        const current = value === undefined || value === null ? '' : String(value);
        const extra = current !== '' && !choices.includes(current) ? [current] : [];
        control = h(
          'select',
          { id, className: 'wws-select', value: current, onChange: (e) => commit(e.target.value) },
          extra
            .concat(choices)
            .map((opt) => {
              const text = extra.includes(opt) ? `${opt}（当前值）` : (labels?.[opt] ?? opt);
              return h('option', { key: opt, value: opt }, text);
            }),
        );
      } else if (type === 'format') {
        // A preset picker and a free-text box editing the same value. Picking a
        // preset writes it straight through; 自定义 hands control to the box so a
        // format nobody anticipated (「第 数字 页 共 数字 页」 and friends) stays
        // reachable. Both point at the same `path`, so no state can disagree.
        const current = value === undefined || value === null ? '' : String(value);
        const pick = current !== '' && presets.includes(current) ? current : '__custom__';
        control = h(
          'div',
          { className: 'wws-format' },
          h(
            'select',
            {
              id: id + '-preset',
              className: 'wws-select',
              value: pick,
              onChange: (e) => {
                if (e.target.value !== '__custom__') commit(e.target.value);
              },
            },
            presets
              .map((opt) => h('option', { key: opt, value: opt }, opt))
              .concat([h('option', { key: '__custom__', value: '__custom__' }, '自定义…')]),
          ),
          h('input', {
            id,
            className: 'wws-input',
            type: 'text',
            value: current,
            placeholder,
            spellCheck: false,
            onChange: (e) => commit(e.target.value),
          }),
        );
      } else if (type === 'color') {
        // A swatch alone is unusable for pasting a brand hex, and a hex box alone
        // is unusable for picking. Both edit the same value: the swatch needs a
        // leading "#" that the rule table does not store, hence the two shapes.
        const hex = value === undefined || value === null ? '' : String(value);
        control = h(
          'div',
          { className: 'wws-color' },
          h('input', {
            id,
            className: 'wws-color-swatch',
            type: 'color',
            value: /^[0-9A-Fa-f]{6}$/.test(hex) ? '#' + hex : '#000000',
            onChange: (e) => commit(e.target.value),
          }),
          h('input', {
            id: id + '-hex',
            className: 'wws-input wws-color-hex',
            type: 'text',
            value: hex,
            maxLength: 7,
            spellCheck: false,
            onChange: (e) => commit(e.target.value),
          }),
        );
      } else {
        // The named multiples show their own factor; every other mode shows the
        // stored number. This keeps the box truthful without writing the factor
        // into valuePt, which the other modes still need.
        const displayValue = fixedMultiple
          ? String(spacingMode.factor)
          : value === undefined || value === null ? '' : String(value);
        control = h('input', {
          id,
          className: 'wws-input',
          type: type === 'number' || type === 'spacingValue' ? 'number' : 'text',
          value: displayValue,
          min,
          max,
          step,
          placeholder,
          // 单倍/1.5 倍/2 倍 carry their number in the mode itself, so the box is
          // read-only rather than hidden: hiding it would shift the grid and make
          // the label look like it belongs to the next field.
          disabled: fixedMultiple ? true : undefined,
          onChange: (e) => commit(e.target.value),
        });
      }

      if (type === 'bool') {
        return h(
          'div',
          { className: 'wws-field' + (dirty ? ' wws-dirty' : ''), title: path },
          h('label', { htmlFor: id }, label),
          h('div', { className: 'wws-check' }, control, dirty ? h('span', { className: 'wws-note' }, '已修改') : null),
        );
      }
      // A field can explain itself beyond its label: the page-number format shows
      // what the reader will actually see, and complains when the counter has no
      // place to go. Rendered here (not inside the control branch) so it also
      // appears for inline fields.
      const hintText = typeof preview === 'function' ? preview(value) : null;
      const hint =
        hintText === null
          ? null
          : h(
              'span',
              { className: 'wws-hint' + (String(hintText).startsWith('缺少') ? ' wws-hint-bad' : '') },
              hintText,
            );
      if (inline === true) {
        return h(
          'label',
          { className: 'wws-inline' + (dirty ? ' wws-dirty' : ''), title: path, htmlFor: id },
          h('span', { className: 'wws-inline-label' }, label),
          control,
        );
      }
      return h(
        'div',
        { className: 'wws-field' + (dirty ? ' wws-dirty' : ''), title: path },
        h('label', { htmlFor: id }, displayLabel),
        control,
        hint,
      );
    }

    function Group(props) {
      const { id, title, fields, spec, draft, onChange, onReset, dirtyCount } = props;
      // Fields flagged inline sit in the header row; everything else keeps the
      // grid. Splitting here (rather than at render time) keeps the count in the
      // header honest about what the grid actually contains.
      const gridFields = fields.filter((f) => f.inline !== true);
      const inlineFields = fields.filter((f) => f.inline === true);
      return h(
        'section',
        { className: 'wws-card', id: 'wws-sec-' + id },
        h(
          'div',
          { className: 'wws-card-head' },
          h('h3', null, title, h('span', { className: 'wws-count' }, `${gridFields.length} 项`)),
          dirtyCount > 0 ? h('span', { className: 'wws-note' }, `${dirtyCount} 项已改`) : null,
          h('span', { className: 'wws-spacer' }),
          inlineFields.map((f) =>
            h(Field, {
              key: f.path,
              ...f,
              spec,
              draft,
              onChange,
            }),
          ),
          h(
            'button',
            {
              className: 'wws-btn wws-btn-sm',
              onClick: () => onReset(fields.map((f) => f.path)),
              disabled: dirtyCount === 0,
              title: '把本组字段还原成磁盘上的值',
            },
            '恢复本组默认',
          ),
        ),
        h(
          'div',
          { className: 'wws-grid' },
          gridFields.map((f) =>
            h(Field, {
              key: f.path,
              ...f,
              spec,
              draft,
              onChange,
            }),
          ),
        ),
      );
    }

    /* ---------------------------------------------------------------- */
    /* Panel                                                             */
    /* ---------------------------------------------------------------- */

    function Panel() {
      const [view, setView] = React.useState(null); // { spec, path, mtimeMs, bytes }
      const [draft, setDraft] = React.useState(null);
      const [error, setError] = React.useState('');
      const [notice, setNotice] = React.useState('');
      const [busy, setBusy] = React.useState(false);
      const [tab, setTab] = React.useState('form');
      const [jsonText, setJsonText] = React.useState('');
      const [preview, setPreview] = React.useState('');
      const [backups, setBackups] = React.useState(null);
      // Which nav entry is highlighted. Kept as state rather than derived from
      // scroll position: deriving it needs a scroll listener on an ancestor the
      // settings shell owns, which we do not control.
      const [activeSection, setActiveSection] = React.useState('page');
      const viewRef = React.useRef(null);

      /**
       * Scroll a section into view.
       *
       * scrollIntoView is optional in the harness the offline smoke test runs, and
       * a missing section must not throw: the nav is a convenience, and errors
       * here would surface as an unrelated「页面出错」.
       */
      const jumpTo = (id) => {
        setActiveSection(id);
        if (typeof document === 'undefined' || typeof document.getElementById !== 'function') return;
        const el = document.getElementById('wws-sec-' + id);
        if (el === null) return;
        if (typeof el.scrollIntoView === 'function') {
          el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      };

      const absorb = (result) => {
        if (!result || result.ok !== true) return;
        const next = { spec: result.spec, path: result.path, mtimeMs: result.mtimeMs, bytes: result.bytes };
        viewRef.current = next;
        setView(next);
        setDraft(clone(next.spec));
        setJsonText(JSON.stringify(next.spec, null, 2));
      };

      const load = React.useCallback(async () => {
        try {
          const result = await callHost('get');
          if (result.ok !== true) {
            setError(result.message || '读取规则表失败');
            return;
          }
          absorb(result);
          setError('');
        } catch (err) {
          setError(err.message);
        }
      }, []);

      React.useEffect(() => {
        load();
      }, [load]);

      // Live updates: another window (or the user editing the JSON on disk) must
      // show up here without a manual reload. EventSource is optional — without
      // it the page still works, just not automatically.
      const sseRef = React.useRef(null);
      React.useEffect(() => {
        if (typeof EventSource !== 'function') return undefined;
        const source = new EventSource(API);
        sseRef.current = source;
        source.onmessage = (ev) => {
          let pushed;
          try {
            pushed = JSON.parse(ev.data);
          } catch {
            return;
          }
          const current = viewRef.current;
          if (current !== null && pushed.mtimeMs === current.mtimeMs) return;
          // Do not clobber unsaved edits: only adopt a pushed table when the
          // form is clean, otherwise the user's typing would vanish.
          const dirtyNow = current === null ? false : hasChanges(current.spec, draft);
          if (dirtyNow) return;
          absorb({ ok: true, ...pushed });
        };
        return () => {
          sseRef.current = null;
          source.close();
        };
      }, [draft]);

      const setField = (path, value) => {
        setDraft((prev) => {
          if (prev === null) return prev;
          const next = clone(prev);
          const segs = path.split('.');
          let cur = next;
          for (let i = 0; i < segs.length - 1; i++) {
            const seg = segs[i];
            if (cur[seg] === undefined) cur[seg] = /^\d+$/.test(segs[i + 1]) ? [] : {};
            cur = cur[seg];
          }
          cur[segs[segs.length - 1]] = value;
          return next;
        });
        setNotice('');
      };

      /** Put a set of paths back to the values currently on disk. */
      const resetPaths = (paths) => {
        if (view === null) return;
        setDraft((prev) => {
          if (prev === null) return prev;
          const next = clone(prev);
          for (const path of paths) {
            const original = getPath(view.spec, path);
            const segs = path.split('.');
            let cur = next;
            for (let i = 0; i < segs.length - 1; i++) {
              const seg = segs[i];
              if (cur[seg] === undefined) cur[seg] = /^\d+$/.test(segs[i + 1]) ? [] : {};
              cur = cur[seg];
            }
            cur[segs[segs.length - 1]] = clone(original);
          }
          return next;
        });
        setNotice('');
      };

      /** How many fields in a set differ from disk — drives「恢复本组默认」. */
      const dirtyCountOf = (fields) => {
        if (view === null || draft === null) return 0;
        let n = 0;
        for (const f of fields) {
          if (!sameValue(getPath(draft, f.path), getPath(view.spec, f.path))) n++;
        }
        return n;
      };

      /**
       * Write a numbering scheme into the draft: the id plus every level's
       * pattern and example. One picker click has to move all three, because a
       * rule table saying scheme=chineseChapter while level 1 still reads
       * "编号形如 1" is the same kind of self-contradiction as A5-sized A4 paper.
       */
      const setNumberingScheme = (id) => {
        const scheme = NUMBERING_SCHEMES[id];
        if (scheme === undefined || draft === null) return;
        setDraft((prev) => {
          if (prev === null) return prev;
          const next = clone(prev);
          const levels = Array.isArray(next.headings?.levels) ? next.headings.levels : [];
          next.headings = next.headings ?? {};
          next.headings.numberingScheme = id;
          for (let i = 0; i < levels.length; i++) {
            const spec = scheme.levels[i];
            if (spec === undefined) continue;
            levels[i].numbering = spec.pattern;
            levels[i].numberingExample = spec.example;
          }
          return next;
        });
        setNotice('');
      };

      /** Field-level ops: only what actually changed, so a save is minimal. */
      const collectOps = (base, next) => {
        const ops = [];
        const walk = (fields) => {
          for (const f of fields) {
            const before = getPath(base, f.path);
            const after = getPath(next, f.path);
            if (!sameValue(after, before) && after !== undefined) ops.push([f.path, after]);
          }
        };
        for (const g of GROUPS) walk(g.fields);
        const levels = Array.isArray(next.headings?.levels) ? next.headings.levels : [];
        for (let i = 0; i < levels.length; i++) walk(levelFields(i));
        // The picker's own outputs are not form fields, so their changes would
        // be dropped by the walk above and the save would silently do nothing.
        walk(schemePaths(levels.length).map((path) => ({ path })));
        return ops;
      };

      const save = async () => {
        if (draft === null || view === null) return;
        const ops = collectOps(view.spec, draft);
        if (!ops.length) {
          setNotice('没有改动');
          return;
        }
        setBusy(true);
        setError('');
        setNotice('');
        try {
          const result = await callHost('patch', { ops });
          if (result.ok !== true) {
            const lines = [result.message || '保存失败'];
            for (const e of result.errors ?? []) lines.push(`· ${e.path || '(根)'}：${e.message}`);
            setError(lines.join('\n'));
            return;
          }
          absorb(result);
          setNotice(`已保存 ${ops.length} 项改动，备份：${result.backup}`);
          refreshBackups();
          if (result.warnings && result.warnings.length) {
            setError(result.warnings.map((w) => `提示：${w.path} — ${w.message}`).join('\n'));
          }
        } catch (err) {
          setError(err.message);
        } finally {
          setBusy(false);
        }
      };

      const revert = () => {
        if (view === null) return;
        setDraft(clone(view.spec));
        setJsonText(JSON.stringify(view.spec, null, 2));
        setError('');
        setNotice('已还原为磁盘上的版本');
      };

      const runPreview = async () => {
        setBusy(true);
        setError('');
        try {
          const result = await callHost('preview', { spec: draft });
          if (result.ok !== true) {
            setError(result.message || '预览失败');
            return;
          }
          setPreview(result.text);
        } catch (err) {
          setError(err.message);
        } finally {
          setBusy(false);
        }
      };

      const refreshBackups = async () => {
        try {
          const result = await callHost('backups');
          if (result.ok === true) setBackups(result.items ?? []);
        } catch {
          /* the list is a convenience; a failure must not break the page */
        }
      };

      React.useEffect(() => {
        refreshBackups();
      }, []);

      const restore = async (name) => {
        if (!confirm(`用备份 ${name} 覆盖当前规则表？当前内容会先备份。`)) return;
        setBusy(true);
        setError('');
        try {
          const result = await callHost('restore', { name });
          if (result.ok !== true) {
            const lines = [result.message || '恢复失败'];
            for (const e of result.errors ?? []) lines.push(`· ${e.path || '(根)'}：${e.message}`);
            setError(lines.join('\n'));
            return;
          }
          absorb(result);
          setNotice(`已从 ${name} 恢复`);
          refreshBackups();
        } catch (err) {
          setError(err.message);
        } finally {
          setBusy(false);
        }
      };

      const importJson = async () => {
        setBusy(true);
        setError('');
        setNotice('');
        try {
          const result = await callHost('import', { text: jsonText });
          if (result.ok !== true) {
            const lines = [result.message || '导入失败'];
            for (const e of result.errors ?? []) lines.push(`· ${e.path || '(根)'}：${e.message}`);
            setError(lines.join('\n'));
            return;
          }
          absorb(result);
          setNotice(`已导入并保存，备份：${result.backup}`);
          refreshBackups();
        } catch (err) {
          setError(err.message);
        } finally {
          setBusy(false);
        }
      };

      if (view === null) {
        return h(
          'div',
          { className: 'wws-root' },
          h('div', { className: 'wws-empty' }, error === '' ? '正在读取规则表…' : error),
          error !== '' ? h('div', { className: 'wws-row' }, h('button', { className: 'wws-btn', onClick: load }, '重试')) : null,
        );
      }

      const dirty = hasChanges(view.spec, draft);
      const levels = Array.isArray(draft.headings?.levels) ? draft.headings.levels : [];

      return h(
        'div',
        { className: 'wws-root' },
        h(
          'div',
          { className: 'wws-head' },
          h(
            'div',
            null,
            h('h2', { className: 'wws-title' }, 'Word 写作规范'),
            h(
              'p',
              { className: 'wws-sub' },
              '这里的取值同时决定「生成时的约束」与「审查时的判据」——改一处，两边一起变。',
            ),
          ),
          h(
            'div',
            { className: 'wws-row' },
            h('button', { className: 'wws-btn wws-btn-sm', onClick: load, disabled: busy }, '重新载入'),
            h(
              'button',
              { className: 'wws-btn wws-btn-sm', onClick: revert, disabled: busy || !dirty },
              '放弃改动',
            ),
            h(
              'button',
              { className: 'wws-btn wws-btn-sm wws-btn-primary', onClick: save, disabled: busy || !dirty },
              busy ? '处理中…' : dirty ? '保存' : '已保存',
            ),
          ),
        ),

        h('div', { className: 'wws-meta' }, `${view.path} · ${view.bytes} 字节`),

        error !== '' ? h('div', { className: 'wws-err' }, error) : null,
        notice !== '' ? h('div', { className: 'wws-ok' }, notice) : null,

        h(
          'div',
          { className: 'wws-tabs' },
          h(
            'button',
            { className: 'wws-tab' + (tab === 'form' ? ' wws-tab-on' : ''), onClick: () => setTab('form') },
            '表单',
          ),
          h(
            'button',
            { className: 'wws-tab' + (tab === 'json' ? ' wws-tab-on' : ''), onClick: () => setTab('json') },
            'JSON',
          ),
        ),

        tab === 'form'
          ? h(
              'div',
              { className: 'wws-layout' },
              h(
                'nav',
                { className: 'wws-nav' },
                GROUPS.concat([{ id: 'headings', title: '标题层级', fields: null }]).map((g) =>
                  h(
                    'button',
                    {
                      key: g.id,
                      className: 'wws-nav-item' + (activeSection === g.id ? ' wws-nav-on' : ''),
                      onClick: () => jumpTo(g.id),
                      type: 'button',
                    },
                    g.title,
                    g.fields !== null && dirtyCountOf(g.fields) > 0 ? h('span', { className: 'wws-dot' }) : null,
                  ),
                ),
              ),
              h(
                'div',
                { className: 'wws-body' },
                GROUPS.map((g) =>
                  h(Group, {
                    key: g.id,
                    id: g.id,
                    title: g.title,
                    fields: g.fields,
                    spec: view.spec,
                    draft,
                    onChange: setField,
                    onReset: resetPaths,
                    dirtyCount: dirtyCountOf(g.fields),
                  }),
                ),
                h(
                  'section',
                  { className: 'wws-card', id: 'wws-sec-headings' },
                  h(
                    'div',
                    { className: 'wws-card-head' },
                    h('h3', null, '标题层级', h('span', { className: 'wws-count' }, `${levels.length} 级`)),
                    h('span', { className: 'wws-spacer' }),
                    h(
                      'label',
                      {
                        className: 'wws-inline',
                        htmlFor: 'wws-headings-numberingScheme',
                        title: 'headings.numberingScheme',
                      },
                      h('span', { className: 'wws-inline-label' }, '编号方案'),
                      h(
                        'select',
                        {
                          id: 'wws-headings-numberingScheme',
                          className: 'wws-select',
                          value: draft?.headings?.numberingScheme ?? 'arabicDotted',
                          onChange: (e) => setNumberingScheme(e.target.value),
                        },
                        Object.keys(NUMBERING_SCHEMES).map((id) =>
                          h('option', { key: id, value: id }, NUMBERING_SCHEMES[id].label),
                        ),
                        // A rule table can name a scheme this build does not
                        // know (edited in the JSON tab, or a newer build wrote
                        // it). Show it read-only instead of silently snapping
                        // the picker to the first option on load.
                        Object.keys(NUMBERING_SCHEMES).includes(draft?.headings?.numberingScheme)
                          ? null
                          : h(
                              'option',
                              { value: draft?.headings?.numberingScheme ?? '' },
                              `${draft?.headings?.numberingScheme ?? '（未设置）'}（当前值）`,
                            ),
                      ),
                    ),
                    h(
                      'button',
                      {
                        className: 'wws-btn wws-btn-sm',
                        onClick: () =>
                          resetPaths(
                            levels
                              .flatMap((_, i) => levelFields(i).map((f) => f.path))
                              .concat(schemePaths(levels.length)),
                          ),
                        disabled:
                          levels.every((_, i) => dirtyCountOf(levelFields(i)) === 0) &&
                          dirtyCountOf(schemePaths(levels.length).map((path) => ({ path }))) === 0,
                        title: '把标题层级字段还原成磁盘上的值',
                      },
                      '恢复本组默认',
                    ),
                  ),
                  h(
                    'div',
                    { className: 'wws-note' },
                    `编号方案决定各级编号形态；当前「${
                      NUMBERING_SCHEMES[draft?.headings?.numberingScheme]?.label ??
                      draft?.headings?.numberingScheme ??
                      '未设置'
                    }」。切换方案会同时改写各级的编号示例与正则。`,
                  ),
                  levels.map((lv, i) =>
                    h(
                      'div',
                      { key: i, style: { marginTop: i === 0 ? 0 : 8 } },
                      h('div', { className: 'wws-note' }, `第 ${i + 1} 级 · 编号形如 ${lv.numberingExample || '?'}`),
                      h(
                        'div',
                        { className: 'wws-grid' },
                        levelFields(i).map((f) =>
                          h(Field, { key: f.path, ...f, spec: view.spec, draft, onChange: setField }),
                        ),
                      ),
                    ),
                  ),
                ),
              ),
            )
          : h(
              'section',
              { className: 'wws-card' },
              h('h3', null, '整表编辑', h('span', { className: 'wws-count' }, '导入前会先校验 schema')),
              h('div', { className: 'wws-note' }, '适合批量改或用一份现成配置替换。保存成功后表单会同步刷新。'),
              h('textarea', {
                className: 'wws-textarea',
                value: jsonText,
                spellCheck: false,
                onChange: (e) => setJsonText(e.target.value),
              }),
              h(
                'div',
                { className: 'wws-row' },
                h('button', { className: 'wws-btn', onClick: () => setJsonText(JSON.stringify(draft, null, 2)) }, '从当前表单取值'),
                h('button', { className: 'wws-btn wws-btn-primary', onClick: importJson, disabled: busy }, '校验并保存'),
              ),
            ),

        h(
          'section',
          { className: 'wws-card' },
          h('h3', null, '提示词预览', h('span', { className: 'wws-count' }, '模型实际看到的文本')),
          h(
            'div',
            { className: 'wws-row' },
            h('button', { className: 'wws-btn', onClick: runPreview, disabled: busy }, '用当前改动渲染'),
            h('span', { className: 'wws-spacer' }),
            preview !== '' ? h('button', { className: 'wws-btn wws-btn-sm', onClick: () => setPreview('') }, '收起') : null,
          ),
          preview !== ''
            ? h('pre', { className: 'wws-pre' }, preview)
            : h('div', { className: 'wws-note' }, '渲染的是当前表单里的值（含未保存的改动），用来确认改动的实际效果。'),
        ),

        h(
          'section',
          { className: 'wws-card' },
          h('h3', null, '历史备份', h('span', { className: 'wws-count' }, `${backups === null ? '…' : backups.length} 份`)),
          h('div', { className: 'wws-note' }, '每次保存都会先把上一版复制到 rules/.backup/，可用它回退。'),
          backups !== null && backups.length === 0
            ? h('div', { className: 'wws-note' }, '还没有备份——保存过一次之后就会出现。')
            : h(
                'div',
                { className: 'wws-list' },
                (backups ?? []).map((b) =>
                  h(
                    'div',
                    { key: b.name, className: 'wws-item' },
                    h('span', { className: 'wws-name' }, b.name),
                    h('span', { className: 'wws-note' }, `${b.bytes} 字节`),
                    h(
                      'button',
                      { className: 'wws-btn wws-btn-sm', onClick: () => restore(b.name), disabled: busy },
                      '恢复',
                    ),
                  ),
                ),
              ),
        ),
      );
    }

    /** Is the draft different from the saved table in any exposed field? */
    function hasChanges(base, draft) {
      if (base === null || draft === null) return false;
      if (JSON.stringify(base) === JSON.stringify(draft)) return false;
      return true;
    }

    /* ---------------------------------------------------------------- */
    /* Plugin                                                            */
    /* ---------------------------------------------------------------- */

    function apply(ctx) {
      const slots = ctx.slots !== undefined ? ctx.slots : ctx.get('slots');
      if (slots === undefined) {
        console.error('[word-writing-spec] slots service unavailable; settings page not registered');
        return;
      }

      const styleEl = document.createElement('style');
      styleEl.setAttribute('data-dsh-plugin', 'word-writing-spec');
      styleEl.textContent = CSS;
      document.head.appendChild(styleEl);
      ctx.effect(() => () => {
        if (styleEl.parentNode) styleEl.parentNode.removeChild(styleEl);
      });

      // No locale service is injected: the page is authored in Chinese and the
      // rule table it edits is Chinese prose, so a translated shell would be
      // the only English text on screen. `label` stays a string for that reason
      // — an older shell renders it directly and would crash on a function.
      ctx.effect(() =>
        slots.inject('settings.section', () =>
          slots.register(
            {
              name: 'settings.section',
              id: 'word-writing-spec',
              // Between the official sections (0–25) and 指令记忆 (1200): this is
              // a third-party tool page, not a core setting.
              order: 1100,
              label: '写作规范',
            },
            () => h(Panel, null),
          ),
        ),
      );
    }

    exports.name = NS;
    // Hard dependency: without it the runtime calls apply() before the slot
    // service is published and the page silently never registers.
    exports.inject = ['slots'];
    exports.apply = apply;
    return module.exports;
  },
});
