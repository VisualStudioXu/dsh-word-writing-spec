# dsh-word-writing-spec

把《AI 辅助开发文档格式规范指南》变成 DSH 里**可执行的**东西：生成时约束、事后审查、可自动修复。

规则只有一份 —— `rules/docx-writing-spec.json`。提示词约束、审查判定、修复写入三者都读它，改一处同时生效，不会出现「提示词说 22 磅、审查却说 28 磅」这种分裂。

## 装上之后得到什么

**1. 生成约束（systemPrompt 注入）**

插件注册一个 `systemPrompt` 段（order 900，约 2.3K 字符），内容是规则表的自然语言渲染：页面设置、正文字体字号行距缩进、三级标题层级与编号空格、表题/图题位置与格式、代码块、页眉页脚页码、分节分页、参考文献著录、命名与备份约定、生成前自检清单。

任何在本会话里写 .docx 的请求，模型开工前就已经拿到这些数值，写出来即合规，不必事后返工。

**2. 四个工具**

| 工具 | 作用 | 是否改文件 |
| --- | --- | --- |
| `word_spec_audit` | 按规范审查 .docx，返回 verdict 与逐条 finding（应/实/可否自动修） | 只读 |
| `word_spec_plan_headings` | §8.6 第一步：扫出「像标题但不是标题样式」的段落，给出候选、置信度、证据 | 只读 |
| `word_spec_fix` | 按规范修复并另存新文件；**修改前自动备份**原文件 | 写新文件 |
| `word_spec_rules` | 读规则表本身，可按小节取 | 只读 |

`word_spec_fix` 的 `acceptedHeadings` 参数接 `word_spec_plan_headings` 返回的候选子集 —— 这是规则里「转换前必须先展示候选等用户确认」的落地方式，工具不会自作主张改标题。

## 审查覆盖的规则组

`page` `body` `headings` `table` `figure` `code` `headerFooter` `reference` `appendix` `cover` `toc` `naming`

判定「正文范围」这件事单独抽到了 `lib/scope.js`：封面标题、目录域条目、题注、代码行都不算正文。不这么分，一份正常文档会报出几十条「字体不符」的假错误。审查和修复共用同一个范围判定，所以「报的问题」和「改的地方」不会错位。

`cover` 组是 `checkCoverAndToc`：`cover.title` 为空时取封面上字号最大的段落当标题（不能用 `meta.title` 做子串匹配 —— 封面的「（遵照《…规范》生成）」引用行也含这个子串，会把真标题判成格式错误），另查目录标题样式与 TOC 域是否存在。

`figure` 组的图题方向读 `figure.captionPosition`（`above`/`below`），编号形态读 `figure.numberPattern`；另查真实有效 DPI（`figure.minDpi`）与图中 `wp:inline`/`wp:anchor`（`figure.embedding`）。表格表题同理读 `table.numberPattern`。

## 修复范围与边界

能自动修：页面设置、正文格式（含颜色）、标题格式与样式定义本身、一级标题分页、附录编号空格、表格字体字号、插图段落行距对齐与图题格式、代码块、页眉页脚字体对齐与页码数字样式、封面标题格式、目录标题格式。

只检查不修（需要人判断或涉及域）：页脚 PAGE 域、页眉首页不同、前置页罗马数字编页、表格是否缺表题、插图是否缺图题（缺图是内容问题，凭空补一行「图 x-x」会在下次插图时错号）、插图 DPI 与嵌入方式、标题是否用了内置样式、TOC 域是否存在。这些在 finding 里 `fixable: false`。

跳过的内容：目录、页码这类域在改完后需要在 Word 里更新（Ctrl+A 后 F9），插件不改域代码。

## 用法

装好后直接在对话里说就行，模型会自己调工具。也可以显式要求：

```
用 word_spec_audit 审一下 D:\项目\文档\设计说明书.docx
```

```
按规范修这份文档，输出到 D:\项目\文档\设计说明书_fixed.docx
```

## 规则表

`rules/docx-writing-spec.json` 是从源文档 `D:\AITrae\word\AI辅助开发文档格式规范指南.docx` 解析出来的，结构由 `rules/docx-writing-spec.schema.json`（draft-07）约束。

`lib/spec.js` 的 `loadSpec()` 按 mtime 热重载 —— 改完 JSON 不用重启 DSH，下一次注入和下一次审查就生效。提示词段传的是 thunk（`text: sectionText`，不是 `sectionText()`），宿主每轮装配都会重新渲染，所以注入和工具读的是同一版规则。

改完跑一次校验：

```bash
node scripts/docx-format-spec/validate-rules.mjs
```

它检查四件事：类型/枚举是否合 schema、必填键是否齐、`checklist[].rule` 路径是否还能解析、`checklist[].label` 里的模板占位符有没有解析不出来的（标签是模板字符串，如 `行距（{body.lineSpacing:spacing}）`，渲染时按规则表现算 —— 这样改规则不会漏改文案）。

## 两张语义表

规则表里有两处值是**代号**而不是完整定义，真值在 `lib/spec.js`：

- `LINE_SPACING_MODES` —— 行距的六档方式。规则表存 `type` 与 `valuePt`，但「`single` 是 1 倍、`atLeast` 用磅、`multiple` 读数值」这件事由这张表定义。写入 OOXML、审查比对、提示词文案三处都读它。
- `NUMBERING_SCHEMES` —— 标题编号的四套方案（`arabicDotted` 1/1.1/1.1.1、`chineseChapter` 第一章/第一节/一、、`chineseOrdinal` 一、/（一）/1.、`arabicParen` 1/（1）/①）。规则表存方案 ID 与各级解析后的正则/示例，审计器与修复器读同一份。

两处都在设置页有下拉，改规则表即改行为，不用动代码。加一个新档位只需改这两个常量（行距）或 `NUMBERING_SCHEMES`（编号），并在 `lib/client.js` 里同步那份镜像表 —— 客户端跑在浏览器里，拿不到 `lib/spec.js`。

## 怎么改

- **改规范数值** → 只动 `rules/*.json`，不用重启。
- **改审查/修复逻辑** → 动 `lib/audit.js` / `lib/fix.js`，需重启。
- **改插件形状**（工具名、参数、注入顺序）→ 动 `lib/index.js`，需重启。
- **改设置页**（字段、下拉档位、样式）→ 动 `lib/client.js`，走 HMR，刷新设置面板即可见。

详版见 `docs/docx-format-spec/editing-guide.md`。

## 目录

```
lib/index.js    插件入口：systemPrompt 段 + 四个工具
lib/spec.js     规则表加载 + 样式名归一化 + 行距语义层（LINE_SPACING_MODES）+ 标题编号方案表（NUMBERING_SCHEMES）
lib/docx.js     docx 包读取（zip/xml 层），段落/表格/节/样式建模
lib/image-size.js  图片像素尺寸与有效 DPI 推算（PNG/JPEG/GIF/BMP，纯字节解析）
lib/zip.js      zip 读写
lib/xml.js      最小 XML DOM
lib/scope.js    正文范围判定（审查与修复共用）
lib/outline.js  标题候选识别与置信度
lib/audit.js    审查引擎
lib/fix.js      修复引擎（含 §8.6 两步式标题转换、备份保存）
lib/settings-route.js  设置页的宿主接口（读写规则表、备份、SSE 推送）
lib/net-guard.js       接口的 Host/Origin/Sec-Fetch-Site 校验
lib/schema-check.js    schema 校验与字段路径解析（路由与 CLI 共用）
lib/client.js          设置页客户端半（手写，无构建步骤）
rules/          规则唯一来源 + JSON Schema
```

## 设计取舍

**为什么自己解析 docx，不用 python-docx 或 mammoth**：审查要读到 `w:firstLineChars`、`w:firstLine` 并存时的真实语义、`w:rFonts` 的主题字体继承、样式链上的有效 rPr —— 这些高层库会抹平或者不给。修复要按 CT_PPr/CT_RPr 的 schema 顺序插入子元素，顺序错了 Word 会拒绝打开文件。

**为什么备份是强制的**：规则里写了「修改任何已有文档前先备份」。`saveWithBackup()` 先 `copyFile` 原件再写输出，不靠调用方记得。

**为什么审计和修复共用 `scope.js`**：否则报错的位置和修改的位置会错位，用户改完发现错误还在。

## 已知限制

- 图片分辨率（`figure.minDpi`）的 DPI 由 `lib/image-size.js` 从图片字节直接推算（支持 PNG / JPEG / GIF / BMP）。读不出像素尺寸的格式（如 EMF/WMF、SVG、外部链接图）会跳过这项检查而不是误报。
- 颜色只比对「显式写了什么」，不解析主题色（`w:themeColor`）到具体值。规则表要求非黑、而文档用主题色时，会按「不符」报出来。
- 页码数字样式只写到最后一节的 `pgNumType`，不逐节处理 —— 前面的节通常合法地使用罗马数字。
- `pageNumberFormat` 只认「数字」一个占位符。需要「第 1 页 / 共 20 页」这种双计数器时，`NUMPAGES` 那部分要手工在 Word 里加 —— 单键模板没有表达两个变量的语法。
- 表格样式名依赖 `styles.xml` 定义；样式 id 是纯数字的文档（Word 导出常见）已做兼容，但仍以英文样式名为准。
- LibreOffice 渲染时若系统缺 `楷体_GB2312` / `仿宋_GB2312` / `Courier`，会回退到其他字体显示 —— 这不影响文件本身，只是预览观感。

## 许可

MIT
