/**
 * Rule-table validation against `rules/docx-writing-spec.schema.json`.
 *
 * No JSON-Schema library is installed, so this walks the subset the schema
 * actually uses (type/required/enum/properties/items/minItems/maxItems/minimum/
 * maximum/pattern). Deliberately narrow: an unsupported keyword shows up as a
 * hole in THIS file, not as a silently passing document.
 *
 * Shared by the settings route (reject a bad edit before it reaches the disk)
 * and by scripts/docx-format-spec/validate-rules.mjs (check the shipped table),
 * so the page and the CLI cannot disagree about what "valid" means.
 */

import { checklistLabel } from './spec.js';

const typeOf = (v) => {
  if (v === null) return 'null';
  if (Array.isArray(v)) return 'array';
  if (Number.isInteger(v)) return 'integer';
  return typeof v;
};

/**
 * Walk one value against its schema node.
 * `errors` collects `{ path, message }` with a dotted path for the UI.
 */
export function walkSchema(node, schema, at, errors) {
  if (!schema || typeof schema !== 'object') return;

  if (schema.enum && !schema.enum.includes(node)) {
    errors.push({ path: at, message: `必须是 ${JSON.stringify(schema.enum)} 之一（实际 ${JSON.stringify(node)}）` });
  }
  if (schema.type) {
    const want = Array.isArray(schema.type) ? schema.type : [schema.type];
    const got = typeOf(node);
    // integer is also a valid number
    const ok = want.includes(got) || (got === 'integer' && want.includes('number'));
    if (!ok) {
      errors.push({ path: at, message: `应为 ${want.join('|')}，实际 ${got}（${JSON.stringify(node)}）` });
      return; // no point descending into a wrong-typed node
    }
  }
  if (typeof node === 'number') {
    if (schema.minimum != null && node < schema.minimum) {
      errors.push({ path: at, message: `不能小于 ${schema.minimum}（实际 ${node}）` });
    }
    if (schema.maximum != null && node > schema.maximum) {
      errors.push({ path: at, message: `不能大于 ${schema.maximum}（实际 ${node}）` });
    }
  }
  if (typeof node === 'string' && schema.pattern) {
    let ok = true;
    try {
      ok = new RegExp(schema.pattern).test(node);
    } catch {
      ok = true; // a broken pattern in the schema is not the data's fault
    }
    if (!ok) errors.push({ path: at, message: `不匹配 /${schema.pattern}/` });
  }
  if (Array.isArray(node)) {
    if (schema.minItems != null && node.length < schema.minItems) {
      errors.push({ path: at, message: `至少 ${schema.minItems} 项（实际 ${node.length}）` });
    }
    if (schema.maxItems != null && node.length > schema.maxItems) {
      errors.push({ path: at, message: `最多 ${schema.maxItems} 项（实际 ${node.length}）` });
    }
    if (schema.items) node.forEach((v, i) => walkSchema(v, schema.items, `${at}[${i}]`, errors));
  }
  if (node && typeof node === 'object' && !Array.isArray(node)) {
    if (schema.required) {
      for (const k of schema.required) {
        if (!(k in node)) errors.push({ path: at, message: `缺少必需字段「${k}」` });
      }
    }
    if (schema.properties) {
      for (const [k, v] of Object.entries(node)) {
        if (schema.properties[k]) walkSchema(v, schema.properties[k], at ? `${at}.${k}` : k, errors);
        else if (schema.additionalProperties === false) {
          errors.push({ path: at, message: `多余的字段「${k}」` });
        }
      }
    }
  }
}

/**
 * Validate a rule table and its checklist cross-references.
 *
 * The checklist is validated too because it is the self-check the host renders
 * into the prompt: a `rule` path that no longer resolves makes a self-check item
 * silently point at nothing.
 */
export function validateSpec(spec, schema) {
  const errors = [];
  walkSchema(spec, schema, '', errors);

  for (const item of spec.checklist ?? []) {
    if (!item.rule) continue;
    if (resolveRulePath(spec, item.rule) === undefined) {
      errors.push({ path: `checklist.${item.id}`, message: `rule 路径「${item.rule}」解析不到任何字段` });
    }
  }
  return errors;
}

/**
 * Resolve a dotted `rule` path from the checklist.
 *
 * The paths may name a GROUP rather than one key: "body.font" covers fontCN +
 * fontEN, "headerFooter.header" covers headerContent and friends. So every
 * segment but the last must be an exact key, and the last may be a prefix of
 * one or more keys. `headings.1` addresses a 1-based level entry.
 */
export function resolveRulePath(spec, dotted) {
  const levelMatch = /^headings\.(\d+)$/.exec(dotted);
  if (levelMatch) return spec.headings?.levels?.[Number(levelMatch[1]) - 1];

  const segs = dotted.split('.');
  let cur = spec;
  for (let i = 0; i < segs.length; i++) {
    const seg = segs[i];
    if (cur == null || typeof cur !== 'object') return undefined;
    if (Array.isArray(cur)) {
      const idx = Number(seg);
      if (!Number.isInteger(idx) || idx < 0 || idx >= cur.length) return undefined;
      cur = cur[idx];
      continue;
    }
    if (seg in cur) {
      cur = cur[seg];
      continue;
    }
    if (i === segs.length - 1) {
      const hits = Object.keys(cur).filter((k) => k.startsWith(seg));
      return hits.length ? hits : undefined;
    }
    return undefined;
  }
  return cur;
}

/**
 * Labels are templates now ("行距（{body.lineSpacing:spacing}）"), so a stale
 * number is impossible: the renderer reads the same table the audit does.
 *
 * What can still go wrong is the template itself — a path typo or a formatter
 * that does not exist would render as a literal `{body.fnt}` in the checklist.
 * `resolveLabel` leaves unknown tokens in place instead of blanking them, and
 * this reports each one so the typo is visible before the list is published.
 */
export function staleLabelWarnings(spec) {
  const warnings = [];
  for (const item of spec.checklist ?? []) {
    const raw = String(item?.label ?? '');
    const rendered = checklistLabel(item, spec);
    const left = rendered.match(/\{[A-Za-z0-9_.]+(?::[A-Za-z]+)?\}/g);
    if (!left) continue;
    for (const token of new Set(left)) {
      warnings.push({
        path: `checklist.${item.id}`,
        message: `自检项「${raw}」里的占位符 ${token} 解析不到，检查路径或格式化器（spacing/bold/pageBreak/range/align/color）`,
      });
    }
  }
  return warnings;
}

/** Apply `[[dottedPath, value], ...]` onto a copy of the table. */
export function applyOps(spec, ops) {
  const next = JSON.parse(JSON.stringify(spec));
  for (const [dotted, value] of ops) {
    const segs = String(dotted).split('.');
    let cur = next;
    for (let i = 0; i < segs.length - 1; i++) {
      const seg = segs[i];
      if (cur == null || typeof cur !== 'object') throw new Error(`路径「${dotted}」在「${segs.slice(0, i).join('.')}」处不存在`);
      cur = cur[seg];
    }
    const last = segs[segs.length - 1];
    if (cur == null || typeof cur !== 'object') throw new Error(`路径「${dotted}」的父字段不存在`);
    // Only overwrite a key that already exists: an unknown key would pass the
    // schema (additionalProperties is not false everywhere) and then be ignored
    // by every reader, so the page would look like it saved something.
    if (Array.isArray(cur)) {
      const idx = Number(last);
      if (!Number.isInteger(idx) || idx < 0 || idx >= cur.length) throw new Error(`数组下标「${last}」越界`);
      cur[idx] = value;
    } else {
      if (!(last in cur)) throw new Error(`规则表没有字段「${dotted}」`);
      cur[last] = value;
    }
  }
  return next;
}
