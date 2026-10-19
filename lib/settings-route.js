/**
 * Settings route: the HTTP surface the settings page talks to.
 *
 * Why a route and not a tool: the settings page runs in the browser and cannot
 * touch the filesystem, while the rule table lives next to this package. A tiny
 * same-origin POST API is the smallest bridge, and it keeps the write path in
 * ONE place that validates and backs up — a second write path (a CLI, an editor
 * plugin) is another chance to skip the checks.
 *
 * Exclusions, on purpose:
 * - No arbitrary file access. Every method reads or writes exactly
 *   `rules/docx-writing-spec.json` (plus its own backup directory). An API that
 *   took a path would be a remote file-write primitive on the user's machine.
 * - No `ctx.inject = ['webServer']` hard dependency. headless and tui sessions
 *   have no web server; a hard dependency would leave the whole row pending and
 *   take the prompt section and all four tools down with it.
 */

import fs from 'node:fs';
import path from 'node:path';

import { invalidateSpec, loadSpec, RULES_PATH, SCHEMA_PATH } from './spec.js';
import { applyOps, staleLabelWarnings, validateSpec } from './schema-check.js';
import { requestOriginVerdict } from './net-guard.js';

const MAX_BODY_CHARS = 1024 * 1024;
const BACKUP_KEEP = 20;
const POLL_MS = 1500;

/* ------------------------------------------------------------------ */
/* Disk                                                                */
/* ------------------------------------------------------------------ */

const stamp = (d) => {
  const p = (n, w = 2) => String(n).padStart(w, '0');
  // Milliseconds matter: two saves inside one second would otherwise share a
  // name, and the second `copyFileSync` would overwrite the first backup —
  // silently narrowing the window the page offers to undo its own mistake.
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}${p(d.getMilliseconds(), 3)}`;
};

const BACKUP_NAME = /^docx-writing-spec-\d{8}-\d{9}\.json$/;

const readSchema = () => JSON.parse(fs.readFileSync(SCHEMA_PATH, 'utf8'));

function readRaw() {
  const stat = fs.statSync(RULES_PATH);
  return { text: fs.readFileSync(RULES_PATH, 'utf8'), mtimeMs: stat.mtimeMs, bytes: stat.size };
}

/**
 * Copy the current table aside before overwriting it.
 *
 * The rule table is the single source of truth, so a bad paste from the page
 * would silently change both the generation guidance and every audit verdict.
 * Timestamped file copies survive a restart, unlike an in-memory undo ring.
 */
function backupRules() {
  const dir = path.join(path.dirname(RULES_PATH), '.backup');
  fs.mkdirSync(dir, { recursive: true });
  const target = path.join(dir, `docx-writing-spec-${stamp(new Date())}.json`);
  fs.copyFileSync(RULES_PATH, target);
  const kept = fs.readdirSync(dir).filter((f) => f.endsWith('.json')).sort();
  for (const old of kept.slice(0, Math.max(0, kept.length - BACKUP_KEEP))) {
    fs.rmSync(path.join(dir, old), { force: true });
  }
  return path.relative(path.dirname(path.dirname(RULES_PATH)), target);
}

/**
 * Write through a temp file in the same directory, then rename.
 *
 * A partial write here would leave the table unparseable, and every later
 * `loadSpec()` call — including the prompt section — would throw.
 */
function writeAtomic(text) {
  const tmp = `${RULES_PATH}.tmp-${process.pid}-${Date.now()}`;
  fs.writeFileSync(tmp, text, 'utf8');
  try {
    fs.renameSync(tmp, RULES_PATH);
  } catch (err) {
    fs.rmSync(tmp, { force: true });
    throw err;
  }
}

/* ------------------------------------------------------------------ */
/* Snapshot                                                            */
/* ------------------------------------------------------------------ */

function snapshot() {
  const raw = readRaw();
  const spec = JSON.parse(raw.text);
  return {
    spec,
    path: RULES_PATH,
    mtimeMs: raw.mtimeMs,
    bytes: raw.bytes,
  };
}

/** Validate a candidate table; returns `{ ok, errors, warnings }`. */
function check(candidate) {
  const errors = validateSpec(candidate, readSchema());
  return { ok: errors.length === 0, errors, warnings: errors.length ? [] : staleLabelWarnings(candidate) };
}

/* ------------------------------------------------------------------ */
/* Methods                                                             */
/* ------------------------------------------------------------------ */

function makeDispatch(broadcast) {
  return async function dispatch(method, args) {
    switch (method) {
      case 'get':
        return { ok: true, ...snapshot() };

      case 'save': {
        const candidate = args?.spec;
        if (candidate === null || typeof candidate !== 'object' || Array.isArray(candidate)) {
          return { ok: false, message: 'spec 必须是一个对象' };
        }
        const verdict = check(candidate);
        if (!verdict.ok) {
          return { ok: false, message: `规则表未通过校验（${verdict.errors.length} 处）`, errors: verdict.errors };
        }
        const backup = backupRules();
        writeAtomic(JSON.stringify(candidate, null, 2) + '\n');
        invalidateSpec();
        const next = snapshot();
        broadcast();
        return { ok: true, backup, warnings: verdict.warnings, ...next };
      }

      case 'patch': {
        // Field-level edits from the form: [[path, value], ...]. Applying them
        // here (not in the browser) keeps one definition of what a path means.
        const ops = Array.isArray(args?.ops) ? args.ops : [];
        if (!ops.length) return { ok: false, message: 'ops 为空' };
        let candidate;
        try {
          candidate = applyOps(loadSpec(), ops);
        } catch (err) {
          return { ok: false, message: err.message };
        }
        const verdict = check(candidate);
        if (!verdict.ok) {
          return { ok: false, message: `规则表未通过校验（${verdict.errors.length} 处）`, errors: verdict.errors };
        }
        const backup = backupRules();
        writeAtomic(JSON.stringify(candidate, null, 2) + '\n');
        invalidateSpec();
        const next = snapshot();
        broadcast();
        return { ok: true, backup, warnings: verdict.warnings, ...next };
      }

      case 'preview': {
        // Deliberately rendered by the host, not by the page: this is meant to
        // show the EXACT text the model will be told, and a second renderer in
        // the browser would be a second thing to keep in sync.
        const { renderSpecSection } = await import('./index.js');
        const text = renderSpecSection(args?.spec ?? loadSpec());
        return { ok: true, text };
      }

      case 'import': {
        const text = args?.text;
        if (typeof text !== 'string') return { ok: false, message: 'text 必须是字符串' };
        let candidate;
        try {
          candidate = JSON.parse(text);
        } catch (err) {
          return { ok: false, message: `JSON 解析失败：${err.message}` };
        }
        const verdict = check(candidate);
        if (!verdict.ok) {
          return {
            ok: false,
            message: `规则表未通过校验（${verdict.errors.length} 处）`,
            errors: verdict.errors,
          };
        }
        const backup = backupRules();
        writeAtomic(JSON.stringify(candidate, null, 2) + '\n');
        invalidateSpec();
        const next = snapshot();
        broadcast();
        return { ok: true, backup, warnings: verdict.warnings, ...next };
      }

      case 'backups': {
        const dir = path.join(path.dirname(RULES_PATH), '.backup');
        if (!fs.existsSync(dir)) return { ok: true, items: [] };
        const items = fs
          .readdirSync(dir)
          .filter((f) => f.endsWith('.json'))
          .sort()
          .reverse()
          .map((f) => {
            const full = path.join(dir, f);
            const stat = fs.statSync(full);
            return { name: f, bytes: stat.size, mtimeMs: stat.mtimeMs };
          });
        return { ok: true, items };
      }

      case 'restore': {
        const name = String(args?.name ?? '');
        // Reject any path separator: the only addressable files are the
        // timestamped copies this route itself wrote.
        if (!BACKUP_NAME.test(name)) {
          return { ok: false, message: '备份名不合法' };
        }
        const full = path.join(path.dirname(RULES_PATH), '.backup', name);
        if (!fs.existsSync(full)) return { ok: false, message: '备份不存在' };
        const candidate = JSON.parse(fs.readFileSync(full, 'utf8'));
        const verdict = check(candidate);
        if (!verdict.ok) {
          return { ok: false, message: '该备份本身未通过校验', errors: verdict.errors };
        }
        const backup = backupRules();
        writeAtomic(JSON.stringify(candidate, null, 2) + '\n');
        invalidateSpec();
        const next = snapshot();
        broadcast();
        return { ok: true, backup, restored: name, warnings: verdict.warnings, ...next };
      }

      default:
        return { ok: false, message: `未知方法「${method}」` };
    }
  };
}

/* ------------------------------------------------------------------ */
/* Route                                                               */
/* ------------------------------------------------------------------ */

/**
 * Register the settings route on this plugin's fiber.
 *
 * Called from `apply()` through `ctx.inject(['webServer'], ...)`; does nothing
 * (and does not throw) where no web server exists.
 */
export function registerSettingsRoute(ctx, { apiPrefix }) {
  let routeDispose = null;
  let guardHost = true;
  /** SSE subscribers: Set of `send` functions. */
  const streams = new Set();

  const broadcast = () => {
    if (!streams.size) return;
    let frame;
    try {
      frame = 'data: ' + JSON.stringify(snapshot()) + '\n\n';
    } catch {
      return; // a broken table is reported by the page's own fetch, not here
    }
    for (const send of streams) {
      try {
        send(frame);
      } catch {
        /* dropped below by its own close handler */
      }
    }
  };

  const dispatch = makeDispatch(broadcast);

  const readBody = (req) =>
    new Promise((resolve) => {
      let size = 0;
      let overflow = false;
      const chunks = [];
      req.on('data', (chunk) => {
        size += chunk.length;
        if (size > MAX_BODY_CHARS) {
          overflow = true;
          return; // keep draining so the socket does not stall
        }
        chunks.push(chunk);
      });
      req.on('end', () => resolve({ text: Buffer.concat(chunks).toString('utf8'), overflow }));
      req.on('error', () => resolve({ text: '', overflow: false }));
    });

  const json = (res, status, payload) => {
    const body = JSON.stringify(payload);
    res.writeHead(status, {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': 'no-store',
      'content-length': Buffer.byteLength(body),
    });
    res.end(body);
  };

  /** Push snapshots whenever the file changes, so two windows stay in step. */
  let lastMtime = 0;
  const tick = () => {
    if (!streams.size) return;
    let mtime = 0;
    try {
      mtime = fs.statSync(RULES_PATH).mtimeMs;
    } catch {
      return;
    }
    if (mtime !== lastMtime) {
      lastMtime = mtime;
      broadcast();
    }
  };
  const poller = setInterval(tick, POLL_MS);
  if (typeof poller.unref === 'function') poller.unref();

  const handleEventStream = (req, res) => {
    res.writeHead(200, {
      'content-type': 'text/event-stream; charset=utf-8',
      'cache-control': 'no-store',
      connection: 'keep-alive',
    });
    res.write('retry: 5000\n\n');

    let closed = false;
    const cleanup = () => {
      if (closed) return;
      closed = true;
      streams.delete(send);
      try {
        res.end();
      } catch {
        /* already gone */
      }
    };
    const send = (frame) => {
      if (closed) return;
      try {
        res.write(frame);
      } catch {
        cleanup();
      }
    };

    req.on('close', cleanup);
    req.on('error', cleanup);
    res.on('error', cleanup);
    streams.add(send);
    try {
      lastMtime = fs.statSync(RULES_PATH).mtimeMs;
    } catch {
      /* reported by the first fetch */
    }
    try {
      send('data: ' + JSON.stringify(snapshot()) + '\n\n');
    } catch {
      cleanup();
    }
  };

  const handler = async (req, res) => {
    // An interrupted socket must not take the host process down.
    req.on('error', () => {});
    res.on('error', () => {});

    const verdict = requestOriginVerdict(req.headers, guardHost);
    if (!verdict.ok) {
      json(res, 403, { ok: false, message: verdict.reason });
      return;
    }

    const accept = String(req.headers.accept ?? '');
    if (req.method === 'GET') {
      if (accept.includes('text/event-stream')) {
        handleEventStream(req, res);
        return;
      }
      json(res, 405, { ok: false, message: '请用 POST 调用，或 GET + Accept: text/event-stream 订阅' });
      return;
    }
    if (req.method !== 'POST') {
      json(res, 405, { ok: false, message: `不支持 ${req.method}` });
      return;
    }
    if (!String(req.headers['content-type'] ?? '').toLowerCase().includes('application/json')) {
      json(res, 415, { ok: false, message: 'Content-Type 必须是 application/json' });
      return;
    }

    const { text, overflow } = await readBody(req);
    if (overflow) {
      json(res, 413, { ok: false, message: '请求体过大' });
      return;
    }
    let payload;
    try {
      payload = JSON.parse(text);
    } catch (err) {
      json(res, 400, { ok: false, message: `请求体不是 JSON：${err.message}` });
      return;
    }

    let result;
    try {
      result = await dispatch(payload?.method, payload?.args ?? null);
    } catch (err) {
      json(res, 200, { ok: false, message: `处理失败：${err.message}` });
      return;
    }
    // Errors travel as 200 + ok:false so the page can render the message; a
    // non-2xx would only give it a status code.
    json(res, 200, result);
  };

  // `inject` is how a plugin waits for a service it did not declare as a hard
  // dependency. A host without it (headless, or an older build) cannot serve the
  // page at all, so bail out instead of throwing: an exception here would take
  // the prompt section and the four tools down with the settings page.
  if (typeof ctx.inject !== 'function') {
    console.error('[word-writing-spec] ctx.inject unavailable; settings page not served');
    return;
  }

  ctx.inject(['webServer'], (scope) => {
    const webServer = scope.webServer;
    if (webServer === undefined || typeof webServer.register !== 'function') {
      console.error('[word-writing-spec] webServer service unavailable; settings page cannot reach the host');
      return;
    }
    const boundHost = typeof webServer.host === 'string' ? webServer.host.trim().toLowerCase() : '';
    // Loopback names are compared here rather than through net-guard's
    // canonicalIp: `host` may be a name ("localhost"), and a non-loopback bind
    // means the operator published the server deliberately, where no Host value
    // is knowable in advance.
    guardHost =
      boundHost === '' ||
      boundHost === 'localhost' ||
      boundHost === '127.0.0.1' ||
      boundHost === '::1' ||
      boundHost === '[::1]';
    try {
      routeDispose = webServer.register({ kind: 'prefix', path: apiPrefix, handler });
    } catch (err) {
      console.error('[word-writing-spec] failed to register settings route:', err);
      routeDispose = null;
    }
    // The sub-fiber owns the route, but fall back to the plugin fiber if the
    // scope has no `effect` of its own — losing the disposer would leak the
    // route across a reload.
    const owner = typeof scope.effect === 'function' ? scope : ctx;
    owner.effect(() => () => {
      if (routeDispose !== null) {
        const dispose = routeDispose;
        routeDispose = null;
        try {
          dispose();
        } catch {
          /* already disposed */
        }
      }
    }, 'word-writing-spec: settings route');
  });

  ctx.effect(() => () => {
    clearInterval(poller);
    for (const send of [...streams]) {
      try {
        send('event: bye\ndata: {}\n\n');
      } catch {
        /* closing anyway */
      }
    }
    streams.clear();
  }, 'word-writing-spec: settings route teardown');
}
