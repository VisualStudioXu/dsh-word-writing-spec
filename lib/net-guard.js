/**
 * Host / Origin / Sec-Fetch-Site checks for the settings route.
 *
 * Why a route that only the local GUI calls still needs them: the content-type
 * check alone stops a cross-site form post, and the absence of CORS headers
 * stops a cross-origin JSON fetch in preflight. Both assume the browser sees the
 * request as cross-origin. DNS rebinding removes that assumption — a page at
 * evil.com whose DNS is re-pointed at 127.0.0.1 becomes same-origin with the
 * server it then talks to, so no preflight happens and the plugin would happily
 * rewrite the user's rule table. The one thing the attacker cannot forge is the
 * `Host` header the browser sends: it still says evil.com.
 *
 * So when the server is bound to loopback (the default) the route demands a
 * loopback `Host`. An operator who binds a non-loopback address has
 * deliberately published the server beyond this machine, where no Host value is
 * knowable in advance; the check is disabled there rather than 403-ing every
 * legitimate client.
 */

import { isIP } from 'node:net';

const LOOPBACK_HOSTNAMES = new Set(['localhost', '127.0.0.1', '::1', '0:0:0:0:0:0:0:1', '[::1]', 'localhost.']);

/**
 * Split a `Host`-style value into a bare hostname.
 *
 * A value with more than one colon is an unbracketed IPv6 literal (the colons
 * are the address, not a port); with at most one colon it is `host:port`.
 */
export function hostnameOf(value) {
  if (typeof value !== 'string') return '';
  const trimmed = value.trim().toLowerCase();
  if (trimmed === '') return '';
  if (trimmed.startsWith('[')) {
    const end = trimmed.indexOf(']');
    return end === -1 ? trimmed.slice(1) : trimmed.slice(1, end);
  }
  if (trimmed.indexOf(':') !== trimmed.lastIndexOf(':')) return trimmed;
  const colon = trimmed.lastIndexOf(':');
  return colon === -1 ? trimmed : trimmed.slice(0, colon);
}

/**
 * Normalise a hostname the same way on BOTH sides of the Origin comparison.
 *
 * `new URL('http://[::1]:8080').hostname` keeps the brackets, while
 * `hostnameOf('[::1]:8080')` strips them; comparing the raw forms rejects every
 * same-origin request over IPv6 loopback. Strip one surrounding pair so both
 * sides spell an IPv6 literal identically.
 */
const normalizeHostname = (value) =>
  typeof value === 'string' && value.startsWith('[') && value.endsWith(']') && value.length >= 2
    ? value.slice(1, -1)
    : value;

/**
 * Expand an IP literal to a comparable canonical string; null if not an IP.
 *
 * Needed because the two sides of the Origin check spell the SAME address
 * differently: a browser sends Host `[::ffff:127.0.0.1]` while the URL parser
 * canonicalises Origin to `[::ffff:7f00:1]`.
 */
export function canonicalIp(value) {
  if (typeof value !== 'string') return null;
  let addr = normalizeHostname(value.trim().toLowerCase());
  const percent = addr.indexOf('%');
  if (percent !== -1) addr = addr.slice(0, percent);
  if (isIP(addr) === 4) return 'v4:' + addr.split('.').map((part) => Number(part)).join('.');
  if (isIP(addr) !== 6) return null;

  // Translate a trailing dotted quad (::ffff:127.0.0.1) into two hex groups so
  // it can meet the parser's canonical form.
  const v4 = /(\d{1,3}(?:\.\d{1,3}){3})$/i.exec(addr);
  if (v4 !== null) {
    const octets = v4[1].split('.').map((part) => Number(part));
    if (octets.some((octet) => octet > 255)) return null;
    const prefix = addr.slice(0, v4.index);
    const hi = ((octets[0] << 8) | octets[1]).toString(16);
    const lo = ((octets[2] << 8) | octets[3]).toString(16);
    // Keep the trailing ':' of the prefix as the separator; it may also be the
    // second ':' of a '::' compression, which must survive.
    addr = (prefix.endsWith(':') ? prefix : prefix + ':') + hi + ':' + lo;
  }

  const halves = addr.split('::');
  if (halves.length > 2) return null;
  const normalizeGroups = (groups) => groups.map((g) => parseInt(g, 16).toString(16).padStart(4, '0')).join(':');
  const validGroups = (groups) => groups.length === 8 && groups.every((g) => /^[0-9a-f]{1,4}$/i.test(g));

  if (halves.length === 1) {
    const groups = halves[0].split(':');
    return validGroups(groups) ? normalizeGroups(groups) : null;
  }
  const [leftRaw, rightRaw] = halves;
  const left = leftRaw === '' ? [] : leftRaw.split(':');
  const right = rightRaw === '' ? [] : rightRaw.split(':');
  const missing = 8 - left.length - right.length;
  if (missing < 0) return null;
  const groups = left.concat(Array(missing).fill('0'), right);
  return validGroups(groups) ? normalizeGroups(groups) : null;
}

/** Same endpoint: exact text or the same IP literal in another spelling. */
export function sameHostname(a, b) {
  if (a === b) return true;
  const ca = canonicalIp(a);
  const cb = canonicalIp(b);
  return ca !== null && cb !== null && ca === cb;
}

/** Loopback by exact name or by equivalent IP spelling. */
export function isLoopbackName(name) {
  if (LOOPBACK_HOSTNAMES.has(name)) return true;
  const canonical = canonicalIp(name);
  if (canonical === null) return false;
  for (const known of LOOPBACK_HOSTNAMES) {
    const knownCanonical = canonicalIp(known);
    if (knownCanonical !== null && knownCanonical === canonical) return true;
  }
  return false;
}

/**
 * Decide whether a settings-route request may proceed.
 *
 * `guarded` is true when the server is bound to loopback. Each rejection is
 * returned with a human-readable reason so the client can show it instead of a
 * bare 403.
 */
export function requestOriginVerdict(headers, guarded) {
  const read = (name) => {
    if (headers === null || typeof headers !== 'object') return '';
    const value = headers[name];
    return typeof value === 'string' ? value : '';
  };

  const requestHost = read('host').trim().toLowerCase();
  const requestHostname = hostnameOf(requestHost);

  if (guarded && !isLoopbackName(requestHostname)) {
    return {
      ok: false,
      reason: 'Host 头不是回环地址（可能是 DNS 重绑定攻击）：' + (requestHost === '' ? '(缺失)' : requestHost),
    };
  }

  // Only `cross-site` is refused: schemeful same-site ignores ports, so the GUI
  // and this route count as same-site despite different ports. Non-browser
  // clients never send the header, so they are governed by the checks above.
  const fetchSite = read('sec-fetch-site').trim().toLowerCase();
  if (fetchSite === 'cross-site') {
    return { ok: false, reason: '请求被浏览器标记为跨站（Sec-Fetch-Site: cross-site）' };
  }

  const origin = read('origin').trim();
  if (origin !== '') {
    let originHostname = '';
    try {
      originHostname = normalizeHostname(new URL(origin).hostname.toLowerCase());
    } catch {
      originHostname = '';
    }
    // Compare hostnames, not host:port — a proxy may rewrite the port, and the
    // hostname is what rebinding actually attacks. `Origin: null` parses to no
    // hostname and is refused.
    if (originHostname === '' || !sameHostname(originHostname, requestHostname)) {
      return { ok: false, reason: 'Origin 与 Host 不同源：' + origin };
    }
  }

  return { ok: true, reason: null };
}
