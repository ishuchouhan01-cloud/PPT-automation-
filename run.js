/**
 * EDOS Weekly Review Cloud — launcher.
 *
 * ONE process, ONE address. It starts the generator in the background, serves the website from
 * ./web, forwards anything under /api to the generator, and opens the browser.
 *
 * This file belongs to the distribution only. The development workspace runs the website through a
 * separate dev server on another port; an end user should not have to know that, or start two things,
 * or type a command. Nothing here touches how a deck is produced — it is plumbing.
 *
 * No dependencies beyond Node's own modules, so it cannot break if the shipped packages change.
 */
'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');

const ROOT = __dirname;
const WEB = path.join(ROOT, 'web');
const API_PORT = Number(process.env.EDOS_API_PORT || 4000);
const WEB_PORT = Number(process.env.EDOS_PORT || 3000);

// ------------------------------------------------------------------------------------------------
// SHARING — the password gate and the public tunnel.
//
// Both are OFF unless asked for. Started normally, this file behaves exactly as it always has: the
// website is on 127.0.0.1 only, and there is no login. Nothing below changes how a deck is produced.
//
// Settings live in sharing.json rather than edos_config.json, because that file is business data
// that gets edited and passed around; a password does not belong in it.
// ------------------------------------------------------------------------------------------------
const SHARE = process.argv.includes('--share');
let shareCfg = { password: '', cloudflared: '' };
try {
  const raw = fs.readFileSync(path.join(ROOT, 'sharing.json'), 'utf8');
  shareCfg = { ...shareCfg, ...JSON.parse(raw) };
} catch { /* absent or unreadable — sharing simply stays off */ }
const PASSWORD = String(process.env.EDOS_PASSWORD || shareCfg.password || '');
const GATED = PASSWORD.length > 0;

/**
 * The session token.
 *
 * Derived from a secret generated fresh at every start, so stopping the window invalidates every
 * session that was open — which is the behaviour you want from something you switch on to show
 * somebody and then switch off. The password itself is never written into the cookie.
 */
const TOKEN = crypto.createHmac('sha256', crypto.randomBytes(32)).update('edos-session').digest('hex');

/** Timing-safe string comparison — a plain `===` on a secret leaks its length and prefix. */
const sameSecret = (a, b) => {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
};

/**
 * Failed attempts, per client.
 *
 * A public tunnel hostname WILL be found and probed — Cloudflare's own certificate log makes that
 * certain — so a shared password with no back-off is a password that gets guessed. Eight tries, then
 * that client waits a minute. The counter is in memory: restarting clears it, which is fine, because
 * restarting also changes the session secret and costs the attacker everything they had.
 */
const fails = new Map();
const clientOf = (req) => String(req.headers['cf-connecting-ip'] || req.socket.remoteAddress || '?');
const lockedOut = (ip) => {
  const f = fails.get(ip);
  return !!(f && f.until > Date.now());
};
const noteFail = (ip) => {
  const f = fails.get(ip) || { n: 0, until: 0 };
  f.n++;
  if (f.n >= 8) { f.until = Date.now() + 60_000; f.n = 0; }
  fails.set(ip, f);
};

const cookiesOf = (req) => Object.fromEntries(
  String(req.headers.cookie || '').split(';').map((c) => {
    const i = c.indexOf('=');
    return i < 0 ? [c.trim(), ''] : [c.slice(0, i).trim(), decodeURIComponent(c.slice(i + 1).trim())];
  }).filter(([k]) => k));

const LOGIN_PAGE = (msg) => `<!doctype html><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>EDOS Weekly Review Cloud</title>
<style>
 body{margin:0;min-height:100vh;display:grid;place-items:center;background:#0f1115;color:#e8e8ea;
      font:16px/1.5 "Helvetica Neue",Segoe UI,system-ui,sans-serif}
 form{width:min(92vw,360px);background:#171a21;border:1px solid #262b36;border-radius:14px;padding:28px}
 h1{font-size:17px;margin:0 0 4px} p{margin:0 0 20px;color:#8b93a4;font-size:13px}
 input{width:100%;box-sizing:border-box;padding:11px 12px;border-radius:8px;border:1px solid #333a49;
       background:#0f1115;color:#e8e8ea;font-size:15px}
 button{width:100%;margin-top:12px;padding:11px;border:0;border-radius:8px;background:#c2410c;
        color:#fff;font-size:15px;font-weight:600;cursor:pointer}
 .err{color:#fca5a5;font-size:13px;margin:10px 0 0}
</style>
<form method="POST" action="/__login">
  <h1>EDOS Weekly Review Cloud</h1>
  <p>This copy is shared. Enter the password you were given.</p>
  <input type="password" name="password" autofocus autocomplete="current-password" placeholder="Password">
  <button type="submit">Open</button>
  ${msg ? `<p class="err">${msg}</p>` : ''}
</form>`;

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.woff': 'font/woff', '.woff2': 'font/woff2',
  '.ttf': 'font/ttf', '.map': 'application/json; charset=utf-8',
};

function fail(msg, hint) {
  console.error('\n============================================================');
  console.error('  Cannot start: ' + msg);
  if (hint) console.error('  ' + hint);
  console.error('============================================================\n');
  process.exitCode = 1;
}

// ---- the pieces this package cannot run without ------------------------------------------------
const serverJs = path.join(ROOT, 'backend', 'dist', 'server.js');
if (!fs.existsSync(serverJs)) {
  fail('backend\\dist\\server.js is missing.',
    'The package is incomplete — extract the ZIP again, keeping every folder.');
  return;
}
if (!fs.existsSync(path.join(ROOT, 'backend', 'node_modules'))) {
  fail('backend\\node_modules is missing.',
    'The package is incomplete — extract the ZIP again, keeping every folder.');
  return;
}
if (!fs.existsSync(path.join(WEB, 'index.html'))) {
  fail('the web folder is missing.',
    'The package is incomplete — extract the ZIP again, keeping every folder.');
  return;
}

// A public link with no password would put the upload form, and everything already generated, in
// front of anyone who finds the hostname. Refused rather than warned about: by the time a warning
// scrolls past, the link is already live.
if (SHARE && !GATED) {
  fail('a public link was asked for, but no password is set.',
    'Open sharing.json and put a password in it, then run SHARE.bat again.');
  return;
}

// ---- 1. the generator ---------------------------------------------------------------------------
// cwd is the package root, so the app finds edos_config.json, assets\ and fonts\ beside it.
const api = spawn(process.execPath, [serverJs], {
  cwd: ROOT,
  env: { ...process.env, PORT: String(API_PORT) },
  stdio: ['ignore', 'inherit', 'inherit'],
});
api.on('exit', (code) => {
  if (code) console.error(`\n[!] The generator stopped unexpectedly (exit code ${code}).`);
  process.exit(code || 0);
});
const shutdown = () => { try { api.kill(); } catch { /* already gone */ } process.exit(0); };
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

// ---- 2. the website, plus /api forwarding -------------------------------------------------------
const serveFile = (res, file) => {
  fs.readFile(file, (err, buf) => {
    if (err) { res.writeHead(404).end('Not found'); return; }
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
      'Cache-Control': 'no-cache',
    });
    res.end(buf);
  });
};

const web = http.createServer((req, res) => {
  // ---- the password gate ------------------------------------------------------------------------
  // In front of BOTH the website and /api, so there is one door rather than two. The browser app
  // needs no knowledge of it: an unauthenticated page request gets the login form, and an
  // unauthenticated /api request gets a 401 in the JSON shape the app already handles — never the
  // login HTML, which would arrive at code expecting JSON and fail in a confusing way.
  if (GATED) {
    const ip = clientOf(req);
    const authed = sameSecret(cookiesOf(req).edos_auth || '', TOKEN);

    if (req.url === '/__login' && req.method === 'POST') {
      if (lockedOut(ip)) {
        res.writeHead(429, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(LOGIN_PAGE('Too many attempts. Wait a minute and try again.'));
        return;
      }
      let body = '';
      req.on('data', (c) => { body += c; if (body.length > 4096) req.destroy(); });
      req.on('end', () => {
        const sent = new URLSearchParams(body).get('password') || '';
        if (!sameSecret(sent, PASSWORD)) {
          noteFail(ip);
          res.writeHead(401, { 'Content-Type': 'text/html; charset=utf-8' });
          res.end(LOGIN_PAGE('That password was not right.'));
          return;
        }
        fails.delete(ip);
        // Secure only when the request actually arrived over HTTPS — through the tunnel it did, and
        // on http://127.0.0.1 it did not. Setting it unconditionally would make the cookie be
        // dropped on the local address and lock you out of your own laptop.
        const https = String(req.headers['x-forwarded-proto'] || '').includes('https');
        res.writeHead(302, {
          'Set-Cookie': `edos_auth=${TOKEN}; Path=/; HttpOnly; SameSite=Lax${https ? '; Secure' : ''}`,
          Location: '/',
        });
        res.end();
      });
      return;
    }

    if (req.url === '/__logout') {
      res.writeHead(302, { 'Set-Cookie': 'edos_auth=; Path=/; HttpOnly; Max-Age=0', Location: '/' });
      res.end();
      return;
    }

    if (!authed) {
      if (req.url.startsWith('/api/')) {
        res.writeHead(401, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          stage: 'Signing in',
          reason: 'This session is not signed in.',
          details: 'Reload the page and enter the password again.',
        }));
        return;
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(LOGIN_PAGE(''));
      return;
    }
  }

  // Anything under /api belongs to the generator. Streamed both ways, so a 100 MB previous deck
  // uploads and a generated file downloads without being held in memory here.
  if (req.url.startsWith('/api/')) {
    const proxy = http.request(
      { host: '127.0.0.1', port: API_PORT, method: req.method, path: req.url, headers: req.headers },
      (r) => { res.writeHead(r.statusCode || 502, r.headers); r.pipe(res); },
    );
    proxy.on('error', () => {
      res.writeHead(502, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        stage: 'Reaching the generator',
        reason: 'The generator is not responding yet.',
        details: 'If this happens right after starting, wait a few seconds and try again. If it ' +
          'continues, close this window and run START.bat again.',
      }));
    });
    req.pipe(proxy);
    return;
  }

  // Static website. Any unknown path returns index.html, so a refresh cannot land on a blank page.
  const clean = decodeURIComponent((req.url || '/').split('?')[0]);
  const target = path.join(WEB, path.normalize(clean).replace(/^(\.\.[/\\])+/, ''));
  if (!target.startsWith(WEB)) { res.writeHead(403).end('Forbidden'); return; }
  fs.stat(target, (err, st) => {
    if (!err && st.isFile()) return serveFile(res, target);
    serveFile(res, path.join(WEB, 'index.html'));
  });
});

web.on('error', (e) => {
  if (e && e.code === 'EADDRINUSE') {
    fail(`something is already using port ${WEB_PORT}.`,
      'Another copy is probably still running — close its window, or set EDOS_PORT to a free port.');
  } else {
    fail(String((e && e.message) || e));
  }
  try { api.kill(); } catch { /* ignore */ }
});

/**
 * Wait until the generator is actually answering before opening the browser.
 *
 * Without this the page can open a second or two before the generator is listening, and the first
 * thing the user sees is "Backend not reachable" — which looks like a broken package and is not. The
 * delay is real and can be long: starting from a network drive or a cold profile, loading the shipped
 * packages took 50 seconds on a slow filesystem, against well under a second from a local disk.
 */
function waitForApi(attempt = 0) {
  const req = http.request(
    { host: '127.0.0.1', port: API_PORT, path: '/api/health', method: 'GET', timeout: 2000 },
    (r) => { r.resume(); r.statusCode === 200 ? ready() : retry(); },
  );
  req.on('error', retry);
  req.on('timeout', () => { req.destroy(); retry(); });
  req.end();

  function retry() {
    if (attempt === 4) {
      console.log('  Still starting — this can take up to a minute the first time,');
      console.log('  or when running from a network drive. Please wait.');
    }
    if (attempt > 150) {                       // ~2.5 minutes, then say so rather than hang silently
      console.error('\n[!] The generator has not started. Close this window and run START.bat again.');
      console.error('    If it keeps happening, see TROUBLESHOOTING.txt.\n');
      return;
    }
    setTimeout(() => waitForApi(attempt + 1), 1000);
  }
}

function ready() {
  const url = `http://127.0.0.1:${WEB_PORT}`;
  console.log('');
  console.log('============================================================');
  console.log('  EDOS Weekly Review Cloud is ready');
  console.log('');
  console.log('  Open:  ' + url);
  if (GATED) console.log('  Password required to open the page.');
  console.log('');
  console.log('  Keep this window open while you use it.');
  console.log('  Close it (or press Ctrl+C) to stop.');
  console.log('============================================================');
  console.log('');
  // Open the default browser. Nothing depends on this working — the address is printed above.
  const cmd = process.platform === 'win32' ? 'cmd' : (process.platform === 'darwin' ? 'open' : 'xdg-open');
  const args = process.platform === 'win32' ? ['/c', 'start', '', url] : [url];
  try { spawn(cmd, args, { detached: true, stdio: 'ignore' }).unref(); } catch { /* open it yourself */ }

  if (SHARE) openTunnel();
}

/**
 * The public link, via a Cloudflare Quick Tunnel.
 *
 * cloudflared dials OUT from this laptop and Cloudflare hands back a hostname that forwards to it.
 * Nothing is opened on the router and no port is exposed — which is why the website keeps listening
 * on 127.0.0.1 only. The tunnel is the only way in, and closing this window closes it.
 *
 * A quick tunnel is anonymous and free, and its hostname is RANDOM AND TEMPORARY: a new one every
 * time, gone when the window closes. That is the right shape for testing and the wrong shape for
 * anything permanent — a named tunnel on a Cloudflare account is the answer for that, and it is a
 * different setup.
 */
function openTunnel() {
  const exe = shareCfg.cloudflared
    || (process.platform === 'win32' ? path.join(ROOT, 'cloudflared.exe') : 'cloudflared');
  if (process.platform === 'win32' && !fs.existsSync(exe)) {
    console.error('  [!] cloudflared.exe is not here, so there is no public link.');
    console.error('      Run SHARE.bat, which fetches it, or see SHARING.txt.');
    return;
  }
  console.log('  Opening the public link…');
  const tun = spawn(exe, ['tunnel', '--no-autoupdate', '--url', `http://127.0.0.1:${WEB_PORT}`], {
    cwd: ROOT, stdio: ['ignore', 'pipe', 'pipe'],
  });
  let announced = false;
  // cloudflared writes its banner to stderr; the hostname is the only part worth showing.
  const scan = (buf) => {
    const m = String(buf).match(/https:\/\/[-a-z0-9]+\.trycloudflare\.com/i);
    if (!m || announced) return;
    announced = true;
    console.log('');
    console.log('============================================================');
    console.log('  PUBLIC LINK — anyone with this address can reach the app');
    console.log('');
    console.log('    ' + m[0]);
    console.log('');
    console.log('  Password:  ' + PASSWORD);
    console.log('');
    console.log('  Send the two separately. The link dies when this window');
    console.log('  closes, and a restart gives a different one.');
    console.log('============================================================');
    console.log('');
  };
  tun.stdout.on('data', scan);
  tun.stderr.on('data', scan);
  tun.on('exit', (c) => {
    if (!announced) console.error(`  [!] cloudflared stopped before giving a link (exit ${c}). See SHARING.txt.`);
  });
  const stopTunnel = () => { try { tun.kill(); } catch { /* already gone */ } };
  process.on('exit', stopTunnel);
  process.on('SIGINT', stopTunnel);
  process.on('SIGTERM', stopTunnel);
}

web.listen(WEB_PORT, '127.0.0.1', () => {
  console.log(`  Website ready on http://127.0.0.1:${WEB_PORT} — waiting for the generator…`);
  waitForApi();
});
