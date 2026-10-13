import { SYNC_HEADER, SYNC_PATH } from './sync-protocol';

export { SYNC_PORT } from './sync-protocol';

export interface ListenOptions {
  calendarUrl: string;
  port: number;
  /** How long to wait for the extension (ms); below the executor's 5 min cap */
  waitMs: number;
}

export type ListenResult =
  | { ok: true; url: string; labels: string[] }
  | { ok: false; error: 'PORT_IN_USE' | 'SERVER_ERROR' | 'TIMEOUT' };

/**
 * Node script run through `executeNodeScript` (spawn path: it uses child_process).
 * The body runs inside an async function with `args` in scope; the return value is
 * sent back as JSON, so nothing may be written to stdout.
 *
 * It listens on 127.0.0.1 for the extension's POST, then opens the calendar in the
 * user's normal browser (and so their signed-in profile). The executor strips the
 * child env, so binaries are addressed by absolute path.
 */
export const LISTEN_SCRIPT = `
const cp = require('child_process');
const fs = require('fs');
const http = require('http');
const os = require('os');
const path = require('path');
const opts = args[0];

const openInBrowser = (url) => {
  const run = (file, argv, env) => {
    const child = cp.spawn(file, argv, { detached: true, stdio: 'ignore', env: env || {} });
    child.on('error', () => {});
    child.unref();
  };
  if (process.platform === 'darwin') {
    const chrome = '/Applications/Google Chrome.app';
    run('/usr/bin/open', fs.existsSync(chrome) ? ['-a', chrome, url] : [url]);
  } else if (process.platform === 'win32') {
    const home = os.homedir();
    const systemRoot = ['C:\\\\Windows', path.join(path.parse(home).root, 'Windows')]
      .find((p) => fs.existsSync(p)) || 'C:\\\\Windows';
    const local = path.join(home, 'AppData', 'Local');
    // Chrome needs SystemRoot and the profile folders, which the stripped env lacks.
    const env = {
      SystemRoot: systemRoot, windir: systemRoot, USERPROFILE: home,
      LOCALAPPDATA: local, APPDATA: path.join(home, 'AppData', 'Roaming'),
      TEMP: path.join(local, 'Temp'), TMP: path.join(local, 'Temp'),
    };
    const chrome = [
      'C:\\\\Program Files\\\\Google\\\\Chrome\\\\Application\\\\chrome.exe',
      'C:\\\\Program Files (x86)\\\\Google\\\\Chrome\\\\Application\\\\chrome.exe',
      path.join(local, 'Google', 'Chrome', 'Application', 'chrome.exe'),
    ].find((p) => fs.existsSync(p));
    if (chrome) run(chrome, [url], env);
    else run(path.join(systemRoot, 'explorer.exe'), [url], env);
  } else {
    run('/usr/bin/xdg-open', [url], { HOME: os.homedir(), DISPLAY: ':0' });
  }
};

return await new Promise((resolve) => {
  let timer;
  const server = http.createServer((req, res) => {
    if (req.method !== 'POST' || req.url !== ${JSON.stringify(SYNC_PATH)} ||
        req.headers[${JSON.stringify(SYNC_HEADER)}] !== '1') {
      res.writeHead(404);
      res.end();
      return;
    }
    let body = '';
    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > 2000000) req.destroy();
    });
    req.on('end', () => {
      let data;
      try { data = JSON.parse(body); } catch { data = null; }
      if (!data || !Array.isArray(data.labels)) {
        res.writeHead(400);
        res.end();
        return;
      }
      res.writeHead(200, { 'content-type': 'application/json' });
      res.end('{"ok":true}');
      done({
        ok: true,
        url: String(data.url || ''),
        labels: data.labels.filter((l) => typeof l === 'string'),
      });
    });
  });
  const done = (result) => {
    clearTimeout(timer);
    server.close();
    // keep-alive sockets would keep this process (and the executor call) alive
    server.closeAllConnections();
    resolve(result);
  };
  server.on('error', (e) => done({ ok: false, error: e.code === 'EADDRINUSE' ? 'PORT_IN_USE' : 'SERVER_ERROR' }));
  server.listen(opts.port, '127.0.0.1', () => {
    timer = setTimeout(() => done({ ok: false, error: 'TIMEOUT' }), opts.waitMs);
    openInBrowser(opts.calendarUrl);
  });
});
`;
