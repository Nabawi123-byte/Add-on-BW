// Server HTTPS lokal untuk add-in Master (https://localhost:3000).
// Melayani dist/ dan salinan office.js lokal sehingga add-in tetap terbuka tanpa internet.

import https from 'node:https';
import { createReadStream, existsSync, statSync, readFileSync } from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const root = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const PORT = Number(process.env.MASTER_PORT || 3000);
const DIST = path.join(root, 'dist');
const OFFICE = path.join(root, 'node_modules', '@microsoft', 'office-js', 'dist');

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.map': 'application/json',
  '.xml': 'application/xml',
  '.ico': 'image/x-icon',
};

async function httpsOptions() {
  try {
    const devCerts = require('office-addin-dev-certs');
    return await devCerts.getHttpsServerOptions(3650);
  } catch (e) {
    const dir = path.join(os.homedir(), '.office-addin-dev-certs');
    const key = path.join(dir, 'localhost.key');
    const cert = path.join(dir, 'localhost.crt');
    if (existsSync(key) && existsSync(cert)) return { key: readFileSync(key), cert: readFileSync(cert) };
    console.error('Sertifikat HTTPS belum ada. Jalankan: npx office-addin-dev-certs install');
    process.exit(1);
  }
}

function resolveFile(urlPath) {
  let p = decodeURIComponent(urlPath.split('?')[0]);
  if (p === '/' || p === '') p = '/taskpane.html';
  let base = DIST;
  if (p.startsWith('/office-js/')) {
    base = OFFICE;
    p = p.slice('/office-js'.length);
  }
  const file = path.normalize(path.join(base, p));
  if (file !== base && !file.startsWith(base + path.sep)) return null; // cegah path traversal
  if (!existsSync(file) || !statSync(file).isFile()) return null;
  return file;
}

if (!existsSync(path.join(DIST, 'taskpane.html'))) {
  console.error('Folder dist/ belum ada. Jalankan: npm run build');
  process.exit(1);
}

const options = await httpsOptions();

function handler(req, res) {
  const file = resolveFile(req.url || '/');
  if (!file) {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Tidak ditemukan');
    return;
  }
  const isOffice = file.startsWith(OFFICE);
  res.writeHead(200, {
    'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
    'Cache-Control': isOffice ? 'public, max-age=604800' : 'no-cache',
    'X-Content-Type-Options': 'nosniff',
  });
  createReadStream(file).pipe(res);
}

// "localhost" bisa mengarah ke 127.0.0.1 (IPv4) atau ::1 (IPv6); dengarkan keduanya, hanya di komputer ini.
let listening = 0;
for (const host of ['127.0.0.1', '::1']) {
  const server = https.createServer(options, handler);
  server.on('error', (e) => {
    if (e.code === 'EADDRINUSE') {
      console.log(`Port ${PORT} (${host}) sudah dipakai — kemungkinan server Master sudah berjalan.`);
      if (!listening && host === '::1') process.exit(0);
    } else if (e.code === 'EADDRNOTAVAIL' || e.code === 'EAFNOSUPPORT') {
      // IPv6 tidak tersedia: abaikan
    } else {
      throw e;
    }
  });
  server.listen(PORT, host, () => {
    listening++;
    if (listening === 1) console.log(`Master berjalan di https://localhost:${PORT}  (Ctrl+C untuk berhenti)`);
  });
}
