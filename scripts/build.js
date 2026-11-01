// Membangun folder dist/ (yang dilayani server lokal).
import { build } from 'esbuild';
import { cpSync, mkdirSync, rmSync, existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dist = path.join(root, 'dist');
const src = (...p) => path.join(root, 'src', ...p);

rmSync(dist, { recursive: true, force: true });
mkdirSync(path.join(dist, 'assets'), { recursive: true });

const common = {
  bundle: true,
  format: 'iife',
  platform: 'browser',
  // Office modern: WebView2 (Edge/Chromium) di Windows, WKWebView (Safari) di Mac.
  target: ['es2020', 'chrome100', 'edge100', 'safari15'],
  minify: true,
  sourcemap: true,
  legalComments: 'none',
  logLevel: 'warning',
};

await build({ ...common, entryPoints: [src('ui', 'app.js')], outfile: path.join(dist, 'taskpane.js') });
await build({ ...common, entryPoints: [src('functions', 'functions.js')], outfile: path.join(dist, 'functions.js') });

for (const f of ['taskpane.html', 'taskpane.css', 'commands.html']) cpSync(src('ui', f), path.join(dist, f));
for (const f of ['functions.html', 'functions.json']) cpSync(src('functions', f), path.join(dist, f));
if (existsSync(src('ui', 'assets'))) cpSync(src('ui', 'assets'), path.join(dist, 'assets'), { recursive: true });

console.log('Build selesai -> dist/');
