// Uji asap end-to-end: panel Master di Chromium dengan Office.js tiruan dan API Claude tiruan.
// Jalankan: npm run test:e2e   (butuh Playwright + Chromium)

import http from 'node:http';
import { readFileSync, existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');
const here = path.dirname(fileURLToPath(import.meta.url));
const dist = path.resolve(here, '..', '..', 'dist');
const FAKE = readFileSync(path.join(here, 'fake-office.js'), 'utf8');

const server = http.createServer((req, res) => {
  const p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/office-js/office.js') {
    res.writeHead(200, { 'Content-Type': 'application/javascript' });
    return res.end(FAKE);
  }
  const f = path.join(dist, p === '/' ? 'taskpane.html' : p);
  if (!f.startsWith(dist) || !existsSync(f)) {
    res.writeHead(404);
    return res.end();
  }
  const type = { '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css', '.png': 'image/png', '.json': 'application/json' }[path.extname(f)] || 'text/plain';
  res.writeHead(200, { 'Content-Type': type });
  res.end(readFileSync(f));
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const BASE = `http://127.0.0.1:${server.address().port}`;

const failures = [];
const check = (cond, msg) => {
  if (cond) console.log('  ✓', msg);
  else {
    console.log('  ✗', msg);
    failures.push(msg);
  }
};

const browser = await chromium.launch();

async function openPane(host, { apiKey = '', route } = {}) {
  const ctx = await browser.newContext({ viewport: { width: 360, height: 800 } });
  const page = await ctx.newPage();
  const pageErrors = [];
  page.on('pageerror', (e) => pageErrors.push(String(e)));
  page.on('console', (m) => m.type() === 'error' && !/Failed to load resource/.test(m.text()) && pageErrors.push(m.text()));
  if (route) await page.route('https://api.anthropic.com/**', route);
  await page.addInitScript((key) => {
    localStorage.setItem('master.settings', JSON.stringify({ apiKey: key, mode: 'auto', model: 'claude-opus-5-5', effort: 'medium', codePolicy: 'ask' }));
    localStorage.setItem('master.profile', JSON.stringify({ nama: 'Budi Santoso', jabatan: 'Manajer', instansi: 'PT Maju Jaya', kota: 'Bandung' }));
  }, apiKey);
  await page.goto(`${BASE}/taskpane.html?fakeHost=${host}`);
  await page.waitForSelector('.msg.bot');
  return { page, ctx, pageErrors };
}

async function send(page, text) {
  const before = await page.locator('.msg').count();
  await page.fill('#input', text);
  await page.press('#input', 'Enter');
  await page.waitForFunction((n) => document.querySelectorAll('.msg').length > n + 1 && !document.body.classList.contains('busy'), before, { timeout: 15000 });
  await page.waitForTimeout(50);
  const msgs = await page.locator('.msg').allInnerTexts();
  return msgs.slice(before + 1).join('\n');
}

/* ------------------------------ EXCEL ------------------------------ */
console.log('\nExcel (offline, tanpa API key)');
{
  const { page, ctx, pageErrors } = await openPane('Excel');
  check((await page.textContent('#hostBadge')) === 'Excel', 'host terdeteksi Excel');
  check(/tanpa API key/.test(await page.textContent('#netBadge')), 'status offline tanpa key');

  const templateCmds = [
    'buat anggaran bulanan',
    'buat laporan arus kas saldo awal 5 juta',
    'buat buku kas kecil',
    'buat laporan laba rugi',
    'buat neraca',
    'buat daftar gaji 5 karyawan',
    'buat faktur',
    'simulasi cicilan pinjaman 50 juta bunga 10% 24 bulan',
    'cicilan kpr 300 juta bunga 7,5% 2 tahun flat',
    'buat daftar stok barang',
    'buat absensi bulan februari 2028',
    'buat daftar nilai kelas 7A',
    'buat catatan keuangan harian',
    'buat catatan hutang piutang',
    'buat laporan penjualan',
    'buat daftar tugas proyek',
    'buat target tabungan 20 juta 10 bulan',
  ];
  for (const cmd of templateCmds) {
    const out = await send(page, cmd);
    check(/✅/.test(out) && !/gagal|❌/.test(out), `${cmd} → ${out.split('\n')[0].slice(0, 70)}`);
    if (/gagal|❌/.test(out)) console.log(out);
  }
  const sheets = await page.evaluate(() => window.__workbook.sheets.map((s) => s.name).filter((n) => n !== '_MasterList'));
  check(sheets.length === 18, `17 template + Sheet1 = 18 sheet (${sheets.length})`);
  const pmt = await page.evaluate(() => {
    const s = window.__workbook.sheets.find((x) => x.name === 'Cicilan');
    return s.cell(6, 1).f;
  });
  check(pmt === '=ROUND(PMT($B$4/12,$B$5,-$B$3),0)', `rumus angsuran: ${pmt}`);
  const terb = await page.evaluate(() => {
    const s = window.__workbook.sheets.find((x) => x.name === 'Faktur');
    for (const [, v] of s.cells) if (typeof v.f === 'string' && /rupiah$/.test(v.f)) return v.f;
    return null;
  });
  check(!!terb, `faktur berisi terbilang: ${terb}`);

  // Data contoh untuk perintah olah data
  await page.evaluate(() => {
    const wb = window.__workbook;
    const s = wb.worksheets.add('Data');
    wb.active = s;
    const rows = [
      ['Produk', 'Kota', 'Jumlah'],
      ['Teh', 'Bandung', 1500000],
      ['Kopi', 'Jakarta', 2500000],
      ['Teh', 'Jakarta', 500000],
      ['Gula', 'Bandung', 750000],
    ];
    s.getRange('A1:C5').formulas = rows;
    wb.selection = s.getRange('B3');
  });
  for (const cmd of ['jumlahkan', 'urutkan berdasarkan jumlah turun', 'format rupiah', 'grafik batang', 'tandai nilai di atas 1 juta merah', 'persentase dari total', 'ranking', 'rekap total per Produk', 'filter', 'bekukan header', 'beri garis', 'huruf besar', 'hapus spasi', 'nomor urut']) {
    const out = await send(page, cmd);
    check(/✅|ℹ️/.test(out) && !/❌|gagal/.test(out), `${cmd} → ${out.split('\n')[0].slice(0, 80)}`);
  }
  const recap = await page.evaluate(() => {
    const s = window.__workbook.sheets.find((x) => x.name === 'Data');
    const out = [];
    for (const [k, v] of s.cells) if (typeof v.f === 'string' && v.f.startsWith('=SUMIF(')) out.push(k + ' ' + v.f);
    return out;
  });
  check(recap.length === 3, `rekap SUMIF per produk (tanpa baris TOTAL): ${recap.length} kelompok`);

  await page.evaluate(() => {
    window.__workbook.selection = window.__workbook.active.getRange('E10');
  });
  let out = await send(page, 'terbilang 1.250.000');
  check(/Satu juta dua ratus lima puluh ribu rupiah/.test(out), 'terbilang angka langsung');
  out = await send(page, 'buat tabel karyawan dengan kolom Nama, Jabatan, Gaji 5 baris');
  check(/Tabel 4 kolom × 5 baris/.test(out), `tabel bebas: ${out.slice(0, 60)}`);
  out = await send(page, 'buat tabel 3x4 lalu rapikan kolom');
  check((out.match(/✅/g) || []).length === 2, 'rantai perintah (2 langkah)');
  out = await send(page, 'analisis tren penjualan ini');
  check(/butuh AI|API key/.test(out), 'perintah AI tanpa key diarahkan ke Pengaturan');
  out = await send(page, 'bantuan');
  check((await page.locator('.help-title').count()) >= 5, 'bantuan menampilkan kategori perintah');

  // Tab Template: form cicilan
  await page.click('.tab[data-tab=templates]');
  check((await page.locator('#templateList .card').count()) === 16, '16 kartu template Excel');
  await page.locator('#templateList .card', { hasText: 'Simulasi Cicilan' }).click();
  await page.fill('#templateForm [data-key=pokok]', '12.000.000');
  await page.fill('#templateForm [data-key=tenor]', '6');
  await page.click('#templateForm button[type=submit]');
  await page.waitForTimeout(300);
  const lastMsg = (await page.locator('.msg').allInnerTexts()).pop();
  check(/Simulasi Cicilan dibuat/.test(lastMsg), `form template cicilan: ${lastMsg.slice(0, 60)}`);

  // Makro
  await page.click('.tab[data-tab=macros]');
  await page.fill('#macroName', 'rapikan');
  await page.fill('#macroCommands', 'rapikan kolom\nbekukan header');
  await page.click('#saveMacro');
  await page.click('.tab[data-tab=chat]');
  out = await send(page, 'jalankan makro rapikan');
  check(/Menjalankan makro/.test(out), 'makro berjalan');

  check(pageErrors.length === 0, `tanpa error JS (${pageErrors.slice(0, 3).join(' | ')})`);
  await ctx.close();
}

/* ------------------------------ AI (tiruan) ------------------------------ */
console.log('\nExcel + AI (API Claude tiruan)');
{
  const requests = [];
  const sse = (events) => events.map((e) => `event: ${e.type}\ndata: ${JSON.stringify(e)}\n\n`).join('');
  const start = (id) => ({ type: 'message_start', message: { id, type: 'message', role: 'assistant', model: 'claude-opus-5-5', content: [], stop_reason: null, stop_sequence: null, usage: { input_tokens: 10, output_tokens: 1 } } });
  const route = async (r) => {
    const body = JSON.parse(r.request().postData());
    requests.push({ body, headers: r.request().headers() });
    let events;
    if (requests.length === 1) {
      const input = JSON.stringify({ ops: [{ op: 'write', address: 'A1', values: [['Halo dari AI', '=1+1']] }, { op: 'format', address: 'A1', bold: true }] });
      events = [
        start('msg_1'),
        { type: 'content_block_start', index: 0, content_block: { type: 'text', text: '' } },
        { type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: 'Saya tulis datanya.' } },
        { type: 'content_block_stop', index: 0 },
        { type: 'content_block_start', index: 1, content_block: { type: 'tool_use', id: 'toolu_1', name: 'excel_apply', input: {} } },
        { type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: input.slice(0, 30) } },
        { type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: input.slice(30) } },
        { type: 'content_block_stop', index: 1 },
        { type: 'message_delta', delta: { stop_reason: 'tool_use', stop_sequence: null }, usage: { output_tokens: 30 } },
        { type: 'message_stop' },
      ];
    } else {
      events = [
        start('msg_2'),
        { type: 'content_block_start', index: 0, content_block: { type: 'text', text: '' } },
        { type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: 'Selesai: A1 berisi sapaan.' } },
        { type: 'content_block_stop', index: 0 },
        { type: 'message_delta', delta: { stop_reason: 'end_turn', stop_sequence: null }, usage: { output_tokens: 8 } },
        { type: 'message_stop' },
      ];
    }
    await r.fulfill({ status: 200, headers: { 'content-type': 'text/event-stream', 'access-control-allow-origin': '*' }, body: sse(events) });
  };
  const { page, ctx, pageErrors } = await openPane('Excel', { apiKey: 'sk-ant-test', route });
  check(/AI Online/.test(await page.textContent('#netBadge')), 'status AI online');
  const out = await send(page, 'analisis data ini dan tulis sapaan di A1');
  check(/Selesai: A1 berisi sapaan/.test(out), `jawaban AI: ${out.replace(/\n/g, ' ').slice(0, 90)}`);
  const a1 = await page.evaluate(() => window.__workbook.active.cell(0, 0).f);
  check(a1 === 'Halo dari AI', 'alat excel_apply menulis ke sheet');
  const b = requests[0] && requests[0].body;
  check(b && b.model === 'claude-opus-5-5', 'model default claude-opus-5-5');
  check(b && b.fallbacks === 'default' && /server-side-fallback-2026-07-01/.test(requests[0].headers['anthropic-beta'] || ''), 'fallback server-side aktif');
  check(b && b.output_config && b.output_config.effort === 'medium', 'effort medium dikirim');
  check(b && b.tools.every((t) => t.eager_input_streaming === true), 'eager_input_streaming di semua alat');
  check(b && b.stream === true, 'permintaan streaming');
  check(requests.length === 2 && requests[1].body.messages.length === 3 && requests[1].body.messages[2].content[0].type === 'tool_result', 'hasil alat dikirim balik ke Claude');
  check(!JSON.stringify(requests[1].body.messages[2]).includes('is_error'), 'alat sukses tanpa error');
  check(pageErrors.length === 0, `tanpa error JS (${pageErrors.slice(0, 3).join(' | ')})`);
  await ctx.close();
}

console.log('\nExcel + AI saat koneksi putus');
{
  const { page, ctx } = await openPane('Excel', { apiKey: 'sk-ant-test', route: (r) => r.abort('internetdisconnected') });
  const out = await send(page, 'analisis data penjualan ini');
  check(/antrean/.test(out), `perintah masuk antrean: ${out.slice(0, 60)}`);
  check(await page.isVisible('#queueBar'), 'bar antrean tampil');
  const out2 = await send(page, 'jumlahkan');
  check(!/antrean/.test(out2), 'perintah offline tetap jalan saat koneksi putus');
  await ctx.close();
}

/* ------------------------------ WORD ------------------------------ */
console.log('\nWord (offline)');
{
  const { page, ctx, pageErrors } = await openPane('Word');
  check((await page.textContent('#hostBadge')) === 'Word', 'host terdeteksi Word');
  await page.fill('#input', 'buat surat izin sakit alasan: demam tinggi');
  await page.press('#input', 'Enter');
  await page.waitForSelector('#log form.tform');
  check((await page.inputValue('#log form [data-key=nama]')) === 'Budi Santoso', 'form surat terisi dari profil');
  check((await page.inputValue('#log form [data-key=alasan]')) === 'demam tinggi', 'form surat terisi dari perintah');
  await page.click('#log form button[type=submit]');
  await page.waitForTimeout(300);
  let html = await page.evaluate(() => window.__doc.html.join(''));
  check(/demam tinggi/.test(html) && /Budi Santoso/.test(html) && /Bandung/.test(html), 'surat izin disisipkan ke dokumen');

  for (const [cmd, re] of [
    ['buat tabel kolom No, Nama, Alamat 3 baris', /Tabel 3 kolom × 3 baris/],
    ['terbilang 2.500.000 rupiah', /Dua juta lima ratus ribu rupiah/],
    ['tanggal hari ini', /✅/],
    ['nomor halaman', /Nomor halaman/],
    ['daftar isi', /Daftar isi/],
    ['header PT Maju Jaya', /Header/],
    ['halaman baru', /Halaman baru/],
    ['blok tanda tangan', /tanda tangan/],
    ['judul Laporan Kegiatan', /Judul/],
    ['hitung kata', /kata/],
  ]) {
    const out = await send(page, cmd);
    check(re.test(out) && !/❌/.test(out), `${cmd} → ${out.split('\n')[0].slice(0, 70)}`);
  }

  await page.click('.tab[data-tab=templates]');
  check((await page.locator('#templateList .card').count()) === 18, '18 kartu template Word');
  await page.locator('#templateList .card', { hasText: 'Kwitansi' }).click();
  await page.fill('#templateForm [data-key=jumlah]', '750.000');
  await page.click('#templateForm button[type=submit]');
  await page.waitForTimeout(300);
  html = await page.evaluate(() => window.__doc.html.join(''));
  check(/Tujuh ratus lima puluh ribu rupiah/.test(html), 'kwitansi dengan terbilang otomatis');
  check(pageErrors.length === 0, `tanpa error JS (${pageErrors.slice(0, 3).join(' | ')})`);
  await page.screenshot({ path: path.join(here, 'screenshot-word.png') });
  await ctx.close();
}

await browser.close();
server.close();
console.log(failures.length ? `\n${failures.length} pemeriksaan GAGAL` : '\nSemua pemeriksaan lulus');
process.exit(failures.length ? 1 : 0);
