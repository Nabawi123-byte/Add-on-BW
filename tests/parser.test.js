import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseCommand, parseChain, prefersAI } from '../src/core/parser.js';

const X = (t) => parseCommand(t, 'excel');
const W = (t) => parseCommand(t, 'word');

test('excel: template keuangan + parameter', () => {
  let r = X('tolong buatkan simulasi cicilan pinjaman 50 juta bunga 10% 24 bulan');
  assert.equal(r.action, 'excel.template');
  assert.equal(r.params.id, 'cicilan');
  assert.equal(r.params.pokok, 50000000);
  assert.equal(r.params.bunga, 0.1);
  assert.equal(r.params.tenor, 24);
  r = X('cicilan kpr 300 juta bunga 7,5% 15 tahun flat');
  assert.equal(r.params.id, 'cicilan');
  assert.equal(r.params.tenor, 180);
  assert.equal(r.params.metode, 'flat');
  assert.equal(r.params.bunga, 0.075);
  assert.equal(X('buat anggaran bulanan').params.id, 'anggaran');
  assert.equal(X('bikin laporan arus kas saldo awal 5 juta').params.saldo, 5000000);
  assert.equal(X('buat laporan laba rugi').params.id, 'laba-rugi');
  assert.equal(X('buat daftar gaji 15 karyawan').params.baris, 15);
  assert.equal(X('buat absensi bulan desember').params.bulan, 11);
  assert.equal(X('buat invoice').params.id, 'faktur');
  assert.equal(X('buat daftar nilai kelas 7A kkm 80').params.kkm, 80);
});

test('excel: tabel', () => {
  const r = X('buat tabel karyawan dengan kolom Nama, Jabatan, Gaji 12 baris');
  assert.equal(r.action, 'excel.table');
  assert.deepEqual(r.params.headers, ['Nama', 'Jabatan', 'Gaji']);
  assert.equal(r.params.rows, 12);
  assert.equal(r.params.name, 'karyawan');
  assert.deepEqual(X('buat tabel 5x10').params, { cols: 5, rows: 10, headers: null });
  assert.equal(X('jadikan tabel').action, 'excel.toTable');
});

test('excel: operasi', () => {
  assert.equal(X('jumlahkan').params.fn, 'SUM');
  assert.equal(X('jumlahkan total penjualan').action, 'excel.aggregate');
  assert.equal(X('rata-rata').params.fn, 'AVERAGE');
  const s = X('urutkan berdasarkan gaji turun');
  assert.equal(s.action, 'excel.sort');
  assert.equal(s.params.ascending, false);
  assert.equal(s.params.column, 'gaji');
  const h = X('tandai nilai di atas 1 juta merah');
  assert.equal(h.action, 'excel.highlight');
  assert.equal(h.params.rule, 'greater');
  assert.equal(h.params.value, 1000000);
  assert.equal(X('tandai duplikat kuning').params.rule, 'duplicate');
  assert.equal(X('grafik lingkaran').params.type, 'Pie');
  assert.equal(X('format rupiah').params.kind, 'rupiah');
  assert.equal(X('hapus duplikat').action, 'excel.dedupe');
  assert.equal(X('terbilang 1.250.000').params.number, 1250000);
  assert.equal(X('sheet baru Laporan 2026').params.name, 'Laporan 2026');
  assert.equal(X('rekap total per Produk').action, 'excel.groupSum');
  assert.equal(X('rekap total per Produk').params.column, 'produk');
  assert.equal(X('bekukan header').action, 'excel.freeze');
  assert.equal(X('huruf kapital tiap kata').params.mode, 'title');
  assert.equal(X('apa kabar dunia'), null);
});

test('word: surat & edit', () => {
  const r = W('buat surat izin sakit nama: Budi, alasan: demam');
  assert.equal(r.action, 'word.template');
  assert.equal(r.params.id, 'izin');
  assert.equal(r.params.given.alasan, 'demam');
  assert.equal(W('surat lamaran kerja').params.id, 'lamaran');
  assert.equal(W('buat notulen rapat').params.id, 'notulen');
  assert.equal(W('buat kwitansi').params.id, 'kwitansi');
  assert.equal(W('surat peringatan sp1').params.id, 'peringatan');
  const g = W('ganti "PT Lama" dengan "PT Baru"');
  assert.deepEqual(g.params, { find: 'PT Lama', replace: 'PT Baru' });
  const f = W('font Times New Roman ukuran 12 semua');
  assert.deepEqual(f.params, { name: 'Times New Roman', size: 12, scope: 'document' });
  assert.equal(W('spasi 1,5 semua').params.value, 1.5);
  assert.equal(W('rata kiri kanan').params.align, 'Justified');
  assert.equal(W('buat tabel 4x3').action, 'word.table');
  assert.equal(W('terbilang 2.500.000 rupiah').params.number, 2500000);
  assert.equal(W('header PT Maju Jaya').params.text, 'PT Maju Jaya');
  assert.equal(W('judul Laporan Kegiatan').params.text, 'Laporan Kegiatan');
  assert.equal(W('daftar isi').action, 'word.toc');
  assert.equal(W('sorot kata penting kuning').params.find, 'penting');
});

test('rantai perintah & preferensi AI', () => {
  const c = parseChain('buat anggaran bulanan lalu bekukan header', 'excel');
  assert.equal(c.length, 2);
  assert.equal(c[1].action, 'excel.freeze');
  assert.equal(parseChain('analisis data ini lalu jelaskan', 'excel'), null);
  assert.equal(prefersAI('analisis data penjualan ini dan beri insight'), true);
  assert.equal(prefersAI('jumlahkan'), false);
});
