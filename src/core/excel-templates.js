// Template Excel siap pakai (berjalan 100% offline).
// Setiap template mengembalikan SheetPlan berisi operasi yang dijalankan executor.

import { SheetPlan, RP } from './plan.js';
import { BULAN, HARI, FORMAT_TANGGAL, FORMAT_PERSEN, FORMAT_ANGKA, columnLetter, tanggalIndonesia, clamp } from './util.js';

const nowLabel = () => {
  const d = new Date();
  return `${BULAN[d.getMonth()]} ${d.getFullYear()}`;
};

// Baris kosong tetap bersih: rumus hanya dihitung bila kolom kunci diisi.
const guard = (keyCell, formula) => `=IF(${keyCell}="","",${formula})`;

function anggaran(p) {
  const plan = new SheetPlan('Anggaran');
  plan.title('ANGGARAN BULANAN', p.periode || nowLabel(), 7);
  const ref = plan.inputs(3, [
    { key: 'masuk', label: 'Total Pemasukan', value: p.pemasukan || 10000000, format: RP },
  ]);
  const kategori = [
    ['Makan & Minum', 3000000],
    ['Transportasi', 800000],
    ['Tempat Tinggal / Sewa', 2500000],
    ['Listrik, Air & Internet', 700000],
    ['Pendidikan', 500000],
    ['Kesehatan', 300000],
    ['Tabungan & Investasi', 1000000],
    ['Hiburan', 400000],
    ['Cicilan', 0],
    ['Lain-lain', 300000],
  ];
  const H = ['No', 'Kategori', 'Rencana', 'Realisasi', 'Selisih', '% Terpakai', 'Status'];
  const t = plan.table({
    startRow: 9,
    headers: H,
    count: kategori.length + 2,
    row: (r, i, C) => [
      i + 1,
      kategori[i] ? kategori[i][0] : '',
      kategori[i] ? kategori[i][1] : '',
      '',
      guard(`${C('Kategori')}${r}`, `${C('Rencana')}${r}-${C('Realisasi')}${r}`),
      guard(`${C('Kategori')}${r}`, `IF(${C('Rencana')}${r}=0,0,${C('Realisasi')}${r}/${C('Rencana')}${r})`),
      guard(`${C('Kategori')}${r}`, `IF(${C('Realisasi')}${r}>${C('Rencana')}${r},"MELEBIHI","AMAN")`),
    ],
    formats: { Rencana: RP, Realisasi: RP, Selisih: RP, '% Terpakai': FORMAT_PERSEN },
    totals: { Rencana: 'SUM', Realisasi: 'SUM', Selisih: 'SUM' },
    name: 'Anggaran',
  });
  plan.inputs(4, [
    { key: 'rencana', label: 'Total Rencana Pengeluaran', value: `=${t.total('Rencana')}`, format: RP, input: false },
    { key: 'real', label: 'Total Realisasi', value: `=${t.total('Realisasi')}`, format: RP, input: false },
    { key: 'sisa', label: 'Sisa Dana', value: `=${ref.masuk}-${t.total('Realisasi')}`, format: RP, input: false },
  ]);
  plan.push({ op: 'conditional', address: t.col('Status'), rule: 'text', value: 'MELEBIHI', fill: '#F8CBAD', fontColor: '#9C0006' });
  plan.push({ op: 'conditional', address: t.col('% Terpakai'), rule: 'dataBar' });
  plan.push({ op: 'chart', address: `B${t.headerRow}:D${t.last}`, type: 'ColumnClustered', title: 'Rencana vs Realisasi', position: 'I3' });
  plan.summary = 'Isi kolom Realisasi; selisih, persentase, status dan sisa dana terhitung otomatis.';
  return plan.finish({ freezeRows: t.headerRow });
}

function bukuKas(p, { sheetName, judul, labelSaldo, masuk, keluar, extra }) {
  const plan = new SheetPlan(sheetName);
  plan.title(judul, p.periode || nowLabel(), 7);
  const ref = plan.inputs(3, [{ key: 'awal', label: labelSaldo, value: p.saldo || 0, format: RP }]);
  const H = ['Tanggal', extra, 'Keterangan', masuk, keluar, 'Saldo'];
  const n = clamp(p.baris || 30, 5, 1000);
  const start = 8;
  const first = start + 1;
  const t = plan.table({
    startRow: start,
    headers: H,
    count: n,
    row: (r, i, C) => [
      '',
      '',
      '',
      '',
      '',
      `=${ref.awal}+SUM(${C(masuk)}$${first}:${C(masuk)}${r})-SUM(${C(keluar)}$${first}:${C(keluar)}${r})`,
    ],
    formats: { Tanggal: FORMAT_TANGGAL, [masuk]: RP, [keluar]: RP, Saldo: RP },
    totals: { [masuk]: 'SUM', [keluar]: 'SUM' },
    name: sheetName.replace(/\s+/g, ''),
  });
  plan.inputs(4, [
    { key: 'tm', label: `Total ${masuk}`, value: `=${t.total(masuk)}`, format: RP, input: false },
    { key: 'tk', label: `Total ${keluar}`, value: `=${t.total(keluar)}`, format: RP, input: false },
    { key: 'akhir', label: 'Saldo Akhir', value: `=${ref.awal}+${t.total(masuk)}-${t.total(keluar)}`, format: RP, input: false },
  ]);
  if (extra === 'Kategori') {
    plan.push({ op: 'validation', address: t.col('Kategori'), list: ['Penjualan', 'Pendapatan Lain', 'Gaji', 'Operasional', 'Pembelian', 'Pajak', 'Lain-lain'] });
  }
  plan.push({ op: 'conditional', address: t.col('Saldo'), rule: 'less', value: 0, fill: '#F8CBAD', fontColor: '#9C0006' });
  plan.summary = 'Isi transaksi per baris; saldo berjalan dan saldo akhir terhitung otomatis.';
  return plan.finish({ freezeRows: start });
}

const arusKas = (p) =>
  bukuKas(p, { sheetName: 'Arus Kas', judul: 'LAPORAN ARUS KAS', labelSaldo: 'Saldo Awal', masuk: 'Kas Masuk', keluar: 'Kas Keluar', extra: 'Kategori' });

const kasKecil = (p) =>
  bukuKas(p, { sheetName: 'Kas Kecil', judul: 'BUKU KAS KECIL (PETTY CASH)', labelSaldo: 'Dana Kas Kecil', masuk: 'Debit', keluar: 'Kredit', extra: 'No. Bukti' });

/** Laporan bertingkat (laba rugi, neraca): baris label-nilai di kolom A-B. */
function statement(plan, startRow, sections) {
  let r = startRow;
  const cells = {};
  for (const s of sections) {
    if (s.header) {
      plan.write(`A${r}`, [[s.header]]);
      plan.format(`A${r}:B${r}`, { bold: true, fill: '#1F3864', fontColor: '#FFFFFF' });
      r++;
      continue;
    }
    if (s.blank) {
      r++;
      continue;
    }
    const value = typeof s.value === 'function' ? s.value(cells) : s.value;
    plan.write(`A${r}:B${r}`, [[s.indent ? '    ' + s.label : s.label, value === undefined ? '' : value]]);
    plan.format(`B${r}`, { numberFormat: s.format || RP });
    if (s.input) plan.format(`B${r}`, { fill: '#FFF2CC' });
    if (s.total) plan.format(`A${r}:B${r}`, { bold: true, border: 'top', fill: s.total === 'grand' ? '#D9E1F2' : undefined });
    if (s.key) cells[s.key] = `B${r}`;
    if (s.group) (cells[s.group] = cells[s.group] || []).push(`B${r}`);
    r++;
  }
  plan.format(`A${startRow}:A${r}`, { columnWidth: 38 });
  plan.format(`B${startRow}:B${r}`, { columnWidth: 20 });
  return { cells, lastRow: r - 1 };
}

const sum = (list) => (list && list.length ? `=SUM(${list.join(',')})` : '=0');

function labaRugi(p) {
  const plan = new SheetPlan('Laba Rugi');
  plan.title('LAPORAN LABA RUGI', `Periode ${p.periode || nowLabel()}`, 2);
  statement(plan, 4, [
    { header: 'PENDAPATAN' },
    { label: 'Penjualan', value: 50000000, input: true, indent: true, group: 'pdpt' },
    { label: 'Pendapatan Lain-lain', value: 0, input: true, indent: true, group: 'pdpt' },
    { label: 'Total Pendapatan', value: (c) => sum(c.pdpt), total: true, key: 'tp' },
    { label: 'Harga Pokok Penjualan (HPP)', value: 30000000, input: true, key: 'hpp' },
    { label: 'LABA KOTOR', value: (c) => `=${c.tp}-${c.hpp}`, total: true, key: 'lk' },
    { blank: true },
    { header: 'BEBAN OPERASIONAL' },
    ...['Gaji Karyawan', 'Sewa Tempat', 'Listrik, Air & Internet', 'Transportasi', 'Pemasaran', 'Penyusutan', 'Beban Lain-lain'].map((label) => ({
      label,
      value: 0,
      input: true,
      indent: true,
      group: 'beban',
    })),
    { label: 'Total Beban Operasional', value: (c) => sum(c.beban), total: true, key: 'tb' },
    { label: 'LABA OPERASIONAL', value: (c) => `=${c.lk}-${c.tb}`, total: true, key: 'lo' },
    { label: 'Tarif Pajak Penghasilan', value: p.pajak !== undefined ? p.pajak : 0.005, input: true, format: '0.0%', key: 'tarif' },
    { label: 'Pajak Penghasilan', value: (c) => `=MAX(0,${c.tp})*${c.tarif}`, key: 'pph' },
    { label: 'LABA BERSIH', value: (c) => `=${c.lo}-${c.pph}`, total: 'grand', key: 'lb' },
    { label: 'Margin Laba Bersih', value: (c) => `=IF(${c.tp}=0,0,${c.lb}/${c.tp})`, format: FORMAT_PERSEN },
  ]);
  plan.summary = 'Isi sel kuning. Tarif pajak default 0,5% dari omzet (PPh final UMKM); ubah sesuai kondisi usaha Anda.';
  return plan.finish({ autofit: false });
}

function neraca(p) {
  const plan = new SheetPlan('Neraca');
  plan.title('NERACA', `Per ${p.tanggal || tanggalIndonesia()}`, 2);
  const { cells, lastRow } = statement(plan, 4, [
    { header: 'ASET' },
    ...['Kas & Bank', 'Piutang Usaha', 'Persediaan', 'Biaya Dibayar di Muka'].map((label) => ({ label, value: 0, input: true, indent: true, group: 'al' })),
    { label: 'Total Aset Lancar', value: (c) => sum(c.al), total: true, key: 'tal' },
    ...['Tanah & Bangunan', 'Kendaraan', 'Peralatan'].map((label) => ({ label, value: 0, input: true, indent: true, group: 'at' })),
    { label: 'Akumulasi Penyusutan (isi negatif)', value: 0, input: true, indent: true, group: 'at' },
    { label: 'Total Aset Tetap', value: (c) => sum(c.at), total: true, key: 'tat' },
    { label: 'TOTAL ASET', value: (c) => `=${c.tal}+${c.tat}`, total: 'grand', key: 'ta' },
    { blank: true },
    { header: 'KEWAJIBAN & EKUITAS' },
    ...['Utang Usaha', 'Utang Pajak', 'Utang Bank Jangka Pendek'].map((label) => ({ label, value: 0, input: true, indent: true, group: 'kl' })),
    { label: 'Total Kewajiban Lancar', value: (c) => sum(c.kl), total: true, key: 'tkl' },
    { label: 'Utang Jangka Panjang', value: 0, input: true, indent: true, key: 'kjp' },
    { label: 'Total Kewajiban', value: (c) => `=${c.tkl}+${c.kjp}`, total: true, key: 'tk' },
    ...['Modal Pemilik', 'Laba Ditahan', 'Laba Tahun Berjalan'].map((label) => ({ label, value: 0, input: true, indent: true, group: 'eq' })),
    { label: 'Total Ekuitas', value: (c) => sum(c.eq), total: true, key: 'te' },
    { label: 'TOTAL KEWAJIBAN & EKUITAS', value: (c) => `=${c.tk}+${c.te}`, total: 'grand', key: 'tke' },
  ]);
  const check = lastRow + 2;
  plan.write(`A${check}:B${check}`, [['Pemeriksaan', `=IF(ABS(${cells.ta}-${cells.tke})<1,"SEIMBANG","TIDAK SEIMBANG")`]]);
  plan.format(`A${check}:B${check}`, { bold: true });
  plan.push({ op: 'conditional', address: `B${check}`, rule: 'text', value: 'TIDAK', fill: '#F8CBAD', fontColor: '#9C0006' });
  plan.push({ op: 'conditional', address: `B${check}`, rule: 'text', value: 'SEIMBANG', fill: '#C6EFCE', fontColor: '#006100' });
  plan.summary = 'Isi sel kuning. Baris "Pemeriksaan" menunjukkan apakah neraca sudah seimbang.';
  return plan.finish({ autofit: false });
}

function gaji(p) {
  const plan = new SheetPlan('Daftar Gaji');
  plan.title('DAFTAR GAJI KARYAWAN', p.periode || nowLabel(), 10);
  const ref = plan.inputs(3, [
    { key: 'bpjs', label: 'Potongan BPJS Karyawan (%)', value: p.bpjs !== undefined ? p.bpjs : 0.03, format: '0.0%' },
  ]);
  const H = ['No', 'Nama', 'Jabatan', 'Gaji Pokok', 'Tunjangan', 'Lembur', 'Gaji Kotor', 'Potongan BPJS', 'Potongan Lain', 'Gaji Bersih'];
  const contoh = [
    ['Contoh Karyawan 1', 'Staf', 5000000, 500000, 0],
    ['Contoh Karyawan 2', 'Supervisor', 7000000, 1000000, 250000],
  ];
  const t = plan.table({
    startRow: 5,
    headers: H,
    count: clamp(p.baris || 10, 1, 1000),
    row: (r, i, C) => {
      const c = contoh[i] || ['', '', '', '', ''];
      return [
        i + 1,
        c[0],
        c[1],
        c[2],
        c[3],
        c[4],
        guard(`${C('Nama')}${r}`, `${C('Gaji Pokok')}${r}+${C('Tunjangan')}${r}+${C('Lembur')}${r}`),
        guard(`${C('Nama')}${r}`, `ROUND(${C('Gaji Pokok')}${r}*${ref.bpjs},0)`),
        '',
        guard(`${C('Nama')}${r}`, `${C('Gaji Kotor')}${r}-${C('Potongan BPJS')}${r}-${C('Potongan Lain')}${r}`),
      ];
    },
    formats: Object.fromEntries(['Gaji Pokok', 'Tunjangan', 'Lembur', 'Gaji Kotor', 'Potongan BPJS', 'Potongan Lain', 'Gaji Bersih'].map((h) => [h, RP])),
    totals: Object.fromEntries(['Gaji Pokok', 'Tunjangan', 'Lembur', 'Gaji Kotor', 'Potongan BPJS', 'Potongan Lain', 'Gaji Bersih'].map((h) => [h, 'SUM'])),
    name: 'DaftarGaji',
  });
  plan.summary = 'Baris contoh boleh diganti. Gaji kotor, potongan BPJS, dan gaji bersih terhitung otomatis.';
  return plan.finish({ freezeRows: t.headerRow });
}

function faktur(p, profile = {}) {
  const plan = new SheetPlan('Faktur');
  plan.title('FAKTUR / INVOICE', null, 6);
  const d = new Date();
  const no = `INV-${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}-001`;
  plan.write('A3:B5', [
    ['Dari', profile.instansi || 'Nama Usaha Anda'],
    ['Kepada', p.kepada || 'Nama Pelanggan'],
    ['Alamat', p.alamat || ''],
  ]);
  plan.write('E3:F4', [
    ['No. Faktur', no],
    ['Tanggal', tanggalIndonesia(d)],
  ]);
  plan.format('A3:A5', { bold: true });
  plan.format('E3:E4', { bold: true });
  plan.format('B3:B5', { fill: '#FFF2CC' });
  plan.format('F3:F4', { fill: '#FFF2CC' });
  const H = ['No', 'Deskripsi', 'Qty', 'Satuan', 'Harga Satuan', 'Jumlah'];
  const t = plan.table({
    startRow: 7,
    headers: H,
    count: clamp(p.baris || 10, 1, 200),
    row: (r, i, C) => [i + 1, i === 0 ? 'Contoh barang / jasa' : '', i === 0 ? 1 : '', i === 0 ? 'unit' : '', i === 0 ? 100000 : '', `=IF(${C('Qty')}${r}="","",${C('Qty')}${r}*${C('Harga Satuan')}${r})`],
    formats: { 'Harga Satuan': RP, Jumlah: RP, Qty: FORMAT_ANGKA },
    name: 'ItemFaktur',
  });
  let r = t.last + 2;
  const sub = `F${r}`;
  plan.write(`E${r}:F${r + 4}`, [
    ['Subtotal', `=SUM(${t.col('Jumlah')})`],
    ['Diskon', `=${sub}*D${r + 1}`],
    ['DPP', `=${sub}-F${r + 1}`],
    ['PPN', `=F${r + 2}*D${r + 3}`],
    ['TOTAL', `=F${r + 2}+F${r + 3}`],
  ]);
  plan.write(`C${r + 1}:D${r + 1}`, [['Diskon (%)', p.diskon || 0]]);
  plan.write(`C${r + 3}:D${r + 3}`, [['PPN (%)', p.ppn !== undefined ? p.ppn : 0.11]]);
  plan.format(`D${r + 1}`, { numberFormat: '0%', fill: '#FFF2CC' });
  plan.format(`D${r + 3}`, { numberFormat: '0%', fill: '#FFF2CC' });
  plan.format(`F${r}:F${r + 4}`, { numberFormat: RP });
  plan.format(`E${r}:E${r + 4}`, { bold: true });
  plan.format(`E${r + 4}:F${r + 4}`, { bold: true, fill: '#D9E1F2', border: 'top-double', fontSize: 12 });
  const terb = r + 6;
  plan.write(`A${terb}`, [['Terbilang:']]);
  plan.format(`A${terb}`, { bold: true });
  plan.push({ op: 'terbilang', source: `F${r + 4}`, target: `B${terb}`, rupiah: true });
  plan.format(`B${terb}:F${terb}`, { italic: true, merge: true, wrap: true });
  plan.write(`E${terb + 2}:E${terb + 6}`, [['Hormat kami,'], [''], [''], [''], [profile.nama || '(..............................)']]);
  plan.format(`E${terb + 2}:E${terb + 6}`, { hAlign: 'Center' });
  plan.format('B1', { columnWidth: 34 });
  plan.summary = 'Isi item barang/jasa. PPN default 11% (silakan ubah). Ketik "perbarui terbilang" setelah total berubah.';
  return plan.finish({ autofit: false });
}

function cicilan(p) {
  const pokok = p.pokok || 50000000;
  const bunga = p.bunga !== undefined ? p.bunga : 0.1;
  const tenor = clamp(Math.round(p.tenor || 12), 1, 360);
  const flat = p.metode === 'flat';
  const plan = new SheetPlan('Cicilan');
  plan.title('SIMULASI CICILAN / ANGSURAN PINJAMAN', flat ? 'Metode bunga flat' : 'Metode bunga anuitas', 6);
  const ref = plan.inputs(3, [
    { key: 'P', label: 'Jumlah Pinjaman', value: pokok, format: RP },
    { key: 'i', label: 'Bunga per Tahun', value: bunga, format: '0.00%' },
    { key: 'n', label: 'Tenor (bulan)', value: tenor, format: '0' },
    { key: 'm', label: 'Metode', value: flat ? 'Flat' : 'Anuitas', format: '@', input: false },
  ]);
  const angsuran = flat ? `=ROUND(${ref.P}/${ref.n}+${ref.P}*${ref.i}/12,0)` : `=ROUND(PMT(${ref.i}/12,${ref.n},-${ref.P}),0)`;
  plan.inputs(7, [{ key: 'A', label: 'Angsuran per Bulan', value: angsuran, format: RP, input: false }]);
  plan.write('A8', [['Catatan: jadwal dibuat untuk tenor di atas. Jika tenor diubah, jalankan perintah cicilan lagi.']]);
  plan.format('A8', { italic: true, fontColor: '#7F7F7F' });
  const H = ['Bulan Ke', 'Angsuran', 'Bunga', 'Pokok', 'Sisa Pinjaman'];
  const t = plan.table({
    startRow: 10,
    headers: H,
    count: tenor,
    row: (r, i, C) => {
      const prev = i === 0 ? ref.P : `${C('Sisa Pinjaman')}${r - 1}`;
      if (flat) {
        return [i + 1, '=$B$7', `=ROUND(${ref.P}*${ref.i}/12,0)`, `=${C('Angsuran')}${r}-${C('Bunga')}${r}`, `=MAX(0,${prev}-${C('Pokok')}${r})`];
      }
      const last = i === tenor - 1;
      return [
        i + 1,
        last ? `=${prev}+${C('Bunga')}${r}` : '=$B$7',
        `=ROUND(${prev}*${ref.i}/12,0)`,
        `=${C('Angsuran')}${r}-${C('Bunga')}${r}`,
        `=MAX(0,${prev}-${C('Pokok')}${r})`,
      ];
    },
    formats: { Angsuran: RP, Bunga: RP, Pokok: RP, 'Sisa Pinjaman': RP },
    totals: { Angsuran: 'SUM', Bunga: 'SUM', Pokok: 'SUM' },
    name: 'JadwalCicilan',
  });
  plan.push({ op: 'chart', address: `E${t.headerRow}:E${t.last}`, type: 'Line', title: 'Sisa Pinjaman per Bulan', position: 'H3' });
  plan.summary = `Pinjaman ${pokok.toLocaleString('id-ID')} selama ${tenor} bulan, bunga ${(bunga * 100).toFixed(2)}%/tahun (${flat ? 'flat' : 'anuitas'}).`;
  return plan.finish({ freezeRows: t.headerRow });
}

function stok(p) {
  const plan = new SheetPlan('Stok Barang');
  plan.title('DAFTAR STOK / INVENTARIS BARANG', p.periode || nowLabel(), 11);
  const H = ['Kode', 'Nama Barang', 'Satuan', 'Stok Awal', 'Masuk', 'Keluar', 'Stok Akhir', 'Stok Minimum', 'Status', 'Harga Satuan', 'Nilai Stok'];
  const t = plan.table({
    startRow: 3,
    headers: H,
    count: clamp(p.baris || 20, 1, 2000),
    row: (r, i, C) => [
      i === 0 ? 'BRG-001' : '',
      i === 0 ? 'Contoh barang' : '',
      i === 0 ? 'pcs' : '',
      i === 0 ? 50 : '',
      i === 0 ? 10 : '',
      i === 0 ? 45 : '',
      guard(`${C('Nama Barang')}${r}`, `${C('Stok Awal')}${r}+${C('Masuk')}${r}-${C('Keluar')}${r}`),
      i === 0 ? 20 : '',
      guard(`${C('Nama Barang')}${r}`, `IF(${C('Stok Akhir')}${r}<=${C('Stok Minimum')}${r},"PERLU RESTOK","AMAN")`),
      i === 0 ? 15000 : '',
      guard(`${C('Nama Barang')}${r}`, `${C('Stok Akhir')}${r}*${C('Harga Satuan')}${r}`),
    ],
    formats: { 'Harga Satuan': RP, 'Nilai Stok': RP, 'Stok Awal': FORMAT_ANGKA, 'Stok Akhir': FORMAT_ANGKA },
    totals: { 'Nilai Stok': 'SUM' },
    name: 'StokBarang',
  });
  plan.push({ op: 'conditional', address: t.col('Status'), rule: 'text', value: 'RESTOK', fill: '#F8CBAD', fontColor: '#9C0006' });
  plan.push({ op: 'conditional', address: t.col('Status'), rule: 'text', value: 'AMAN', fill: '#C6EFCE', fontColor: '#006100' });
  plan.summary = 'Stok akhir, status restok, dan nilai persediaan terhitung otomatis.';
  return plan.finish({ freezeRows: t.headerRow });
}

function absensi(p) {
  const now = new Date();
  const bulan = p.bulan !== undefined ? p.bulan : now.getMonth();
  const tahun = p.tahun || now.getFullYear();
  const days = new Date(tahun, bulan + 1, 0).getDate();
  const plan = new SheetPlan(`Absensi ${BULAN[bulan].slice(0, 3)}`);
  const H = ['No', 'Nama'];
  for (let d = 1; d <= days; d++) H.push(String(d));
  H.push('H', 'I', 'S', 'A', '% Hadir');
  plan.title(`DAFTAR HADIR ${BULAN[bulan].toUpperCase()} ${tahun}`, 'Kode: H = Hadir, I = Izin, S = Sakit, A = Alpa', H.length);
  const firstDay = columnLetter(3);
  const lastDay = columnLetter(2 + days);
  const t = plan.table({
    startRow: 4,
    headers: H,
    count: clamp(p.baris || 15, 1, 1000),
    row: (r, i, C) => {
      const row = [i + 1, ''];
      for (let d = 1; d <= days; d++) row.push('');
      const rg = `${firstDay}${r}:${lastDay}${r}`;
      row.push(`=COUNTIF(${rg},"H")`, `=COUNTIF(${rg},"I")`, `=COUNTIF(${rg},"S")`, `=COUNTIF(${rg},"A")`);
      row.push(`=IF(SUM(${C('H')}${r}:${C('A')}${r})=0,"",${C('H')}${r}/SUM(${C('H')}${r}:${C('A')}${r}))`);
      return row;
    },
    formats: { '% Hadir': FORMAT_PERSEN },
    name: 'Absensi',
    style: 'TableStyleLight9',
  });
  const area = `${firstDay}${t.first}:${lastDay}${t.last}`;
  plan.format(`${firstDay}${t.headerRow}:${lastDay}${t.last}`, { columnWidth: 3.5, hAlign: 'Center' });
  for (let d = 1; d <= days; d++) {
    const dow = new Date(tahun, bulan, d).getDay();
    if (dow === 0) plan.format(`${columnLetter(2 + d)}${t.first}:${columnLetter(2 + d)}${t.last}`, { fill: '#E7E6E6' });
  }
  plan.push({ op: 'validation', address: area, list: ['H', 'I', 'S', 'A'] });
  plan.push({ op: 'conditional', address: area, rule: 'equal', value: '"A"', fill: '#F8CBAD', fontColor: '#9C0006' });
  plan.push({ op: 'conditional', address: area, rule: 'equal', value: '"S"', fill: '#FFEB9C', fontColor: '#9C5700' });
  plan.push({ op: 'conditional', address: area, rule: 'equal', value: '"I"', fill: '#DDEBF7', fontColor: '#1F4E78' });
  plan.summary = `Kolom hari Minggu diberi warna abu-abu (${HARI[0]}). Pilih H/I/S/A dari daftar di setiap sel.`;
  return plan.finish({ freezeRows: t.headerRow, autofit: false });
}

function nilai(p) {
  const plan = new SheetPlan('Nilai Siswa');
  plan.title('DAFTAR NILAI SISWA', p.kelas ? `Kelas ${p.kelas}` : 'Mata pelajaran: ..........', 8);
  const ref = plan.inputs(3, [
    { key: 'tg', label: 'Bobot Tugas', value: 0.3, format: '0%' },
    { key: 'uts', label: 'Bobot UTS', value: 0.3, format: '0%' },
    { key: 'uas', label: 'Bobot UAS', value: 0.4, format: '0%' },
    { key: 'kkm', label: 'KKM', value: p.kkm || 75, format: '0' },
  ]);
  const H = ['No', 'Nama Siswa', 'Tugas', 'UTS', 'UAS', 'Nilai Akhir', 'Predikat', 'Keterangan'];
  const t = plan.table({
    startRow: 9,
    headers: H,
    count: clamp(p.baris || 30, 1, 1000),
    row: (r, i, C) => [
      i + 1,
      '',
      '',
      '',
      '',
      guard(`${C('Nama Siswa')}${r}`, `ROUND(${C('Tugas')}${r}*${ref.tg}+${C('UTS')}${r}*${ref.uts}+${C('UAS')}${r}*${ref.uas},2)`),
      guard(`${C('Nama Siswa')}${r}`, `IF(${C('Nilai Akhir')}${r}>=90,"A",IF(${C('Nilai Akhir')}${r}>=80,"B",IF(${C('Nilai Akhir')}${r}>=70,"C",IF(${C('Nilai Akhir')}${r}>=60,"D","E"))))`),
      guard(`${C('Nama Siswa')}${r}`, `IF(${C('Nilai Akhir')}${r}>=${ref.kkm},"TUNTAS","BELUM TUNTAS")`),
    ],
    name: 'NilaiSiswa',
  });
  const na = t.col('Nilai Akhir');
  plan.write('D3:E6', [
    ['Rata-rata Kelas', `=IFERROR(AVERAGE(${na}),0)`],
    ['Nilai Tertinggi', `=MAX(${na})`],
    ['Nilai Terendah', `=IFERROR(SMALL(${na},1),0)`],
    ['Jumlah Tuntas', `=COUNTIF(${t.col('Keterangan')},"TUNTAS")`],
  ]);
  plan.format('D3:D6', { bold: true });
  plan.format('E3:E6', { numberFormat: '0.00', fill: '#E2EFDA' });
  plan.push({ op: 'conditional', address: t.col('Keterangan'), rule: 'text', value: 'BELUM', fill: '#F8CBAD', fontColor: '#9C0006' });
  plan.summary = 'Bobot dan KKM bisa diubah di sel kuning.';
  return plan.finish({ freezeRows: t.headerRow });
}

function keuanganPribadi(p) {
  const plan = new SheetPlan('Keuangan Harian');
  plan.title('CATATAN KEUANGAN HARIAN', p.periode || nowLabel(), 10);
  const kategori = ['Gaji / Pemasukan', 'Makan & Minum', 'Transportasi', 'Belanja', 'Tagihan', 'Hiburan', 'Kesehatan', 'Pendidikan', 'Sedekah / Donasi', 'Lain-lain'];
  const H = ['Tanggal', 'Kategori', 'Keterangan', 'Pemasukan', 'Pengeluaran', 'Metode Bayar'];
  const t = plan.table({
    startRow: 3,
    headers: H,
    count: clamp(p.baris || 50, 5, 5000),
    formats: { Tanggal: FORMAT_TANGGAL, Pemasukan: RP, Pengeluaran: RP },
    totals: { Pemasukan: 'SUM', Pengeluaran: 'SUM' },
    name: 'KeuanganHarian',
  });
  plan.push({ op: 'validation', address: t.col('Kategori'), list: kategori });
  plan.push({ op: 'validation', address: t.col('Metode Bayar'), list: ['Tunai', 'Transfer', 'QRIS', 'Kartu Debit', 'Kartu Kredit', 'E-Wallet'] });
  // Ringkasan per kategori di sisi kanan
  const sr = 3;
  const rows = [['Ringkasan per Kategori', 'Jumlah']];
  kategori.forEach((k) => {
    const col = k === 'Gaji / Pemasukan' ? t.col('Pemasukan') : t.col('Pengeluaran');
    rows.push([k, `=SUMIF(${t.col('Kategori')},"${k}",${col})`]);
  });
  rows.push(['SISA UANG', `=${t.total('Pemasukan')}-${t.total('Pengeluaran')}`]);
  plan.write(`H${sr}`, rows);
  plan.format(`H${sr}:I${sr}`, { bold: true, fill: '#1F3864', fontColor: '#FFFFFF' });
  plan.format(`I${sr + 1}:I${sr + rows.length - 1}`, { numberFormat: RP });
  plan.format(`H${sr + rows.length - 1}:I${sr + rows.length - 1}`, { bold: true, fill: '#D9E1F2', border: 'top-double' });
  plan.format(`H${sr}:I${sr + rows.length - 1}`, { border: 'all' });
  plan.push({ op: 'chart', address: `H${sr + 1}:I${sr + kategori.length}`, type: 'Pie', title: 'Komposisi per Kategori', position: 'K3' });
  plan.summary = 'Catat transaksi harian; ringkasan per kategori dan grafik ikut diperbarui.';
  return plan.finish({ freezeRows: t.headerRow });
}

function hutangPiutang(p) {
  const plan = new SheetPlan('Hutang Piutang');
  plan.title('CATATAN HUTANG & PIUTANG', p.periode || nowLabel(), 9);
  const H = ['Tanggal', 'Nama', 'Jenis', 'Keterangan', 'Jumlah', 'Dibayar', 'Sisa', 'Jatuh Tempo', 'Status'];
  const t = plan.table({
    startRow: 7,
    headers: H,
    count: clamp(p.baris || 30, 1, 2000),
    row: (r, i, C) => [
      '',
      '',
      '',
      '',
      '',
      '',
      guard(`${C('Nama')}${r}`, `${C('Jumlah')}${r}-${C('Dibayar')}${r}`),
      '',
      guard(`${C('Nama')}${r}`, `IF(${C('Sisa')}${r}<=0,"LUNAS",IF(AND(${C('Jatuh Tempo')}${r}<>"",${C('Jatuh Tempo')}${r}<TODAY()),"JATUH TEMPO","BELUM LUNAS"))`),
    ],
    formats: { Tanggal: FORMAT_TANGGAL, 'Jatuh Tempo': FORMAT_TANGGAL, Jumlah: RP, Dibayar: RP, Sisa: RP },
    name: 'HutangPiutang',
  });
  plan.inputs(3, [
    { key: 'h', label: 'Total Hutang (belum lunas)', value: `=SUMIFS(${t.col('Sisa')},${t.col('Jenis')},"Hutang")`, format: RP, input: false },
    { key: 'pi', label: 'Total Piutang (belum lunas)', value: `=SUMIFS(${t.col('Sisa')},${t.col('Jenis')},"Piutang")`, format: RP, input: false },
    { key: 'jt', label: 'Jumlah Jatuh Tempo', value: `=COUNTIF(${t.col('Status')},"JATUH TEMPO")`, format: '0', input: false },
  ]);
  plan.push({ op: 'validation', address: t.col('Jenis'), list: ['Hutang', 'Piutang'] });
  plan.push({ op: 'conditional', address: t.col('Status'), rule: 'text', value: 'JATUH TEMPO', fill: '#F8CBAD', fontColor: '#9C0006' });
  plan.push({ op: 'conditional', address: t.col('Status'), rule: 'text', value: 'LUNAS', fill: '#C6EFCE', fontColor: '#006100' });
  plan.summary = 'Status otomatis berubah menjadi LUNAS / JATUH TEMPO.';
  return plan.finish({ freezeRows: t.headerRow });
}

function penjualan(p) {
  const plan = new SheetPlan('Penjualan');
  plan.title('LAPORAN PENJUALAN', p.periode || nowLabel(), 8);
  const H = ['Tanggal', 'No. Transaksi', 'Produk', 'Pelanggan', 'Qty', 'Harga Satuan', 'Diskon', 'Total'];
  const t = plan.table({
    startRow: 7,
    headers: H,
    count: clamp(p.baris || 50, 1, 5000),
    row: (r, i, C) => ['', '', '', '', '', '', '', guard(`${C('Produk')}${r}`, `${C('Qty')}${r}*${C('Harga Satuan')}${r}*(1-${C('Diskon')}${r})`)],
    formats: { Tanggal: FORMAT_TANGGAL, 'Harga Satuan': RP, Diskon: '0%', Total: RP, Qty: FORMAT_ANGKA },
    totals: { Qty: 'SUM', Total: 'SUM' },
    name: 'Penjualan',
  });
  plan.inputs(3, [
    { key: 'tot', label: 'Total Penjualan', value: `=${t.total('Total')}`, format: RP, input: false },
    { key: 'trx', label: 'Jumlah Transaksi', value: `=COUNTA(${t.col('Produk')})`, format: '0', input: false },
    { key: 'avg', label: 'Rata-rata per Transaksi', value: `=IFERROR(${t.total('Total')}/COUNTA(${t.col('Produk')}),0)`, format: RP, input: false },
  ]);
  plan.summary = 'Ketik "ringkas berdasarkan Produk" setelah data terisi untuk rekap per produk.';
  return plan.finish({ freezeRows: t.headerRow });
}

function tugas(p) {
  const plan = new SheetPlan('Daftar Tugas');
  plan.title('DAFTAR TUGAS / PROYEK', p.periode || nowLabel(), 9);
  const H = ['No', 'Tugas', 'Penanggung Jawab', 'Mulai', 'Tenggat', 'Status', 'Prioritas', 'Sisa Hari', 'Catatan'];
  const t = plan.table({
    startRow: 3,
    headers: H,
    count: clamp(p.baris || 20, 1, 2000),
    row: (r, i, C) => [i + 1, '', '', '', '', '', '', `=IF(OR(${C('Tenggat')}${r}="",${C('Status')}${r}="Selesai"),"",${C('Tenggat')}${r}-TODAY())`, ''],
    formats: { Mulai: FORMAT_TANGGAL, Tenggat: FORMAT_TANGGAL, 'Sisa Hari': '0' },
    name: 'DaftarTugas',
  });
  plan.push({ op: 'validation', address: t.col('Status'), list: ['Belum Mulai', 'Proses', 'Selesai', 'Tertunda'] });
  plan.push({ op: 'validation', address: t.col('Prioritas'), list: ['Tinggi', 'Sedang', 'Rendah'] });
  plan.push({ op: 'conditional', address: t.col('Status'), rule: 'text', value: 'Selesai', fill: '#C6EFCE', fontColor: '#006100' });
  plan.push({ op: 'conditional', address: t.col('Sisa Hari'), rule: 'less', value: 0, fill: '#F8CBAD', fontColor: '#9C0006' });
  plan.push({ op: 'conditional', address: t.col('Prioritas'), rule: 'text', value: 'Tinggi', fontColor: '#C00000' });
  plan.summary = 'Sisa hari menjadi merah jika tenggat terlewat.';
  return plan.finish({ freezeRows: t.headerRow });
}

function tabungan(p) {
  const target = p.target || p.pokok || 12000000;
  const n = clamp(Math.round(p.tenor || 12), 1, 600);
  const plan = new SheetPlan('Target Tabungan');
  plan.title('RENCANA TARGET TABUNGAN', null, 6);
  const ref = plan.inputs(3, [
    { key: 'T', label: 'Target Tabungan', value: target, format: RP },
    { key: 'n', label: 'Lama Menabung (bulan)', value: n, format: '0' },
  ]);
  plan.inputs(5, [{ key: 'per', label: 'Setoran per Bulan', value: `=ROUNDUP(${ref.T}/${ref.n},-3)`, format: RP, input: false }]);
  const now = new Date();
  const H = ['Bulan Ke', 'Bulan', 'Rencana Setoran', 'Realisasi Setoran', 'Akumulasi', 'Progres'];
  const t = plan.table({
    startRow: 8,
    headers: H,
    count: n,
    row: (r, i, C) => {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      return [i + 1, `${BULAN[d.getMonth()]} ${d.getFullYear()}`, '=$B$5', '', `=SUM(${C('Realisasi Setoran')}$9:${C('Realisasi Setoran')}${r})`, `=IF(${ref.T}=0,0,${C('Akumulasi')}${r}/${ref.T})`];
    },
    formats: { 'Rencana Setoran': RP, 'Realisasi Setoran': RP, Akumulasi: RP, Progres: FORMAT_PERSEN },
    totals: { 'Rencana Setoran': 'SUM', 'Realisasi Setoran': 'SUM' },
    name: 'TargetTabungan',
  });
  plan.push({ op: 'conditional', address: t.col('Progres'), rule: 'dataBar' });
  plan.summary = 'Isi kolom Realisasi Setoran setiap bulan untuk melihat progres.';
  return plan.finish({ freezeRows: t.headerRow });
}

export const EXCEL_TEMPLATES = [
  { id: 'anggaran', title: 'Anggaran Bulanan', icon: '💰', keywords: ['anggaran', 'budget', 'rencana keuangan', 'bujet'], build: anggaran, desc: 'Rencana vs realisasi pengeluaran, sisa dana, grafik.' },
  { id: 'arus-kas', title: 'Arus Kas', icon: '🔁', keywords: ['arus kas', 'cash flow', 'cashflow', 'buku kas', 'kas masuk', 'buku besar'], build: arusKas, desc: 'Kas masuk/keluar dengan saldo berjalan.' },
  { id: 'kas-kecil', title: 'Kas Kecil', icon: '🪙', keywords: ['kas kecil', 'petty cash'], build: kasKecil, desc: 'Buku kas kecil dengan nomor bukti.' },
  { id: 'laba-rugi', title: 'Laporan Laba Rugi', icon: '📈', keywords: ['laba rugi', 'rugi laba', 'profit loss', 'income statement', 'laporan laba'], build: labaRugi, desc: 'Pendapatan, HPP, beban, pajak, laba bersih.' },
  { id: 'neraca', title: 'Neraca', icon: '⚖️', keywords: ['neraca', 'balance sheet', 'posisi keuangan'], build: neraca, desc: 'Aset, kewajiban, ekuitas + cek seimbang.' },
  { id: 'gaji', title: 'Daftar Gaji', icon: '🧾', keywords: ['gaji', 'payroll', 'penggajian', 'upah karyawan', 'slip gaji'], build: gaji, desc: 'Gaji kotor, BPJS, potongan, gaji bersih.' },
  { id: 'faktur', title: 'Faktur / Invoice', icon: '📄', keywords: ['faktur', 'invoice', 'tagihan', 'nota'], build: faktur, desc: 'Item, diskon, PPN, total + terbilang.' },
  { id: 'cicilan', title: 'Simulasi Cicilan', icon: '🏦', keywords: ['cicilan', 'angsuran', 'pinjaman', 'kredit', 'kpr', 'amortisasi'], build: cicilan, desc: 'Jadwal angsuran anuitas atau flat.' },
  { id: 'stok', title: 'Stok / Inventaris', icon: '📦', keywords: ['stok', 'inventaris', 'persediaan', 'gudang', 'inventory'], build: stok, desc: 'Stok akhir, status restok, nilai stok.' },
  { id: 'absensi', title: 'Absensi Bulanan', icon: '🗓️', keywords: ['absensi', 'kehadiran', 'presensi', 'daftar hadir', 'absen'], build: absensi, desc: 'H/I/S/A per hari, rekap & % hadir.' },
  { id: 'nilai', title: 'Daftar Nilai', icon: '🎓', keywords: ['nilai', 'rapor', 'raport', 'nilai siswa', 'nilai mahasiswa'], build: nilai, desc: 'Nilai akhir berbobot, predikat, KKM.' },
  { id: 'keuangan', title: 'Keuangan Harian', icon: '👛', keywords: ['keuangan pribadi', 'pengeluaran harian', 'catatan keuangan', 'keuangan harian', 'pengeluaran', 'pemasukan', 'keuangan rumah tangga'], build: keuanganPribadi, desc: 'Catatan harian + ringkasan per kategori.' },
  { id: 'hutang', title: 'Hutang & Piutang', icon: '🤝', keywords: ['hutang', 'utang', 'piutang'], build: hutangPiutang, desc: 'Sisa, jatuh tempo, status lunas.' },
  { id: 'penjualan', title: 'Laporan Penjualan', icon: '🛒', keywords: ['penjualan', 'sales', 'omzet', 'omset', 'transaksi'], build: penjualan, desc: 'Transaksi, diskon, total, rata-rata.' },
  { id: 'tugas', title: 'Daftar Tugas / Proyek', icon: '✅', keywords: ['tugas', 'to do', 'todo', 'proyek', 'project', 'jadwal kerja', 'task'], build: tugas, desc: 'Status, prioritas, sisa hari ke tenggat.' },
  { id: 'tabungan', title: 'Target Tabungan', icon: '🐷', keywords: ['tabungan', 'menabung', 'target tabungan', 'saving'], build: tabungan, desc: 'Setoran per bulan & progres.' },
];

export function findExcelTemplate(text) {
  let best = null;
  for (const t of EXCEL_TEMPLATES) {
    for (const k of t.keywords) {
      if (text.includes(k) && (!best || k.length > best.len)) best = { t, len: k.length };
    }
  }
  return best ? best.t : null;
}
