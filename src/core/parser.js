// Pengenal perintah offline (bahasa Indonesia, sebagian Inggris).
// parseCommand(teks, host) -> { action, params, label } atau null.

import { normalize, parseNumber, parsePercent, splitList, parseKeyValues, BULAN } from './util.js';
import { findExcelTemplate, EXCEL_TEMPLATES } from './excel-templates.js';
import { findWordTemplate, WORD_TEMPLATES } from './word-templates.js';

const VERB = /\b(buat|template|siapkan|generate|susun|sisipkan|tambahkan|tambah|masukkan|simulasi|hitung|tulis|insert|create|make)\b/;

/** Merapikan kata-kata sopan & sinonim supaya pola lebih mudah dicocokkan. */
export function canonical(text) {
  let n = normalize(text);
  n = n.replace(/^(tolong|coba|mohon|bisa|bisakah|silakan|silahkan|please|hai master|master|ok|oke)\b[ ,]*/g, '');
  n = n.replace(/^(tolong|coba|mohon|bisa|bisakah|silakan|silahkan|please)\b[ ,]*/g, '');
  n = n.replace(/\b(buatkan|bikinkan|bikin|buatin|bikinin|buat kan)\b/g, 'buat');
  n = n.replace(/\b(dong|ya|yah|deh|sih|kak|min)\b/g, '').replace(/[?!.]+$/g, '');
  return n.replace(/\s+/g, ' ').trim();
}

const COLORS = {
  merah: { fill: '#F8CBAD', font: '#9C0006' },
  hijau: { fill: '#C6EFCE', font: '#006100' },
  kuning: { fill: '#FFEB9C', font: '#9C5700' },
  biru: { fill: '#DDEBF7', font: '#1F4E78' },
  oranye: { fill: '#FCE4D6', font: '#C65911' },
  orange: { fill: '#FCE4D6', font: '#C65911' },
  ungu: { fill: '#E4DFEC', font: '#5B2C6F' },
  abu: { fill: '#EDEDED', font: '#3A3A3A' },
};

function pickColor(n, def = 'merah') {
  const k = Object.keys(COLORS).find((c) => n.includes(c)) || def;
  return Object.assign({ name: k }, COLORS[k]);
}

const NUM = String.raw`(-?(?:rp\.?\s*)?[\d.,]+(?:\s*(?:juta|jt|ribu|rb|miliar|milyar|triliun|k))?)`;

function grab(n, re) {
  const m = n.match(re);
  return m ? m[1] : undefined;
}

/** Parameter umum untuk template (jumlah baris, nominal, periode, dll). */
export function templateParams(n, raw) {
  const p = {};
  const baris = grab(n, /(\d+)\s*(?:baris|row|orang|karyawan|siswa|item|barang)\b/);
  if (baris) p.baris = parseInt(baris, 10);
  const pokok = grab(n, new RegExp(String.raw`(?:pinjaman|pokok|plafon|kredit|utang|hutang|harga)\s*(?:sebesar|senilai|=|:)?\s*` + NUM));
  if (pokok) p.pokok = parseNumber(pokok);
  const target = grab(n, new RegExp(String.raw`target\s*(?:sebesar|senilai|=|:)?\s*` + NUM));
  if (target) p.target = parseNumber(target);
  const bunga = grab(n, /bunga\s*(?:sebesar|=|:)?\s*([\d.,]+\s*%?)/);
  if (bunga) p.bunga = parsePercent(bunga.includes('%') ? bunga : bunga + '%');
  const tenor = n.match(/(?:tenor|selama|jangka waktu|lama)?\s*(\d+)\s*(bulan|bln|tahun|thn)\b/);
  if (tenor && !/\b(bulan|tahun) (ini|depan|lalu)\b/.test(tenor[0])) {
    const v = parseInt(tenor[1], 10);
    p.tenor = /tahun|thn/.test(tenor[2]) ? v * 12 : v;
  }
  if (/\bflat\b/.test(n)) p.metode = 'flat';
  const saldo = grab(n, new RegExp(String.raw`(?:saldo awal|saldo|dana)\s*(?:sebesar|=|:)?\s*` + NUM));
  if (saldo) p.saldo = parseNumber(saldo);
  const masuk = grab(n, new RegExp(String.raw`(?:pemasukan|penghasilan|pendapatan|gaji saya)\s*(?:sebesar|=|:)?\s*` + NUM));
  if (masuk) p.pemasukan = parseNumber(masuk);
  const bulanIdx = BULAN.findIndex((b) => new RegExp(`\\b${b.toLowerCase()}\\b`).test(n));
  if (bulanIdx >= 0) {
    p.bulan = bulanIdx;
    const th = grab(n, new RegExp(`${BULAN[bulanIdx].toLowerCase()}\\s+(\\d{4})`));
    if (th) p.tahun = parseInt(th, 10);
    p.periode = `${BULAN[bulanIdx]} ${p.tahun || new Date().getFullYear()}`;
  }
  const kelas = grab(n, /kelas\s+([a-z0-9 -]{1,12}?)(?:\s+\d+\s+siswa|$|,)/);
  if (kelas) p.kelas = kelas.toUpperCase();
  const kkm = grab(n, /kkm\s*(\d+)/);
  if (kkm) p.kkm = parseInt(kkm, 10);
  const ppn = grab(n, /ppn\s*([\d.,]+\s*%?)/);
  if (ppn) p.ppn = parsePercent(ppn.includes('%') ? ppn : ppn + '%');
  const kepada = grab(raw || '', /(?:untuk|kepada)\s+(?:pelanggan\s+)?([A-Z][\w .]+?)(?:,|$)/);
  if (kepada) p.kepada = kepada.trim();
  return p;
}

function tableSpec(n, raw) {
  // "tabel 5x3" / "tabel 5 kolom 10 baris"
  let m = n.match(/tabel\s+(\d+)\s*[x×]\s*(\d+)/);
  if (m) return { cols: parseInt(m[1], 10), rows: parseInt(m[2], 10), headers: null };
  m = n.match(/tabel\s+(\d+)\s*kolom(?:\s+(?:dan\s+)?(\d+)\s*baris)?/);
  if (m) return { cols: parseInt(m[1], 10), rows: m[2] ? parseInt(m[2], 10) : 10, headers: null };
  // "tabel karyawan dengan kolom nama, jabatan, gaji (sebanyak) 10 baris"
  m = n.match(/tabel\s*(.*?)\s*(?:dengan |yang berisi |berisi )?(?:kolom|kolomnya|header)\s*:?\s*(.+)$/);
  if (m) {
    let colText = m[2];
    let rows = 10;
    const r = colText.match(/[,\s]+(?:sebanyak\s+)?(\d+)\s*baris\b.*$/);
    if (r) {
      rows = parseInt(r[1], 10);
      colText = colText.slice(0, r.index);
    }
    // Ambil kapitalisasi asli dari teks mentah bila ada
    const rawCols = (raw || '').match(/(?:kolom|kolomnya|header)\s*:?\s*(.+)$/i);
    let headers = splitList(colText);
    if (rawCols) {
      const rc = splitList(rawCols[1].replace(/[,\s]+(?:sebanyak\s+)?\d+\s*baris\b.*$/i, ''));
      if (rc.length === headers.length) headers = rc;
    }
    const name = m[1].replace(/^(baru|data|untuk)\s+/, '').trim();
    return { headers, rows, name };
  }
  return null;
}

function columnRef(n) {
  // "berdasarkan kolom gaji" / "berdasarkan gaji" / "kolom C"
  const m = n.match(/(?:berdasarkan|menurut|by|per|kolom)\s+(?:kolom\s+)?([a-z0-9 _.%-]+?)(?:\s+(?:dari|secara|naik|turun|terbesar|terkecil|ascending|descending)\b|$)/);
  return m ? m[1].trim() : undefined;
}

function chartType(n) {
  if (/\b(pie|lingkaran|pai|donat|donut)\b/.test(n)) return /donat|donut/.test(n) ? 'Doughnut' : 'Pie';
  if (/\b(garis|line|tren|trend)\b/.test(n)) return 'Line';
  if (/\b(horizontal|bar)\b/.test(n)) return 'BarClustered';
  if (/\b(area)\b/.test(n)) return 'Area';
  if (/\b(sebar|scatter)\b/.test(n)) return 'XYScatter';
  if (/\b(tumpuk|stacked)\b/.test(n)) return 'ColumnStacked';
  return 'ColumnClustered';
}

function caseMode(n) {
  if (/huruf (besar|kapital) semua|uppercase|huruf besar|kapitalkan semua|jadikan kapital/.test(n)) return 'upper';
  if (/huruf kecil|lowercase/.test(n)) return 'lower';
  if (/kapital (di )?(awal|tiap|setiap)|awal kata|title case|proper/.test(n)) return 'title';
  if (/awal kalimat|sentence case/.test(n)) return 'sentence';
  return null;
}

function conditionalRule(n) {
  const op = [
    [/(?:di atas|lebih dari|lebih besar dari|>|melebihi|diatas)\s*/, 'greater'],
    [/(?:di bawah|kurang dari|lebih kecil dari|<|dibawah)\s*/, 'less'],
    [/(?:sama dengan|=|bernilai)\s*/, 'equal'],
  ];
  for (const [re, rule] of op) {
    const m = n.match(new RegExp(re.source + NUM));
    if (m) return { rule, value: parseNumber(m[1]) };
  }
  const t = n.match(/(?:berisi|mengandung|teks|kata)\s+["']?([^"']+?)["']?(?:\s+(?:dengan\s+)?(?:warna\s+)?(?:merah|hijau|kuning|biru|oranye|orange|ungu|abu))?$/);
  if (t) return { rule: 'text', value: t[1] };
  if (/duplikat|ganda/.test(n)) return { rule: 'duplicate' };
  return null;
}

/* ---------------------------- EXCEL ---------------------------- */

function parseExcel(n, raw) {
  let m;
  // Tabel dengan kolom eksplisit didahulukan ("tabel karyawan kolom nama, gaji" bukan template gaji)
  if (/\btabel\b/.test(n) && /\b(kolom|header|\d+\s*[x×]\s*\d+)/.test(n) && /\b(buat|tambah|sisipkan)\b/.test(n)) {
    const spec = tableSpec(n, raw);
    if (spec) return { action: 'excel.table', params: spec, label: 'Buat tabel' };
  }
  const tpl = findExcelTemplate(n);
  if (tpl && (VERB.test(n) || n.startsWith(tpl.keywords[0]) || /^(laporan|daftar|catatan|simulasi|rencana|jadwal)\b/.test(n))) {
    // Hindari salah tangkap: "jumlahkan total penjualan" bukan permintaan template
    if (!/^(jumlahkan|urutkan|filter|format|grafik|tandai|warnai|hapus|ringkas|rekap)\b/.test(n)) {
      return { action: 'excel.template', params: { id: tpl.id, ...templateParams(n, raw) }, label: `Template: ${tpl.title}` };
    }
  }
  if (/\btabel\b/.test(n) && /\b(buat|tambah|sisipkan)\b/.test(n) && !/jadikan|ubah|format sebagai/.test(n)) {
    const spec = tableSpec(n, raw);
    if (spec) return { action: 'excel.table', params: spec, label: 'Buat tabel' };
  }
  if (/(jadikan|ubah (ke|jadi|menjadi)|format sebagai|konversi (ke|jadi)) tabel/.test(n)) return { action: 'excel.toTable', params: {}, label: 'Jadikan tabel' };
  if (/^(ringkas|rekap|kelompokkan|pivot|ringkasan)\b/.test(n) || /\b(rekap|ringkas|total) (per|berdasarkan)\b/.test(n)) {
    return { action: 'excel.groupSum', params: { column: columnRef(n), value: grab(n, /(?:jumlah|nilai|total)\s+(?:kolom\s+)?([a-z0-9 _]+?)\s+(?:per|berdasarkan)/) }, label: 'Rekap per kelompok' };
  }
  if (/perbarui terbilang|update terbilang|refresh terbilang/.test(n)) return { action: 'excel.refreshTerbilang', params: {}, label: 'Perbarui terbilang' };
  if (/\bterbilang\b/.test(n)) {
    const num = grab(n, new RegExp('terbilang\\s+' + NUM));
    return { action: 'excel.terbilang', params: { number: num ? parseNumber(num) : null }, label: 'Terbilang' };
  }
  if (/^(jumlahkan|total|totalkan|sum|tambahkan total|hitung total|beri total)\b/.test(n) || /\b(baris|kolom) total\b/.test(n)) return { action: 'excel.aggregate', params: { fn: 'SUM' }, label: 'Jumlahkan (SUM)' };
  if (/\b(rata-rata|rata rata|average|rerata)\b/.test(n)) return { action: 'excel.aggregate', params: { fn: 'AVERAGE' }, label: 'Rata-rata' };
  if (/\b(maksimum|tertinggi|terbesar|max)\b/.test(n) && !/urut/.test(n)) return { action: 'excel.aggregate', params: { fn: 'MAX' }, label: 'Nilai maksimum' };
  if (/\b(minimum|terendah|terkecil|min)\b/.test(n) && !/urut/.test(n)) return { action: 'excel.aggregate', params: { fn: 'MIN' }, label: 'Nilai minimum' };
  if (/\b(hitung jumlah data|banyak data|count|cacah)\b/.test(n)) return { action: 'excel.aggregate', params: { fn: 'COUNTA' }, label: 'Hitung data' };
  if (/\b(nomor urut|penomoran|beri nomor|nomori)\b/.test(n)) return { action: 'excel.numbering', params: {}, label: 'Nomor urut' };
  if (/\b(urutkan|sortir|sort|urut)\b/.test(n)) {
    const desc = /\b(turun|menurun|besar ke kecil|terbesar|z-a|z ke a|descending|desc|terbaru|tertinggi)\b/.test(n);
    return { action: 'excel.sort', params: { ascending: !desc, column: columnRef(n) }, label: `Urutkan ${desc ? 'menurun' : 'menaik'}` };
  }
  if (/\b(tandai|warnai|sorot|highlight|beri warna)\b/.test(n)) {
    const rule = conditionalRule(n);
    if (rule) {
      const c = pickColor(n);
      return { action: 'excel.highlight', params: { ...rule, fill: c.fill, fontColor: c.font }, label: `Tandai sel (${rule.rule})` };
    }
  }
  if (/\b(hapus|buang) (data )?(duplikat|ganda|dobel)\b/.test(n)) return { action: 'excel.dedupe', params: {}, label: 'Hapus duplikat' };
  if (/\b(filter|saring|autofilter)\b/.test(n)) return { action: 'excel.filter', params: {}, label: 'Filter' };
  if (/\b(grafik|chart|diagram|bagan)\b/.test(n)) {
    const title = grab(raw ? normalize(raw) : n, /judul(?:nya)?\s+["']?(.+?)["']?$/);
    return { action: 'excel.chart', params: { type: chartType(n), title }, label: 'Buat grafik' };
  }
  m = n.match(/\b(?:format|jadikan|ubah (?:ke|jadi|menjadi)|ganti format)\s+(?:ke\s+|sebagai\s+|jadi\s+)?(rupiah|uang|mata uang|angka|ribuan|persen|persentase|tanggal|teks)\b/);
  if (m) {
    const kind = /rupiah|uang/.test(m[1]) ? 'rupiah' : /angka|ribuan/.test(m[1]) ? 'angka' : /persen/.test(m[1]) ? 'persen' : /tanggal/.test(m[1]) ? 'tanggal' : 'teks';
    return { action: 'excel.numberFormat', params: { kind }, label: `Format ${kind}` };
  }
  if (/\b(persentase|% dari total|porsi)\b/.test(n)) return { action: 'excel.percentOfTotal', params: {}, label: 'Persentase dari total' };
  if (/\b(ranking|rangking|peringkat|rank)\b/.test(n)) return { action: 'excel.rank', params: { ascending: /terkecil|terendah/.test(n) }, label: 'Peringkat' };
  m = n.match(/\b(?:sheet|lembar|worksheet|tab) baru(?:\s+(?:bernama|dengan nama|nama))?\s*(.*)$/) || n.match(/\btambah (?:sheet|lembar)(?:\s+(?:bernama|nama))?\s*(.*)$/);
  if (m) {
    const rawName = (raw || '').match(/\bbaru(?:\s+(?:bernama|dengan nama|nama))?\s+(.+)$/i) || (raw || '').match(/\b(?:sheet|lembar)(?:\s+(?:bernama|dengan nama|nama))?\s+(.+)$/i);
    return { action: 'excel.newSheet', params: { name: (rawName ? rawName[1] : m[1]).trim() }, label: 'Sheet baru' };
  }
  if (/\btranspos/.test(n)) return { action: 'excel.transpose', params: {}, label: 'Transpos' };
  if (/hapus (semua )?baris kosong/.test(n)) return { action: 'excel.removeEmptyRows', params: {}, label: 'Hapus baris kosong' };
  if (/\b(lepas|buka|hapus) (bekuan|freeze|pembekuan)\b|unfreeze/.test(n)) return { action: 'excel.unfreeze', params: {}, label: 'Lepas bekuan' };
  if (/\b(bekukan|freeze|kunci) (baris|header|judul|panel|pane)/.test(n) || /^bekukan\b/.test(n)) return { action: 'excel.freeze', params: {}, label: 'Bekukan header' };
  if (/\b(hapus|buang|rapikan) (spasi|spasi ganda|spasi berlebih)\b|\btrim\b/.test(n)) return { action: 'excel.trim', params: {}, label: 'Rapikan spasi' };
  const cm = caseMode(n);
  if (cm) return { action: 'excel.case', params: { mode: cm }, label: 'Ubah huruf' };
  if (/\b(beri garis|border|garis tepi|bingkai|garis tabel)\b/.test(n)) return { action: 'excel.border', params: {}, label: 'Beri garis' };
  if (/\b(tebalkan|bold)\b/.test(n)) return { action: 'excel.bold', params: { headerOnly: /header|judul|baris pertama/.test(n) }, label: 'Tebalkan' };
  if (/\b(hapus format|bersihkan format|clear format)\b/.test(n)) return { action: 'excel.clearFormat', params: {}, label: 'Hapus format' };
  if (/\b(rapikan|autofit|lebar kolom otomatis|sesuaikan lebar)\b/.test(n)) return { action: 'excel.autofit', params: {}, label: 'Rapikan kolom' };
  if (/\b(tanggal hari ini|sisipkan tanggal|isi tanggal)\b/.test(n)) return { action: 'excel.today', params: {}, label: 'Tanggal hari ini' };
  return null;
}

/* ---------------------------- WORD ---------------------------- */

function parseWord(n, raw) {
  let m;
  const tpl = findWordTemplate(n);
  if (tpl && (VERB.test(n) || /^(surat|memo|notulen|notula|kwitansi|kuitansi|berita acara|permohonan|undangan|lamaran)\b/.test(n))) {
    return { action: 'word.template', params: { id: tpl.id, given: parseKeyValues(raw) }, label: `Template: ${tpl.title}` };
  }
  if (/\btabel\b/.test(n) && /\b(buat|tambah|sisipkan)\b/.test(n)) {
    const spec = tableSpec(n, raw);
    if (spec) return { action: 'word.table', params: spec, label: 'Sisipkan tabel' };
  }
  m = (raw || '').match(/^(?:tolong\s+)?(?:ganti|ubah|replace)\s+(?:kata\s+|semua\s+)?["“']?(.+?)["”']?\s+(?:dengan|menjadi|jadi|ke|with)\s+["“']?(.+?)["”']?\s*$/i);
  if (m) return { action: 'word.replace', params: { find: m[1], replace: m[2] }, label: `Ganti "${m[1]}"` };
  if (/\bterbilang\b/.test(n)) {
    const num = grab(n, new RegExp('terbilang\\s+' + NUM));
    return { action: 'word.terbilang', params: { number: num ? parseNumber(num) : null, rupiah: /rupiah|rp/.test(n) || /rp/.test(num || '') }, label: 'Terbilang' };
  }
  if (/\b(tanggal hari ini|sisipkan tanggal|tulis tanggal)\b/.test(n)) return { action: 'word.date', params: {}, label: 'Tanggal hari ini' };
  if (/\b(hitung kata|jumlah kata|berapa kata|statistik dokumen|word count)\b/.test(n)) return { action: 'word.count', params: {}, label: 'Statistik dokumen' };
  if (/\b(hapus|rapikan|buang) spasi( ganda| berlebih)?\b/.test(n)) return { action: 'word.cleanSpaces', params: {}, label: 'Rapikan spasi' };
  if (/hapus (baris|paragraf|enter) kosong/.test(n)) return { action: 'word.removeEmpty', params: {}, label: 'Hapus paragraf kosong' };
  if (/\b(daftar isi|table of contents|toc)\b/.test(n)) return { action: 'word.toc', params: {}, label: 'Daftar isi' };
  if (/\bnomor halaman\b|page number/.test(n)) return { action: 'word.pageNumbers', params: {}, label: 'Nomor halaman' };
  if (/\b(halaman baru|page break|pindah halaman|ganti halaman)\b/.test(n)) return { action: 'word.pageBreak', params: {}, label: 'Halaman baru' };
  m = (raw || '').match(/^(?:buat\s+|tambah(?:kan)?\s+|isi\s+)?(header|kepala halaman|footer|kaki halaman)\s*:?\s+(.+)$/i);
  if (m) return { action: /footer|kaki/i.test(m[1]) ? 'word.footer' : 'word.header', params: { text: m[2] }, label: m[1] };
  if (/\b(tanda tangan|ttd|blok tanda tangan)\b/.test(n)) return { action: 'word.signature', params: {}, label: 'Blok tanda tangan' };
  m = (raw || '').match(/^(?:sisipkan\s+|buat\s+|tulis\s+)?judul\s*:?\s+(.+)$/i);
  if (m && !/^(ini|yang|terpilih)/i.test(m[1])) return { action: 'word.heading', params: { text: m[1] }, label: 'Sisipkan judul' };
  if (/\bjadikan (judul|heading)( \d)?\b/.test(n)) return { action: 'word.applyHeading', params: { level: parseInt(grab(n, /(?:judul|heading) (\d)/) || '1', 10) }, label: 'Jadikan judul' };
  m = n.match(/\b(?:sorot|stabilo|highlight|tandai)\s+(?:kata\s+|semua\s+)?["']?(.+?)["']?(?:\s+(?:dengan\s+)?(?:warna\s+)?(kuning|hijau|biru|merah|ungu|abu))?$/);
  if (m) return { action: 'word.highlight', params: { find: m[1], color: m[2] || 'kuning' }, label: `Sorot "${m[1]}"` };
  if (/\b(daftar bernomor|penomoran|numbering)\b/.test(n)) return { action: 'word.list', params: { ordered: true }, label: 'Daftar bernomor' };
  if (/\b(poin|bullet|butir|daftar poin)\b/.test(n)) return { action: 'word.list', params: { ordered: false }, label: 'Daftar poin' };
  const scope = /\b(semua|seluruh|dokumen|semuanya)\b/.test(n) ? 'document' : 'selection';
  const cm = caseMode(n);
  if (cm) return { action: 'word.case', params: { mode: cm, scope }, label: 'Ubah huruf' };
  m = n.match(/\bspasi\s+(1[.,]15|1[.,]5|1|2|ganda|tunggal|satu setengah)\b/);
  if (m) {
    const v = { ganda: 2, tunggal: 1, 'satu setengah': 1.5 }[m[1]] || parseFloat(m[1].replace(',', '.'));
    return { action: 'word.spacing', params: { value: v, scope }, label: `Spasi ${v}` };
  }
  m = n.match(/\brata\s+(kiri kanan|kanan kiri|kiri-kanan|kiri|kanan|tengah)\b|\b(justify|center|rata tengah)\b/);
  if (m) {
    const k = m[1] || m[2];
    const a = /kiri.kanan|kanan.kiri|justify/.test(k) ? 'Justified' : /tengah|center/.test(k) ? 'Centered' : k === 'kanan' ? 'Right' : 'Left';
    return { action: 'word.align', params: { align: a, scope }, label: 'Perataan' };
  }
  const font = n.match(/\b(?:font|huruf|jenis huruf)\s+(times new roman|times|arial|calibri|cambria|verdana|tahoma|georgia|garamond|book antiqua|century gothic|segoe ui|bookman old style)\b/);
  const size = grab(n, /\b(?:ukuran|size)\s*(?:huruf|font)?\s*(\d{1,2}(?:[.,]5)?)\b/) || grab(n, /\b(\d{1,2})\s*(?:pt|poin)\b/);
  if (font || size) {
    const names = { times: 'Times New Roman' };
    const fname = font ? names[font[1]] || font[1].replace(/\b\w/g, (c) => c.toUpperCase()) : undefined;
    return { action: 'word.font', params: { name: fname, size: size ? parseFloat(size.replace(',', '.')) : undefined, scope }, label: 'Ubah font' };
  }
  if (/\b(tebalkan|bold|cetak tebal|miringkan|italic|cetak miring|garis bawahi|underline)\b/.test(n)) {
    return {
      action: 'word.style',
      params: { bold: /tebal|bold/.test(n) || undefined, italic: /miring|italic/.test(n) || undefined, underline: /garis bawah|underline/.test(n) || undefined },
      label: 'Gaya teks',
    };
  }
  if (/\b(rapikan dokumen|format standar|format dokumen|rapikan format)\b/.test(n)) return { action: 'word.tidy', params: {}, label: 'Rapikan dokumen' };
  return null;
}

/* ---------------------------- UMUM ---------------------------- */

export function parseCommand(text, host) {
  const n = canonical(text);
  if (!n) return null;
  if (/^(bantuan|help|\?|perintah apa saja|daftar perintah|apa saja perintahnya|contoh perintah)$/.test(n)) return { action: 'help', params: {}, label: 'Bantuan' };
  if (/^(template|daftar template|lihat template|semua template)$/.test(n)) return { action: 'templates', params: {}, label: 'Daftar template' };
  let m = n.match(/^(?:jalankan|run) makro\s+(.+)$/);
  if (m) return { action: 'macro.run', params: { name: m[1] }, label: `Makro ${m[1]}` };
  if (/^(batalkan|undo)$/.test(n)) return { action: 'undoHint', params: {}, label: 'Batalkan' };
  const raw = String(text || '').trim();
  return host === 'word' ? parseWord(n, raw) : parseExcel(n, raw);
}

/**
 * Memecah perintah berantai: baris baru, ";", " lalu ", " kemudian ", " setelah itu ".
 * Rantai hanya dipakai bila SEMUA bagian dikenali mesin offline.
 */
export function parseChain(text, host) {
  const parts = String(text || '')
    .split(/\n+|;\s*|\s+(?:lalu|kemudian|setelah itu|terus|and then|then)\s+/i)
    .map((s) => s.trim())
    .filter(Boolean);
  if (parts.length < 2) {
    const single = parseCommand(text, host);
    return single ? [single] : null;
  }
  const parsed = parts.map((p) => parseCommand(p, host));
  if (parsed.every(Boolean)) return parsed;
  const whole = parseCommand(text, host);
  return whole ? [whole] : null;
}

/**
 * Apakah perintah lebih cocok dikerjakan AI (isi kreatif, analisis, pertanyaan)?
 * Dipakai di mode Otomatis: perintah sederhana -> offline, kompleks -> AI.
 */
export function prefersAI(text) {
  const n = canonical(text);
  if (/\b(analisis|analisa|jelaskan|kenapa|mengapa|bagaimana|apa itu|ringkas(kan)? (dokumen|isi|teks)|rangkum|terjemah|translate|perbaiki (kalimat|tata bahasa|ejaan|rumus)|parafrase|tulis ulang|karang|buat (isi|konten|artikel|esai|cerita|pidato|proposal|laporan kegiatan)|isi (dengan|data)|contoh data|prediksi|saran|rekomendasikan|bandingkan|cari kesalahan|periksa)\b/.test(n)) return true;
  if (n.split(' ').length > 18) return true;
  return false;
}

export const ALL_TEMPLATES = { excel: EXCEL_TEMPLATES, word: WORD_TEMPLATES };

export const HELP = {
  excel: [
    ['Template keuangan', ['buat anggaran bulanan', 'buat laporan arus kas saldo awal 5 juta', 'buat laporan laba rugi', 'buat neraca', 'simulasi cicilan pinjaman 50 juta bunga 10% 24 bulan', 'cicilan kpr 300 juta bunga 7,5% 15 tahun flat', 'buat faktur untuk PT Maju', 'buat daftar gaji 15 karyawan', 'buat catatan keuangan harian', 'buat target tabungan 20 juta 10 bulan']],
    ['Template lain', ['buat daftar stok barang', 'buat absensi bulan desember', 'buat daftar nilai kelas 7A kkm 75', 'buat daftar tugas proyek', 'buat laporan penjualan', 'buat catatan hutang piutang']],
    ['Tabel', ['buat tabel karyawan dengan kolom Nama, Jabatan, Gaji 10 baris', 'buat tabel 5x10', 'jadikan tabel']],
    ['Hitung (pilih sel dulu)', ['jumlahkan', 'rata-rata', 'nilai maksimum', 'persentase dari total', 'ranking', 'rekap total per Produk', 'terbilang', 'terbilang 1.250.000']],
    ['Atur data', ['urutkan berdasarkan Gaji turun', 'filter', 'hapus duplikat', 'hapus baris kosong', 'transpos', 'nomor urut', 'hapus spasi', 'huruf kapital tiap kata']],
    ['Format & visual', ['format rupiah', 'format persen', 'format tanggal', 'rapikan kolom', 'beri garis', 'tebalkan header', 'bekukan header', 'tandai nilai di atas 1 juta merah', 'tandai duplikat kuning', 'grafik batang', 'grafik lingkaran judul Komposisi Biaya']],
    ['Lainnya', ['sheet baru Laporan 2026', 'tanggal hari ini', 'perbarui terbilang', 'buat tabel anggaran lalu format rupiah (rantai perintah)']],
  ],
  word: [
    ['Surat otomatis', ['surat izin sakit', 'surat undangan rapat', 'surat lamaran kerja posisi: Admin, perusahaan: PT Maju', 'surat kuasa', 'surat pernyataan', 'surat keterangan kerja', 'surat pengunduran diri', 'surat permohonan', 'permohonan cuti', 'surat tugas', 'surat edaran', 'surat penawaran harga', 'surat peringatan sp1', 'surat rekomendasi']],
    ['Dokumen', ['buat notulen rapat', 'buat memo', 'buat berita acara serah terima', 'buat kwitansi', 'judul Laporan Kegiatan', 'daftar isi', 'nomor halaman', 'halaman baru', 'blok tanda tangan']],
    ['Tabel & teks', ['buat tabel 4x5', 'buat tabel kolom No, Nama, Alamat 10 baris', 'tanggal hari ini', 'terbilang 2.500.000 rupiah']],
    ['Edit cepat', ['ganti "PT Lama" dengan "PT Baru"', 'sorot kata penting kuning', 'hapus spasi ganda', 'hapus paragraf kosong', 'hitung kata']],
    ['Format', ['font Times New Roman ukuran 12 semua', 'spasi 1,5 semua', 'rata kiri kanan semua', 'huruf besar', 'tebalkan', 'jadikan judul', 'daftar poin', 'daftar bernomor', 'rapikan dokumen', 'header PT Maju Jaya', 'footer Rahasia']],
  ],
  ai: [
    'Tanyakan apa saja atau minta hal yang rumit, misalnya:',
    'Excel: "analisis data penjualan ini dan beri 3 insight", "buat rumus komisi 5% jika penjualan di atas 10 juta", "kenapa rumus di C5 error?", "isi tabel dengan 20 contoh data produk"',
    'Word: "ringkas dokumen ini jadi 5 poin", "perbaiki tata bahasa paragraf terpilih", "buat proposal kegiatan 17 Agustus", "terjemahkan ke bahasa Inggris"',
  ],
};
