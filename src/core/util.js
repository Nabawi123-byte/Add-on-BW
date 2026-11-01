// Utilitas umum: normalisasi teks, angka gaya Indonesia, tanggal, dan alamat sel.

export const BULAN = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
];
export const HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

export const FORMAT_RUPIAH = '"Rp "#,##0;[Red]-"Rp "#,##0';
export const FORMAT_ANGKA = '#,##0';
export const FORMAT_PERSEN = '0.0%';
export const FORMAT_TANGGAL = 'dd/mm/yyyy';

/** Huruf kecil, tanpa aksen, spasi dirapikan. */
export function normalize(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

export function capitalize(s) {
  s = String(s || '');
  return s.charAt(0).toUpperCase() + s.slice(1);
}

export function titleCase(s) {
  return String(s || '').toLowerCase().replace(/(^|[\s\-/(])([a-zÀ-ɏ])/g, (m, a, b) => a + b.toUpperCase());
}

const PENGALI = [
  [/^(triliun|t)$/, 1e12],
  [/^(miliar|milyar|m|b)$/, 1e9],
  [/^(juta|jt|j)$/, 1e6],
  [/^(ribu|rb|k)$/, 1e3],
  [/^(ratus)$/, 100],
];

/**
 * Mengubah teks angka gaya Indonesia menjadi number.
 * Contoh: "Rp 1.500.000" -> 1500000, "1,5 juta" -> 1500000, "50jt" -> 50000000,
 * "2.5" -> 2.5, "12%" -> 12 (persen dibiarkan sebagai angka; pakai parsePercent).
 */
export function parseNumber(input) {
  if (typeof input === 'number') return Number.isFinite(input) ? input : null;
  let s = normalize(input).replace(/^rp\.?\s*/, '').replace(/rupiah$/, '').replace(/%$/, '').trim();
  if (!s) return null;
  const m = s.match(/^(-?[\d.,]+)\s*([a-z]+)?$/);
  if (!m) return null;
  let num = m[1];
  const unit = m[2];
  const dots = (num.match(/\./g) || []).length;
  const commas = (num.match(/,/g) || []).length;
  if (dots && commas) {
    // Format paling umum di Indonesia: 1.234.567,89
    if (num.lastIndexOf(',') > num.lastIndexOf('.')) num = num.replace(/\./g, '').replace(',', '.');
    else num = num.replace(/,/g, '');
  } else if (dots > 1) {
    num = num.replace(/\./g, '');
  } else if (commas > 1) {
    num = num.replace(/,/g, '');
  } else if (dots === 1) {
    // "1.500" = seribu lima ratus; "1.5" atau "2.75" = desimal
    const after = num.split('.')[1];
    if (after.length === 3 && !unit) num = num.replace('.', '');
  } else if (commas === 1) {
    const after = num.split(',')[1];
    if (after.length === 3 && !unit) num = num.replace(',', '');
    else num = num.replace(',', '.');
  }
  let value = parseFloat(num);
  if (!Number.isFinite(value)) return null;
  if (unit) {
    const p = PENGALI.find(([re]) => re.test(unit));
    if (!p) return null;
    value *= p[1];
  }
  return Math.round(value * 1e6) / 1e6;
}

/** "10%" atau "10" -> 0.1 ; "0.1" -> 0.1 */
export function parsePercent(input) {
  const raw = normalize(input);
  const n = parseNumber(raw);
  if (n === null) return null;
  if (raw.endsWith('%') || n > 1) return n / 100;
  return n;
}

export function formatRupiah(n) {
  const v = Math.round(Number(n) || 0);
  const s = Math.abs(v).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (v < 0 ? '-Rp ' : 'Rp ') + s;
}

export function formatAngka(n) {
  const v = Number(n) || 0;
  const [i, d] = Math.abs(v).toFixed(2).split('.');
  const intPart = i.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (v < 0 ? '-' : '') + intPart + (d !== '00' ? ',' + d.replace(/0$/, '') : '');
}

export function tanggalIndonesia(date) {
  const d = date instanceof Date ? date : new Date();
  return `${d.getDate()} ${BULAN[d.getMonth()]} ${d.getFullYear()}`;
}

export function hariTanggal(date) {
  const d = date instanceof Date ? date : new Date();
  return `${HARI[d.getDay()]}, ${tanggalIndonesia(d)}`;
}

/** Angka kolom (1-based) -> huruf kolom Excel. 1 -> A, 27 -> AA */
export function columnLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Huruf kolom -> angka (1-based). */
export function columnNumber(letters) {
  let n = 0;
  for (const ch of String(letters).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** Alamat A1 dari baris & kolom (1-based). */
export function cellAddress(row, col) {
  return columnLetter(col) + row;
}

export function rangeAddress(row1, col1, row2, col2) {
  return `${cellAddress(row1, col1)}:${cellAddress(row2, col2)}`;
}

/** Memecah alamat "Sheet1!B3:D10" menjadi bagian-bagiannya (1-based). */
export function parseAddress(address) {
  let a = String(address || '');
  let sheet = null;
  const bang = a.lastIndexOf('!');
  if (bang >= 0) {
    sheet = a.slice(0, bang).replace(/^'|'$/g, '');
    a = a.slice(bang + 1);
  }
  const m = a.replace(/\$/g, '').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!m) return null;
  const c1 = columnNumber(m[1]);
  const r1 = parseInt(m[2], 10);
  const c2 = m[3] ? columnNumber(m[3]) : c1;
  const r2 = m[4] ? parseInt(m[4], 10) : r1;
  return {
    sheet,
    row: Math.min(r1, r2),
    col: Math.min(c1, c2),
    lastRow: Math.max(r1, r2),
    lastCol: Math.max(c1, c2),
  };
}

/** "nama, jabatan dan gaji" -> ["nama", "jabatan", "gaji"] */
export function splitList(text) {
  return String(text || '')
    .split(/\s*(?:,|;|\/|\s+dan\s+|\s+&\s+|\|)\s*/i)
    .map((s) => s.trim())
    .filter(Boolean);
}

/**
 * Mengambil pasangan "kunci: nilai" dari teks perintah.
 * "surat izin nama: Budi, jabatan: Staf IT; alasan: sakit" -> {nama:'Budi', jabatan:'Staf IT', alasan:'sakit'}
 */
export function parseKeyValues(text) {
  const out = {};
  const re = /([a-z_][a-z0-9_ ]{0,30}?)\s*[:=]\s*([^,;\n]+)/gi;
  let m;
  while ((m = re.exec(String(text || '')))) {
    const key = normalize(m[1]).replace(/\s+/g, '_');
    out[key] = m[2].trim();
  }
  return out;
}

export function escapeHtml(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Kata kunci kolom yang biasanya berisi uang. */
export function looksLikeMoneyColumn(header) {
  return /(gaji|harga|jumlah|total|biaya|nominal|saldo|pendapatan|pengeluaran|pemasukan|upah|tunjangan|potongan|bayar|tagihan|hutang|utang|piutang|nilai stok|anggaran|realisasi|modal|omzet|omset|laba|rugi|debit|kredit|rp)/i.test(
    String(header || ''),
  );
}

export function looksLikeDateColumn(header) {
  return /(tanggal|tgl|tenggat|jatuh tempo|deadline|mulai|selesai|lahir)/i.test(String(header || ''));
}

export function clamp(n, lo, hi) {
  return Math.max(lo, Math.min(hi, n));
}
