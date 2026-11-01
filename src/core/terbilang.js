// Terbilang: mengubah angka menjadi kata-kata bahasa Indonesia.

const SATUAN = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];

function words(n) {
  if (n < 12) return SATUAN[n];
  if (n < 20) return words(n - 10) + ' belas';
  if (n < 100) return words(Math.floor(n / 10)) + ' puluh' + (n % 10 ? ' ' + words(n % 10) : '');
  if (n < 200) return 'seratus' + (n - 100 ? ' ' + words(n - 100) : '');
  if (n < 1000) return words(Math.floor(n / 100)) + ' ratus' + (n % 100 ? ' ' + words(n % 100) : '');
  if (n < 2000) return 'seribu' + (n - 1000 ? ' ' + words(n - 1000) : '');
  const units = [
    [1e12, 'triliun'],
    [1e9, 'miliar'],
    [1e6, 'juta'],
    [1e3, 'ribu'],
  ];
  for (const [size, name] of units) {
    if (n >= size) {
      const head = Math.floor(n / size);
      const rest = n - head * size;
      return words(head) + ' ' + name + (rest ? ' ' + words(rest) : '');
    }
  }
  return '';
}

/**
 * terbilang(1500000) -> "satu juta lima ratus ribu"
 * terbilang(12.5) -> "dua belas koma lima"
 */
export function terbilang(value) {
  let n = Number(value);
  if (!Number.isFinite(n)) return '';
  if (Math.abs(n) >= 1e15) return 'angka terlalu besar';
  const negative = n < 0;
  n = Math.abs(n);
  const intPart = Math.floor(n);
  let text = intPart === 0 ? 'nol' : words(intPart);
  const decimals = n.toFixed(2).split('.')[1].replace(/0+$/, '');
  if (decimals) {
    text += ' koma ' + decimals.split('').map((d) => (d === '0' ? 'nol' : SATUAN[Number(d)])).join(' ');
  }
  return (negative ? 'minus ' : '') + text.replace(/\s+/g, ' ').trim();
}

/** terbilangRupiah(1500000) -> "Satu juta lima ratus ribu rupiah" */
export function terbilangRupiah(value) {
  const n = Math.round(Number(value) || 0);
  const t = terbilang(n);
  return t.charAt(0).toUpperCase() + t.slice(1) + ' rupiah';
}
