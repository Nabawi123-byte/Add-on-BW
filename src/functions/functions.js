/* global CustomFunctions */
// Fungsi khusus Excel (offline): =MASTER.TERBILANG(A1) dan =MASTER.RUPIAH(A1)

import { terbilang, terbilangRupiah } from '../core/terbilang.js';
import { formatRupiah } from '../core/util.js';

/** Mengubah angka menjadi kata-kata bahasa Indonesia. rupiah=TRUE menambahkan kata "rupiah". */
function TERBILANG(angka, rupiah) {
  if (rupiah === null || rupiah === undefined || rupiah) return terbilangRupiah(angka);
  const t = terbilang(angka);
  return t.charAt(0).toUpperCase() + t.slice(1);
}

/** Menampilkan angka sebagai teks Rupiah, mis. "Rp 1.500.000". */
function RUPIAH(angka) {
  return formatRupiah(angka);
}

CustomFunctions.associate('TERBILANG', TERBILANG);
CustomFunctions.associate('RUPIAH', RUPIAH);
