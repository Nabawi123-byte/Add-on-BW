# Panduan Fitur MASTER BW

Semua perintah bisa diketik di **Panel AI** dalam bahasa Indonesia sehari-hari, atau dipilih dari
**ribbon**, **menu klik kanan**, dan **palet perintah (Ctrl+Alt+Space)**. Setiap aksi yang mengubah
dokumen ditampilkan dulu di jendela **Pratinjau aksi**. Di sana Anda bisa mencentang/menghilangkan langkah dan
mengedit argumennya sebelum klik **Jalankan**.

---

## A. MICROSOFT WORD

### 1. Menulis dengan AI
| Fitur | Contoh perintah |
|---|---|
| Tulis bagian baru | "tulis latar belakang penelitian tentang stunting di desa, 4 paragraf, gaya akademik" |
| Olah teks terpilih | Ringkas · Parafrase · Perbaiki ejaan & tata bahasa · Formal · Akademik · Santai · Sederhanakan · Kembangkan · Lanjutkan · Jadikan poin · Terjemah EN/ID · Usulkan judul · Buat abstrak · Buat kesimpulan · Buat soal · Jadikan email |
| Ringkas seluruh dokumen | Dokumen panjang dipotong per bagian, diringkas, lalu digabung (map-reduce) |
| Tanya tentang teks | Klik kanan → MASTER BW → "Tanya AI tentang teks ini…" |

### 2. Dokumen & surat (26 templat)
- **Surat**: surat dinas, undangan, Surat Keputusan (menimbang–mengingat–memutuskan), surat tugas, surat keterangan, surat kuasa, surat pernyataan, surat permohonan, lamaran kerja, nota dinas/memo.
- **Rapat & laporan**: notulen rapat, berita acara, laporan kegiatan, proposal kegiatan (dengan susunan panitia & RAB), KAK/TOR, SOP.
- **Akademik**: makalah lengkap, kerangka skripsi (halaman awal s.d. BAB V), proposal penelitian, artikel jurnal (IMRaD + abstract), laporan PKL, laporan praktikum.
- **Pendidikan**: Modul Ajar Kurikulum Merdeka (informasi umum, komponen inti, asesmen, LKPD), RPP satu lembar, soal ujian + kunci + rubrik.
- **Pribadi**: daftar riwayat hidup (CV).

Bagian bertanda *AI* diisi otomatis bila opsi "AI mengisi bagian isi" dicentang. Data yang tidak Anda
berikan ditandai `[ … ]` dengan sorotan kuning untuk dilengkapi.

### 3. Format & struktur
- **Format baku sekali klik**: Skripsi/Proposal (TNR 12, spasi 2, margin 4-4-3-3), Makalah, Laporan, Jurnal, Surat Dinas (F4).
- **Penomoran bab otomatis**: `BAB I → A. → 1.` · `BAB 1 → 1.1 → 1.1.1` · `1. → 1.1 → 1.1.1`.
- Daftar isi, daftar gambar, daftar tabel, caption bernomor otomatis, perbarui semua field.
- Nomor halaman (arab/romawi, posisi bebas, mulai dari angka tertentu), header/footer.
- Huruf, paragraf, spasi, indentasi, margin, ukuran kertas (A4/F4/Letter/Legal), orientasi, kolom gaya jurnal.
- Tabel (gaya akademik APA, garis penuh, modern), teks → tabel, persamaan matematika, gambar, catatan kaki.
- **Rapikan dokumen**: spasi ganda, baris kosong, enter manual, tab, dan tanda kutip sisa salin dari PDF.

### 4. Akademik
- **Sitasi & daftar pustaka**: APA 7, IEEE, Harvard, MLA, Chicago (author-date), Vancouver. Pustaka disimpan otomatis untuk dipakai ulang. Untuk 3+ penulis dipakai "dkk.".
- **Rapikan daftar pustaka**: urut abjad + indentasi gantung 1,27 cm.
- **Cek ejaan EYD offline**: kata tidak baku (KBBI), *di-* awalan vs kata depan, partikel *pun*, spasi & tanda baca, kapital awal kalimat & nama hari, kata berulang, penulisan Rp dan %. Hasilnya berupa komentar, atau perbaikan otomatis untuk kesalahan yang aman.

### 5. Perkantoran
- Kop surat (dengan logo dari Pengaturan), blok tanda tangan (kanan/kiri/dua kolom, NIP/NIK/NIM), tanggal Indonesia.
- Watermark (RAHASIA/DRAFT/CONTOH), komentar, *track changes*, cari & ganti, sorot kata, tautan.
- **Surat massal**: tulis `{{Nama}}`, `{{Alamat}}`, … di dokumen, lalu perintahkan "gabungkan dengan data D:\data.xlsx". Hasilnya satu dokumen per baris data.
- Simpan PDF, statistik dokumen (kata, halaman, waktu baca).

---

## B. MICROSOFT EXCEL

### 1. Rumus
| Fitur | Contoh |
|---|---|
| Buat rumus dari bahasa biasa | "kolom G berisi total = qty × harga", "di H2 total penjualan wilayah Jawa bulan Maret" |
| Jelaskan rumus | Pilih sel → Jelaskan Rumus. Penjelasan per fungsi, tanpa AI |
| Perbaiki rumus | Diagnosis `#N/A #REF! #VALUE! #DIV/0! #NAME?` + usulan rumus baru (dikonfirmasi dulu) |
| Terbilang | Angka terpilih → terbilang rupiah di kolom kanannya |

Rumus selalu diperiksa: sintaks, kurung, dan **kompatibilitas Excel 2016**. Fungsi XLOOKUP, FILTER, UNIQUE,
IFS, TEXTJOIN, dan sejenisnya ditolak dan diganti alternatif yang tersedia di Excel 2016.

### 2. Olah & bersihkan data
- Bersihkan: spasi berlebih, karakter aneh, nama orang (gelar dipertahankan), HURUF BESAR/kecil, teks/Rupiah → angka, tanggal berantakan → tanggal Excel ("12 Januari 2024", "12-01-24", …), nomor HP → +62.
- Hapus duplikat, hapus baris kosong, pisah/gabung kolom, urutkan (2 kunci), filter, validasi data (dropdown), cari & ganti.
- Rekap per kategori dengan rumus SUMIFS/AVERAGEIFS/COUNTIFS yang ikut berubah saat data berubah.
- Format angka: Rupiah, akuntansi, ribuan, persen, tanggal panjang Indonesia ("1 Januari 2025").
- Format bersyarat (lebih dari, di antara, duplikat, 10 teratas, di atas rata-rata, skala warna, data bar, ikon, rumus), jadikan tabel, bekukan panel, pengaturan cetak (F4/A4, muat 1 halaman, ulangi judul).

### 3. Analisis & visualisasi
- **Analisis & insight AI**: statistik per kolom, rekap per kategori, lalu AI menulis temuan, tren, anomali, dan rekomendasi.
- **Dashboard otomatis**: 4 kartu KPI + PivotTable + grafik per kategori + tren bulanan.
- PivotTable, grafik (kolom, batang, garis, lingkaran, donat, area, sebar, kombinasi).

### 4. Statistik penelitian (setara SPSS, tanpa AI)
| Uji | Keluaran |
|---|---|
| Statistik deskriptif | N, mean, median, modus, SD, varians, min–max, kuartil, CV, skewness, kurtosis |
| Uji validitas & reliabilitas | r hitung vs r tabel per item, r item-total terkoreksi, alpha jika item dihapus, Cronbach's Alpha + kategori |
| Uji normalitas | Kolmogorov-Smirnov (Lilliefors), Shapiro-Wilk, Jarque-Bera + kesimpulan |
| Korelasi | Matriks Pearson/Spearman + Sig. (2-tailed) + kekuatan hubungan |
| Regresi linear sederhana/berganda | R, R², adj. R², uji F, koefisien + uji t, VIF (multikolinearitas), Durbin-Watson, normalitas residual, persamaan regresi |
| Uji t | Satu sampel, independen (Levene + Welch, Cohen's d), berpasangan (pretest–posttest) |
| ANOVA satu jalur | Tabel ANOVA, Levene, eta² |
| Chi-square | Frekuensi harapan, χ², Sig., Cramer's V, peringatan sel < 5 |
| Outlier | IQR / z-score, sel ditandai warna |
| Peramalan | Tren linear, rata-rata bergerak, pemulusan eksponensial + MAPE + grafik |

Setiap hasil dilengkapi **interpretasi bahasa Indonesia** yang siap disalin ke BAB IV.

### 5. Keuangan & pajak
- Tabel angsuran kredit: anuitas, bunga flat, bunga efektif/menurun.
- Penyusutan: garis lurus, saldo menurun, saldo menurun ganda, jumlah angka tahun.
- Kelayakan investasi: NPV, IRR, Profitability Index, Payback & Discounted Payback.
- Pajak (estimasi): PPh 21 tahunan (Pasal 17 + PTKP + biaya jabatan), PPN 12% DPP 11/12, PPh 23, PPh Final UMKM 0,5%.
- Titik impas (BEP), rasio keuangan (likuiditas, solvabilitas, profitabilitas, aktivitas).

### 6. Templat lembar kerja (18)
Anggaran vs realisasi · RAB (PPN + terbilang) · arus kas · BKU sekolah/BOS · invoice · kuitansi · daftar hadir
(tanda tangan zig-zag) · absensi bulanan H/S/I/A/C · daftar nilai/rapor (bobot, predikat, peringkat, KKM) ·
daftar gaji (BPJS JHT/JP/Kesehatan) · inventaris · stok barang (pesan ulang) · data penjualan · piutang (umur &
status) · jadwal kegiatan · tabulasi kuesioner Likert · laporan laba rugi · neraca (cek seimbang).

### 7. Fungsi sel (ketik langsung di sel)
| Fungsi | Contoh |
|---|---|
| `=AI(perintah; data…)` | `=AI("Buat slogan singkat untuk produk"; A2)` |
| `=AI.RINGKAS(teks; [kata])` | `=AI.RINGKAS(B2; 15)` |
| `=AI.KATEGORI(teks; kategori)` | `=AI.KATEGORI(B2; "Makanan,Minuman,Lainnya")` |
| `=AI.SENTIMEN(teks)` | Positif / Negatif / Netral |
| `=AI.TERJEMAH(teks; [bahasa])` | `=AI.TERJEMAH(A2; "Inggris")` |
| `=AI.EKSTRAK(teks; apa)` | `=AI.EKSTRAK(A2; "nomor HP")` |
| `=AI.PERBAIKI(teks)` | Perbaikan EYD |
| `=TERBILANG(angka)` / `=TERBILANG.RUPIAH(angka)` | Satu Juta Lima Ratus Ribu Rupiah |
| `=TANGGAL.INDO(tgl; [0/1/2])`, `=HARI.INDO(tgl)`, `=BULAN.INDO(tgl)` | Senin, 1 Januari 2025 |
| `=ROMAWI(n)`, `=ANGKA.INDO(teks)`, `=TANGGAL.TEKS(teks)` | Konversi format Indonesia |
| `=NAMA.RAPI(nama)`, `=HP.RAPI(no)`, `=USIA(tgl_lahir)`, `=PREDIKAT(nilai)` | Utilitas administrasi |
| `=PPH21.TAHUNAN(bruto; "K/1")`, `=PTKP(status)`, `=PPN.HITUNG(harga)`, `=ANGSURAN(pokok; bunga; tenor)` | Keuangan |

Hasil fungsi AI disimpan di cache, jadi tidak dihitung ulang setiap Excel melakukan kalkulasi.

### 8. Otomasi VBA
Untuk tugas di luar daftar aksi, minta "buat makro untuk …". AI menulis kode VBA, kode diperiksa keamanannya
(hapus file, akses internet, registry, Shell, dll.), lalu ditampilkan dan dipasang hanya bila Anda setuju.

---

## C. PENGATURAN PENTING (tab MASTER BW → Pengaturan)

| Pengaturan | Keterangan |
|---|---|
| `ModelFile` | Nama berkas model di folder `models` (mis. `Qwen3-4B-Q4_K_M.gguf`) |
| `ContextSize` | Panjang konteks (default 8192). Turunkan ke 4096 bila RAM terbatas |
| `Threads` | 0 = otomatis |
| `DefaultReasoning` | Cepat / Seimbang / Mendalam |
| `ThinkBudgetSeimbang`, `ThinkBudgetMendalam` | Batas token "berpikir" agar tidak terlalu lama di CPU |
| `ConfirmBeforeExecute` | Selalu tampilkan pratinjau (default: ya) |
| `PaletteHotkey` | Default `Ctrl+Alt+Space` |
| `UserName`, `UserTitle`, `UserNip`, `Institution`, `InstitutionAddress`, `City`, `LogoPath` | Untuk kop, tanda tangan, sampul |
| `CitationStyle` | Gaya sitasi default (APA7) |
| `IdleUnloadMinutes` | Model dilepas dari RAM setelah sekian menit tidak dipakai |

Pengaturan dan data tersimpan di `%LOCALAPPDATA%\MasterBW` (settings.json, referensi.json, cache, logs).
