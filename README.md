# Master — asisten perintah untuk Microsoft Word & Excel

Master adalah add-in Office. Anda cukup **mengetik perintah biasa** di panel samping, dan Master langsung mengerjakannya di dokumen: membuat tabel, laporan keuangan, simulasi cicilan, surat resmi, notulen, merapikan format, dan lain-lain.

- **Offline (tanpa internet, gratis):** 16 template Excel, 18 template surat/dokumen Word, dan lebih dari 50 perintah olah data & format.
- **Online (AI Claude):** perintah bebas apa saja, misalnya "analisis data penjualan ini", "buat rumus komisi 5% bila di atas 10 juta", "ringkas dokumen ini jadi 5 poin", "perbaiki tata bahasa paragraf terpilih". AI membaca dokumen lalu mengubahnya langsung lewat alat-alat Master.
- **Mode Otomatis:** perintah yang dikenali dijalankan offline (cepat, tanpa biaya), sisanya dikirim ke Claude. Saat internet putus, perintah AI disimpan di **antrean** dan bisa dijalankan begitu internet kembali.

---

## Contoh perintah

### Excel
| Perintah | Hasil |
|---|---|
| `simulasi cicilan pinjaman 50 juta bunga 10% 24 bulan` | Sheet jadwal angsuran anuitas (rumus PMT), total bunga, grafik sisa pinjaman |
| `cicilan kpr 300 juta bunga 7,5% 15 tahun flat` | Simulasi bunga flat |
| `buat anggaran bulanan` | Rencana vs realisasi, sisa dana, status MELEBIHI, grafik |
| `buat laporan arus kas saldo awal 5 juta` | Kas masuk/keluar + saldo berjalan |
| `buat laporan laba rugi` / `buat neraca` | Laporan keuangan dengan rumus hidup + cek "SEIMBANG" |
| `buat daftar gaji 15 karyawan` | Gaji kotor, potongan BPJS, gaji bersih |
| `buat faktur` | Invoice dengan diskon, PPN 11%, total, dan terbilang |
| `buat absensi bulan desember` | Daftar hadir H/I/S/A + rekap & % hadir |
| `buat tabel karyawan dengan kolom Nama, Jabatan, Gaji 10 baris` | Tabel Excel bergaya, kolom uang otomatis format Rupiah |
| `jumlahkan` · `rata-rata` · `ranking` · `persentase dari total` | Baris/kolom rumus di bawah atau di samping data |
| `urutkan berdasarkan Gaji turun` · `filter` · `hapus duplikat` | Olah data |
| `rekap total per Produk` | Tabel rekap SUMIF + grafik |
| `format rupiah` · `bekukan header` · `beri garis` · `rapikan kolom` | Format |
| `tandai nilai di atas 1 juta merah` · `tandai duplikat kuning` | Format bersyarat |
| `grafik batang` · `grafik lingkaran judul Komposisi Biaya` | Grafik dari data |
| `terbilang 1.250.000` | "Satu juta dua ratus lima puluh ribu rupiah" |

Fungsi sel (Excel 365): `=MASTER.TERBILANG(A1)` dan `=MASTER.RUPIAH(A1)`.

### Word
| Perintah | Hasil |
|---|---|
| `buat surat izin sakit` | Formulir terisi otomatis dari Profil → klik Sisipkan |
| `surat lamaran kerja posisi: Admin, perusahaan: PT Maju` | Nilai `kunci: nilai` langsung mengisi formulir |
| `surat undangan rapat`, `surat kuasa`, `surat pernyataan`, `surat keterangan kerja`, `surat pengunduran diri`, `permohonan cuti`, `surat tugas`, `surat edaran`, `surat penawaran harga`, `surat peringatan sp1`, `surat rekomendasi` | Surat resmi format Indonesia |
| `buat notulen rapat` · `buat memo` · `buat berita acara` · `buat kwitansi` | Dokumen kantor (kwitansi lengkap dengan terbilang) |
| `ganti "PT Lama" dengan "PT Baru"` · `sorot kata penting kuning` | Edit cepat |
| `font Times New Roman ukuran 12 semua` · `spasi 1,5 semua` · `rata kiri kanan semua` | Format seluruh dokumen |
| `daftar isi` · `nomor halaman` · `header PT Maju Jaya` · `halaman baru` | Struktur dokumen |
| `hapus spasi ganda` · `hapus paragraf kosong` · `rapikan dokumen` · `hitung kata` | Perapian |

**Rantai perintah:** `buat laporan arus kas lalu bekukan header` atau pisahkan dengan `;` / baris baru.
**Makro:** simpan beberapa perintah dengan satu nama di tab *Makro*, lalu jalankan dengan `jalankan makro <nama>`.
Ketik `bantuan` di panel untuk melihat semua perintah.

---

## Kebutuhan

- **Microsoft 365** atau **Office 2021/2024** (Windows atau Mac). Office 2016/2019 versi lama memakai mesin browser lama dan kemungkinan tidak didukung.
- **Node.js LTS** (https://nodejs.org), untuk menjalankan server lokal Master.
- Internet **hanya** untuk instalasi pertama dan untuk fitur AI. Perintah offline, template, dan makro berjalan tanpa internet.
- Untuk AI: **API key Anthropic** dari https://console.anthropic.com (biayanya ditagih terpisah dari langganan Claude.ai/Claude Pro).

## Instalasi

### Windows
1. Unduh repositori ini (Code → Download ZIP), lalu ekstrak, misalnya ke `C:\Master`.
2. Klik dua kali **`install-windows.bat`**. Klik **Yes** bila Windows meminta izin memasang sertifikat.
3. Tutup lalu buka lagi Excel/Word. Tombol **Master** ada di tab **Beranda (Home)**.
   Jika belum muncul: Beranda → Add-in → Add-in Saya → *Add-in Pengembang* → Master.

Server Master otomatis berjalan setiap kali Windows menyala. Untuk menjalankannya manual: `start-master.bat`. Untuk menghapus: `uninstall-windows.bat`.

### macOS
```bash
cd ~/Downloads/Master
```
```bash
./install-mac.sh
```
Lalu di Excel/Word: Sisipkan (Insert) → Add-in Saya (My Add-ins) → Master.

### Mengaktifkan AI
Buka panel Master → tab **Pengaturan** → tempel API key → **Uji koneksi** → **Simpan**. Di sana juga ada:
- **Model:** Claude Opus 5.5 (default, paling cerdas), Sonnet 5.5 (lebih cepat & hemat), atau Haiku 5.5 (paling hemat).
- **Tingkat berpikir:** rendah/sedang/tinggi. Makin tinggi, makin teliti tetapi makin mahal.
- **Kode khusus dari AI:** bila alat bawaan tidak cukup, AI dapat menulis kode Office.js sendiri. Defaultnya Master **meminta persetujuan Anda** sebelum menjalankan kode itu.
- **Profil:** nama, jabatan, instansi, alamat, dan kota dipakai untuk mengisi surat secara otomatis.

Untuk Opus 5.5 dan Sonnet 5.5, Master mengaktifkan *fallback server-side* Anthropic: jika sebuah permintaan ditolak filter keamanan model, permintaan itu otomatis dicoba ulang di model cadangan.

---

## Cara kerja

```
Panel Master (taskpane)
 ├─ Parser perintah offline (bahasa Indonesia) ── template & aksi ──┐
 ├─ Agen AI Claude (streaming + alat) ─────────── alat Office ─────┤
 └─ Antrean, makro, riwayat (disimpan lokal)                       │
                                                                   ▼
                                    Office.js → Excel / Word (dokumen Anda)
Server lokal https://localhost:3000 menyajikan panel + office.js salinan lokal
```

Alat AI: `excel_info`, `excel_read`, `excel_apply` (tulis, format, tabel, grafik, format bersyarat, dropdown, urut, filter, sheet), `word_read`, `word_insert` (judul, paragraf, daftar, tabel, blok data surat, tanda tangan), `word_edit` (ganti, sorot, format, header/footer, nomor halaman, daftar isi), `master_command` (memakai template bawaan), dan `run_office_js` (kode khusus, dengan persetujuan).

## Keamanan & privasi
- API key disimpan di penyimpanan lokal panel pada komputer Anda dan dikirim langsung ke `api.anthropic.com`. Jangan pakai Master di komputer bersama bila key Anda sensitif.
- Saat AI dipakai, isi dokumen yang relevan (sel terpilih, data yang dibaca AI) dikirim ke Anthropic untuk diproses. Mode **Offline saja** tidak mengirim apa pun.
- Server lokal hanya mendengarkan `localhost`, sehingga tidak bisa diakses dari komputer lain.

## Batasan yang perlu diketahui
- Di Excel, sebagian perubahan dari add-in **tidak bisa di-undo** dengan Ctrl+Z. Simpan file sebelum menjalankan perintah besar.
- Daftar isi perlu diperbarui manual (klik kanan → *Update Field* / F9) dan hanya membaca paragraf bergaya Judul/Heading.
- Jadwal cicilan dibuat sesuai tenor saat perintah dijalankan. Bila tenor diubah, jalankan perintahnya lagi.
- Pengujian otomatis memakai tiruan (simulasi) Office.js dan API Claude. Uji di Word/Excel asli tetap disarankan. Bila ada perintah yang error, catat pesannya dan laporkan.

## Pengembangan
```bash
npm install
```
```bash
npm test
```
```bash
npm run test:e2e
```
```bash
npm start
```
- `npm test`: uji unit (parser, terbilang, angka gaya Indonesia, semua template).
- `npm run test:e2e`: uji panel di Chromium dengan Office.js tiruan dan API Claude tiruan.
- `npm start`: server https://localhost:3000.
- `npm run manifest -- --port 3001`: membuat ulang manifest bila port diganti (jalankan server dengan `MASTER_PORT=3001`).
- `npm run sideload:excel` / `npm run sideload:word`: buka Office dengan add-in termuat.

Struktur:
```
src/core/     parser perintah, template Excel & Word, terbilang, agen AI, validator
src/host/     eksekutor Office.js untuk Excel dan Word
src/ui/       panel (HTML/CSS/JS), ikon
src/functions fungsi sel =MASTER.TERBILANG / =MASTER.RUPIAH
server.js     server HTTPS lokal (melayani dist/ + office.js lokal)
manifest-*.xml manifest add-in Excel & Word
```
