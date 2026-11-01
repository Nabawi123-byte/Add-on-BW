# MASTER BW — Asisten AI Offline untuk Word & Excel

MASTER BW adalah add-in untuk **Microsoft Word dan Excel 2016 ke atas** yang membawa asisten AI
**100% offline**. Model AI berjalan di komputer Anda sendiri lewat llama.cpp, tanpa internet, tanpa
langganan, dan data tidak pernah keluar dari laptop. Perintah ditulis dalam **bahasa Indonesia**, dan setiap
perubahan ditampilkan dulu untuk Anda setujui.

| | Word | Excel |
|---|---|---|
| Perintah bahasa Indonesia → aksi | 46 aksi (format, struktur, tabel, sitasi, surat…) | 52 aksi (rumus, data, pivot, grafik, statistik, keuangan…) |
| Templat siap pakai | 26 templat: surat dinas, undangan, SK, surat tugas, notulen, berita acara, proposal, KAK, SOP, CV, makalah, skripsi, artikel jurnal, laporan PKL, modul ajar, RPP, soal ujian… | 18 templat: anggaran, RAB, arus kas, BKU sekolah/BOS, invoice, kuitansi, daftar hadir, absensi, nilai siswa/rapor, payroll, inventaris, stok, piutang, laba rugi, neraca, kuesioner… |
| Tanpa AI (instan) | Cek ejaan EYD, daftar pustaka APA/IEEE/Harvard/MLA/Chicago/Vancouver, format skripsi, penomoran BAB | Uji statistik setara SPSS, kalkulator pajak & kredit, terbilang, pembersih data |
| Fungsi di sel | — | `=AI()`, `=AI.RINGKAS()`, `=AI.KATEGORI()`, `=TERBILANG.RUPIAH()`, `=TANGGAL.INDO()`, `=PPH21.TAHUNAN()` … |

Daftar fitur lengkap dan contoh perintah ada di **[docs/PANDUAN.md](docs/PANDUAN.md)**.

---

## 1. Kebutuhan

- Windows 10/11 64-bit, Microsoft Office **2016, 2019, 2021, atau 365** (32 atau 64-bit)
- .NET Framework 4.8 (biasanya sudah ada di Windows 10/11)
- RAM 8 GB atau lebih (direkomendasikan 10 GB+), ruang disk ±3 GB untuk model
- Internet **hanya sekali** untuk mengunduh model AI (±1–2,5 GB)

## 2. Cara memasang (tanpa perlu build)

1. Buka tab **Actions** di repositori GitHub ini, klik run **Build MASTER BW** yang paling baru (centang hijau).
2. Di bagian **Artifacts**, unduh **MasterBW-paket-instalasi** (±55 MB), lalu ekstrak ZIP-nya.
3. **Tutup Word dan Excel**, lalu klik ganda **`INSTALL.cmd`**.
4. Ikuti petunjuk di layar:
   - isi nama, instansi, dan kota (dipakai untuk kop & tanda tangan surat; boleh dikosongkan),
   - pilih model AI:
     - **1 = Qwen3-4B** (±2,5 GB): paling pintar, mendukung mode nalar mendalam. *Direkomendasikan untuk RAM 10 GB.*
     - **2 = Qwen3-1.7B** (±1,1 GB): lebih ringan dan cepat.
   - opsional: model suara Whisper (±150 MB) untuk perintah suara.
5. Buka Word atau Excel. Tab **MASTER BW** akan muncul di ribbon.

Pemasangan berlaku per pengguna dan **tidak butuh hak Administrator**. Berkas dipasang ke
`%LOCALAPPDATA%\Programs\MasterBW`. Untuk mencopot, jalankan `UNINSTALL.cmd`.

> Jika unduhan model gagal (misalnya jaringan kantor memblokir HuggingFace), unduh manual file
> `Qwen3-4B-Q4_K_M.gguf` dari https://huggingface.co/Qwen/Qwen3-4B-GGUF lalu salin ke
> `%LOCALAPPDATA%\Programs\MasterBW\models`. Model GGUF lain (Gemma, Llama, Qwen) juga bisa dipakai.
> Atur nama berkasnya di **Pengaturan → ModelFile**.

## 3. Cara memakai

- **Panel AI** (tab MASTER BW → Panel AI): ketik perintah, misalnya
  *"buat surat undangan rapat komite hari Senin jam 9 di aula"* atau
  *"tambah kolom total = qty × harga lalu format rupiah"*.
  AI menyusun rencana aksi → Anda periksa di jendela pratinjau → klik **Jalankan**.
- **Palet perintah**: tekan **Ctrl+Alt+Space** di mana saja, ketik nama perintah, lalu Enter.
- **Klik kanan** pada teks/sel → menu **MASTER BW** (ringkas, parafrase, jelaskan rumus, terbilang…).
- **Mode nalar**: ⚡ Cepat (tanpa berpikir), ⚖ Seimbang, 🧠 Mendalam (*high reasoning*, lebih lambat tetapi lebih teliti).
- **Batalkan**: tombol ↶ di panel. Di Word bisa juga dengan Ctrl+Z (satu langkah untuk seluruh aksi AI).

Perintah pertama setelah Office dibuka butuh ±10–30 detik untuk memuat model ke RAM. Perintah
berikutnya jauh lebih cepat karena bagian prompt yang sama di-*cache*.

## 4. Arsitektur singkat

```
Word / Excel ── MasterBW.Office.dll (COM add-in: ribbon, panel chat, eksekutor aksi)
      │                   │  HTTP 127.0.0.1 + token
      │                   ▼
      │         MasterBW.Engine.exe (llama.cpp via LLamaSharp + Whisper)  ←  models\*.gguf
      │                   ▲
Excel ── MasterBW.Functions.xll (Excel-DNA: =AI(), =TERBILANG(), …)
                MasterBW.Core.dll (dipakai semua: prompt, katalog aksi, EYD, statistik, keuangan, templat)
```

- **Satu mesin AI dipakai bersama** oleh Word, Excel, dan fungsi sel, jadi model hanya dimuat sekali di RAM. Mesin dilepas otomatis dari RAM setelah 20 menit tidak dipakai.
- **Grammar GBNF** memaksa model kecil mengeluarkan JSON rencana aksi yang valid dan hanya memakai nama aksi yang ada. Ini kunci keandalan model 1,7–4B.
- **Pemilihan aksi relevan** membuat prompt tetap pendek sehingga cepat di CPU.
- Fitur deterministik (statistik, pajak, EYD, sitasi, templat) **tidak memakai AI**, sehingga hasilnya pasti dan instan. Uji statistik diverifikasi terhadap SciPy/statsmodels.

## 5. Membangun dari sumber

Butuh .NET SDK 10. Bisa di Windows, Linux, atau macOS:

```bash
./scripts/build-release.sh          # uji + build semua + paket dist/MasterBW-<versi>.zip
dotnet test tests/MasterBW.Core.Tests
```

Di Windows dengan Visual Studio 2022: buka `MasterBW.sln`, lalu build konfigurasi Release.
Struktur proyek:

| Proyek | Isi |
|---|---|
| `src/MasterBW.Core` | netstandard2.0. Prompt, katalog aksi, grammar, EYD, terbilang, statistik, keuangan, pajak, sitasi, templat |
| `src/MasterBW.Engine` | net10.0. Server AI lokal (LLamaSharp + Whisper.net) |
| `src/MasterBW.Office` | net48. COM add-in Word & Excel: ribbon, panel, eksekutor aksi |
| `src/MasterBW.Functions` | net48. Fungsi sel Excel (Excel-DNA) |
| `tests/MasterBW.Core.Tests` | Uji otomatis (64 uji) |

## 6. Mengatasi masalah

| Gejala | Solusi |
|---|---|
| Tab MASTER BW tidak muncul | File → Options → Add-ins → Manage: **COM Add-ins** → Go → centang MASTER BW. Jika ada di *Disabled Items*, aktifkan lagi. Jalankan ulang `INSTALL.cmd`. |
| Error *"type initializer for LLama.Native.NativeApi"* / *"tidak bisa memuat llama.dll"* | Pasang **Microsoft Visual C++ Redistributable x64**: https://aka.ms/vs/17/release/vc_redist.x64.exe, lalu tutup Word & Excel, akhiri proses `MasterBW.Engine` di Task Manager, dan buka lagi. Paket versi terbaru sudah menyertakan DLL ini. |
| "Mesin AI tidak merespons" | Pengaturan → **Mulai ulang mesin AI**. Lihat log di `%LOCALAPPDATA%\MasterBW\logs`. |
| "File model tidak ditemukan" | Pengaturan → **Unduh model AI**, atau salin `.gguf` ke folder models. |
| Lambat | Gunakan mode ⚡ Cepat atau model 1.7B. Tutup aplikasi berat lain. Kecilkan `ContextSize` ke 4096. |
| Fungsi `=AI()` menampilkan `#NAME?` | Excel → File → Options → Add-ins → Excel Add-ins → Browse → `addin\functions\MasterBW.Functions-AddIn64-packed.xll` (atau tanpa "64" untuk Office 32-bit). |
| Makro VBA tidak bisa dipasang otomatis | Trust Center → Macro Settings → centang *Trust access to the VBA project object model*. Kode tetap disalin ke clipboard. |

## 7. Catatan

- Perhitungan pajak (PPh 21, PPN, PPh 23, PPh Final UMKM) adalah **estimasi** berdasarkan UU HPP. Selalu verifikasi dengan ketentuan DJP terbaru.
- AI bisa keliru. Itu sebabnya semua aksi melewati pratinjau. Periksa kembali isi dokumen penting.
- Lisensi model: Qwen3 dirilis dengan lisensi Apache 2.0.
