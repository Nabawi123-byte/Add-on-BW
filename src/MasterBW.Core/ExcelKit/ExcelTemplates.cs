using System;
using System.Collections.Generic;
using System.Linq;

namespace MasterBW.Core.ExcelKit
{
    /// <summary>Kolom pada templat tabel. Formula memakai {r} = baris saat ini, {p} = baris sebelumnya, {first}/{last} = baris data pertama/terakhir.</summary>
    public sealed class TemplateColumn
    {
        public string Header;
        public double Width = 14;
        /// <summary>teks | angka | rupiah | persen | tanggal | desimal</summary>
        public string Format = "teks";
        public string Formula;
        /// <summary>Rumus khusus untuk baris data pertama (mis. saldo awal).</summary>
        public string FirstRowFormula;
        public string[] Choices;
        /// <summary>Rumus baris total, mis. "=SUM({col}{first}:{col}{last})". Kosong = tidak ada.</summary>
        public string Total;
        public bool Number;
    }

    /// <summary>Baris pada templat laporan berbentuk daftar (laba rugi, neraca, kuitansi). Rumus memakai {key} → alamat nilai baris lain.</summary>
    public sealed class TemplateLine
    {
        public string Key;
        public string Label;
        public string Value;
        /// <summary>heading | input | formula | total | note | blank</summary>
        public string Kind = "input";
        public string Format = "rupiah";
    }

    public sealed class SheetTemplate
    {
        public string Id;
        public string Name;
        public string Category;
        public string Description;
        public string Title;
        public List<string> InfoRows = new List<string>();
        public List<TemplateColumn> Columns = new List<TemplateColumn>();
        public List<TemplateLine> Lines = new List<TemplateLine>();
        public int DefaultRows = 20;
        public bool AutoNumber = true;
        public List<(string label, string formula, string format)> Summary = new List<(string, string, string)>();
        public bool IsStatement => Lines.Count > 0;
    }

    public static class ExcelTemplates
    {
        private static TemplateColumn C(string header, double width = 14, string format = "teks", string formula = null, string total = null)
            => new TemplateColumn { Header = header, Width = width, Format = format, Formula = formula, Total = total, Number = format != "teks" && format != "tanggal" };

        private static TemplateColumn Sum(string header, double width = 16, string format = "rupiah", string formula = null)
            => C(header, width, format, formula, "=SUM({col}{first}:{col}{last})");

        private static TemplateLine L(string key, string label, string kind = "input", string value = null, string format = "rupiah")
            => new TemplateLine { Key = key, Label = label, Kind = kind, Value = value, Format = format };

        public static readonly List<SheetTemplate> All = new List<SheetTemplate>
        {
            new SheetTemplate
            {
                Id = "anggaran", Name = "Anggaran vs Realisasi", Category = "Keuangan", Title = "ANGGARAN DAN REALISASI",
                Description = "Pantau anggaran per pos dengan selisih dan persentase serapan.",
                InfoRows = { "Periode:", "Unit/Bagian:" },
                Columns = { C("No", 5), C("Kategori", 18), C("Uraian", 34), Sum("Anggaran"), Sum("Realisasi"),
                            Sum("Selisih", 16, "rupiah", "=D{r}-E{r}"), C("Serapan", 10, "persen", "=IF(D{r}=0,0,E{r}/D{r})", "=IF(D{last+1}=0,0,E{last+1}/D{last+1})") },
            },
            new SheetTemplate
            {
                Id = "rab", Name = "Rencana Anggaran Biaya (RAB)", Category = "Keuangan", Title = "RENCANA ANGGARAN BIAYA",
                Description = "RAB kegiatan/proyek dengan volume × harga satuan, PPN, dan terbilang.",
                InfoRows = { "Kegiatan:", "Lokasi:", "Tahun Anggaran:" },
                Columns = { C("No", 5), C("Uraian Pekerjaan/Barang", 38), C("Volume", 9, "desimal"), C("Satuan", 9), C("Harga Satuan", 16, "rupiah"), Sum("Jumlah", 18, "rupiah", "=C{r}*E{r}") },
                Summary = { ("PPN 11%", "=ROUND(F{last+1}*11%,0)", "rupiah"), ("TOTAL", "=F{last+1}+F{last+2}", "rupiah"), ("Dibulatkan", "=ROUNDDOWN(F{last+3},-3)", "rupiah"), ("Terbilang", "=TERBILANG.RUPIAH(F{last+4})", "teks") },
            },
            new SheetTemplate
            {
                Id = "arus_kas", Name = "Arus Kas / Buku Kas", Category = "Keuangan", Title = "LAPORAN ARUS KAS",
                Description = "Catat kas masuk dan keluar dengan saldo berjalan otomatis.",
                InfoRows = { "Periode:", "Saldo awal:" },
                Columns = { C("No", 5), C("Tanggal", 12, "tanggal"), C("Keterangan", 34), C("Kategori", 14), Sum("Kas Masuk"), Sum("Kas Keluar"),
                            new TemplateColumn { Header = "Saldo", Width = 18, Format = "rupiah", Number = true, Formula = "=G{p}+E{r}-F{r}", FirstRowFormula = "=N(C3)+E{r}-F{r}" } },
            },
            new SheetTemplate
            {
                Id = "buku_kas_umum", Name = "Buku Kas Umum (BKU) Sekolah/BOS", Category = "Keuangan", Title = "BUKU KAS UMUM",
                Description = "BKU format sekolah/instansi: tanggal, no. bukti, kode rekening, penerimaan, pengeluaran, saldo.",
                InfoRows = { "Nama Satuan:", "Bulan:", "Saldo bulan lalu:" },
                Columns = { C("No", 5), C("Tanggal", 12, "tanggal"), C("No. Bukti", 12), C("Kode Rekening", 14), C("Uraian", 34), Sum("Penerimaan"), Sum("Pengeluaran"),
                            new TemplateColumn { Header = "Saldo", Width = 18, Format = "rupiah", Number = true, Formula = "=H{p}+F{r}-G{r}", FirstRowFormula = "=N(C4)+F{r}-G{r}" } },
            },
            new SheetTemplate
            {
                Id = "invoice", Name = "Invoice / Faktur", Category = "Kantor", Title = "INVOICE", AutoNumber = true, DefaultRows = 10,
                Description = "Faktur penjualan dengan diskon, PPN, total, dan terbilang.",
                InfoRows = { "No. Invoice:", "Tanggal:", "Kepada:", "Alamat:", "Jatuh tempo:" },
                Columns = { C("No", 5), C("Deskripsi", 36), C("Qty", 8, "desimal"), C("Harga", 16, "rupiah"), C("Diskon", 9, "persen"), Sum("Jumlah", 18, "rupiah", "=C{r}*D{r}*(1-E{r})") },
                Summary = { ("PPN 11%", "=ROUND(F{last+1}*11%,0)", "rupiah"), ("TOTAL TAGIHAN", "=F{last+1}+F{last+2}", "rupiah"), ("Terbilang", "=TERBILANG.RUPIAH(F{last+3})", "teks") },
            },
            new SheetTemplate
            {
                Id = "daftar_hadir", Name = "Daftar Hadir", Category = "Kantor", Title = "DAFTAR HADIR", DefaultRows = 30,
                Description = "Daftar hadir rapat/kegiatan dengan kolom tanda tangan zig-zag.",
                InfoRows = { "Kegiatan:", "Hari/Tanggal:", "Tempat:" },
                Columns = { C("No", 5), C("Nama", 30), C("NIP/NIM/Instansi", 22), C("Jabatan/Kelas", 18), C("Tanda Tangan", 14, "teks", "=IF(MOD(A{r},2)=1,A{r}&\".\",\"\")"), C("", 14, "teks", "=IF(MOD(A{r},2)=0,A{r}&\".\",\"\")") },
            },
            new SheetTemplate
            {
                Id = "absensi_bulanan", Name = "Absensi Bulanan", Category = "Kantor", Title = "REKAP ABSENSI BULANAN", DefaultRows = 25,
                Description = "Absensi harian 1–31 (H/S/I/A/C) dengan rekap otomatis per orang.",
                InfoRows = { "Bulan:", "Unit/Kelas:" },
                Columns = BuildAttendanceColumns(),
            },
            new SheetTemplate
            {
                Id = "nilai_siswa", Name = "Daftar Nilai Siswa / Rapor", Category = "Pendidikan", Title = "DAFTAR NILAI PESERTA DIDIK", DefaultRows = 32,
                Description = "Nilai tugas, UH, PTS, PAS dengan bobot, predikat A–D, peringkat, dan ketuntasan KKM/KKTP.",
                InfoRows = { "Mata Pelajaran:", "Kelas/Semester:", "KKM/KKTP:", "Bobot (Tugas;UH;PTS;PAS): 20%;30%;20%;30%" },
                Columns = { C("No", 5), C("NIS/NISN", 12), C("Nama Peserta Didik", 30), C("Tugas", 8, "desimal"), C("UH", 8, "desimal"), C("PTS", 8, "desimal"), C("PAS", 8, "desimal"),
                            C("Nilai Akhir", 11, "desimal", "=IF(COUNT(D{r}:G{r})=0,\"\",ROUND(D{r}*20%+E{r}*30%+F{r}*20%+G{r}*30%,0))"),
                            C("Predikat", 9, "teks", "=IF(H{r}=\"\",\"\",IF(H{r}>=90,\"A\",IF(H{r}>=80,\"B\",IF(H{r}>=70,\"C\",\"D\"))))"),
                            C("Peringkat", 10, "angka", "=IF(H{r}=\"\",\"\",RANK(H{r},$H${first}:$H${last},0))"),
                            C("Keterangan", 14, "teks", "=IF(H{r}=\"\",\"\",IF(H{r}>=N($C$4),\"Tuntas\",\"Belum tuntas\"))") },
                Summary = { ("Rata-rata kelas", "=IFERROR(AVERAGE(H{first}:H{last}),0)", "desimal"), ("Nilai tertinggi", "=MAX(H{first}:H{last})", "desimal"), ("Nilai terendah", "=MIN(H{first}:H{last})", "desimal"), ("Jumlah tuntas", "=COUNTIF(K{first}:K{last},\"Tuntas\")", "angka") },
            },
            new SheetTemplate
            {
                Id = "penggajian", Name = "Daftar Gaji / Payroll", Category = "Keuangan", Title = "DAFTAR GAJI KARYAWAN",
                Description = "Gaji pokok, tunjangan, lembur, potongan BPJS (JHT 2%, JP 1%, Kesehatan 1%), PPh 21, gaji bersih.",
                InfoRows = { "Periode:", "Perusahaan:" },
                Columns = { C("No", 5), C("Nama", 26), C("Jabatan", 16), Sum("Gaji Pokok"), Sum("Tunjangan"), Sum("Lembur"), Sum("Bruto", 16, "rupiah", "=D{r}+E{r}+F{r}"),
                            Sum("BPJS JHT 2%", 14, "rupiah", "=ROUND(D{r}*2%,0)"), Sum("BPJS JP 1%", 14, "rupiah", "=ROUND(MIN(D{r},10547400)*1%,0)"), Sum("BPJS Kes 1%", 14, "rupiah", "=ROUND(MIN(D{r}+E{r},12000000)*1%,0)"),
                            Sum("PPh 21", 14), Sum("Potongan Lain", 14), Sum("Gaji Bersih", 18, "rupiah", "=G{r}-SUM(H{r}:L{r})") },
            },
            new SheetTemplate
            {
                Id = "inventaris", Name = "Inventaris Barang", Category = "Kantor", Title = "DAFTAR INVENTARIS BARANG",
                Description = "Kode, nama, merk, tahun, jumlah, kondisi (dropdown), lokasi, nilai aset.",
                InfoRows = { "Unit/Ruang:", "Tanggal pendataan:" },
                Columns = { C("No", 5), C("Kode Barang", 14), C("Nama Barang", 26), C("Merk/Tipe", 16), C("Tahun", 8, "angka"), C("Jumlah", 8, "angka", null, "=SUM({col}{first}:{col}{last})"),
                            new TemplateColumn { Header = "Kondisi", Width = 14, Choices = new[] { "Baik", "Rusak Ringan", "Rusak Berat" } }, C("Lokasi", 14), C("Harga Satuan", 16, "rupiah"), Sum("Nilai Total", 18, "rupiah", "=F{r}*I{r}") },
            },
            new SheetTemplate
            {
                Id = "stok_barang", Name = "Kartu Stok Barang", Category = "Usaha", Title = "LAPORAN STOK BARANG",
                Description = "Stok awal, masuk, keluar, stok akhir, dan status pesan ulang otomatis.",
                InfoRows = { "Gudang/Toko:", "Periode:" },
                Columns = { C("No", 5), C("Kode", 12), C("Nama Barang", 28), C("Satuan", 8), C("Stok Awal", 10, "angka"), C("Masuk", 9, "angka", null, "=SUM({col}{first}:{col}{last})"), C("Keluar", 9, "angka", null, "=SUM({col}{first}:{col}{last})"),
                            C("Stok Akhir", 10, "angka", "=E{r}+F{r}-G{r}"), C("Stok Minimum", 12, "angka"), C("Status", 14, "teks", "=IF(C{r}=\"\",\"\",IF(H{r}<=I{r},\"PESAN ULANG\",\"AMAN\"))"), C("Harga", 14, "rupiah"), Sum("Nilai Stok", 16, "rupiah", "=H{r}*K{r}") },
            },
            new SheetTemplate
            {
                Id = "penjualan", Name = "Data Penjualan", Category = "Usaha", Title = "DATA PENJUALAN", DefaultRows = 50,
                Description = "Struktur data penjualan siap pivot & dashboard.",
                Columns = { C("No", 5), C("Tanggal", 12, "tanggal"), C("No. Faktur", 12), C("Pelanggan", 22), C("Produk", 20), C("Wilayah", 14), C("Qty", 8, "angka", null, "=SUM({col}{first}:{col}{last})"), C("Harga", 14, "rupiah"), Sum("Total", 16, "rupiah", "=G{r}*H{r}") },
            },
            new SheetTemplate
            {
                Id = "piutang", Name = "Kartu Piutang", Category = "Keuangan", Title = "DAFTAR PIUTANG USAHA",
                Description = "Faktur, jatuh tempo, pembayaran, sisa, umur piutang, dan status.",
                Columns = { C("No", 5), C("Pelanggan", 24), C("No. Faktur", 12), C("Tanggal", 12, "tanggal"), C("Jatuh Tempo", 12, "tanggal", "=IF(D{r}=\"\",\"\",D{r}+30)"), Sum("Jumlah"), Sum("Dibayar"), Sum("Sisa", 16, "rupiah", "=F{r}-G{r}"),
                            C("Umur (hari)", 10, "angka", "=IF(OR(D{r}=\"\",H{r}<=0),\"\",TODAY()-E{r})"), C("Status", 14, "teks", "=IF(C{r}=\"\",\"\",IF(H{r}<=0,\"LUNAS\",IF(TODAY()>E{r},\"JATUH TEMPO\",\"BERJALAN\")))") },
            },
            new SheetTemplate
            {
                Id = "jadwal_kegiatan", Name = "Jadwal / Timeline Kegiatan", Category = "Kantor", Title = "JADWAL KEGIATAN",
                Description = "Rencana kegiatan dengan PIC, tanggal, durasi, status, dan progres.",
                Columns = { C("No", 5), C("Kegiatan", 34), C("Penanggung Jawab", 18), C("Mulai", 12, "tanggal"), C("Selesai", 12, "tanggal"), C("Durasi (hari)", 11, "angka", "=IF(OR(D{r}=\"\",E{r}=\"\"),\"\",E{r}-D{r}+1)"),
                            new TemplateColumn { Header = "Status", Width = 14, Choices = new[] { "Belum mulai", "Berjalan", "Selesai", "Tertunda" } }, C("Progres", 10, "persen") },
            },
            new SheetTemplate
            {
                Id = "kuesioner", Name = "Tabulasi Kuesioner (Likert)", Category = "Akademik", Title = "TABULASI DATA KUESIONER", DefaultRows = 30,
                Description = "Tabulasi skor Likert 1–5 per item, total skor, siap uji validitas & reliabilitas.",
                InfoRows = { "Variabel:", "Skala: 1 = Sangat Tidak Setuju ... 5 = Sangat Setuju" },
                Columns = BuildLikertColumns(10),
            },
            new SheetTemplate
            {
                Id = "laba_rugi", Name = "Laporan Laba Rugi", Category = "Keuangan", Title = "LAPORAN LABA RUGI",
                Description = "Laba rugi bentuk bertahap: pendapatan, HPP, laba kotor, beban, laba sebelum/sesudah pajak.",
                Lines =
                {
                    L("h1", "PENDAPATAN", "heading"), L("pend", "Penjualan bersih"), L("pendlain", "Pendapatan lain-lain"),
                    L("totpend", "Total Pendapatan", "total", "={pend}+{pendlain}"), L("b1", "", "blank"),
                    L("hpp", "Harga Pokok Penjualan"), L("lk", "LABA KOTOR", "total", "={totpend}-{hpp}"), L("b2", "", "blank"),
                    L("h2", "BEBAN USAHA", "heading"), L("gaji", "Beban gaji"), L("sewa", "Beban sewa"), L("listrik", "Beban listrik, air, internet"),
                    L("penyusutan", "Beban penyusutan"), L("pemasaran", "Beban pemasaran"), L("bebanlain", "Beban lain-lain"),
                    L("totbeban", "Total Beban Usaha", "total", "=SUM({gaji}:{bebanlain})"), L("b3", "", "blank"),
                    L("lsp", "LABA SEBELUM PAJAK", "total", "={lk}-{totbeban}"), L("pajak", "Pajak penghasilan (tarif badan 22%)", "formula", "=MAX(0,{lsp}*22%)"),
                    L("lb", "LABA BERSIH", "total", "={lsp}-{pajak}"), L("npm", "Margin laba bersih", "formula", "=IF({totpend}=0,0,{lb}/{totpend})", "persen"),
                },
            },
            new SheetTemplate
            {
                Id = "neraca", Name = "Neraca (Laporan Posisi Keuangan)", Category = "Keuangan", Title = "NERACA",
                Description = "Aset lancar & tetap, kewajiban, ekuitas, dengan cek keseimbangan.",
                Lines =
                {
                    L("h1", "ASET", "heading"), L("kas", "Kas dan setara kas"), L("piutang", "Piutang usaha"), L("persediaan", "Persediaan"), L("aslain", "Aset lancar lain"),
                    L("tal", "Total Aset Lancar", "total", "=SUM({kas}:{aslain})"), L("peralatan", "Peralatan dan kendaraan"), L("gedung", "Tanah dan bangunan"),
                    L("akum", "Akumulasi penyusutan (isi negatif)"), L("tat", "Total Aset Tetap", "total", "=SUM({peralatan}:{akum})"),
                    L("ta", "TOTAL ASET", "total", "={tal}+{tat}"), L("b1", "", "blank"),
                    L("h2", "KEWAJIBAN DAN EKUITAS", "heading"), L("utangusaha", "Utang usaha"), L("utangpajak", "Utang pajak"), L("utangbank", "Utang bank jangka panjang"),
                    L("tk", "Total Kewajiban", "total", "=SUM({utangusaha}:{utangbank})"), L("modal", "Modal disetor"), L("laba", "Laba ditahan"),
                    L("te", "Total Ekuitas", "total", "={modal}+{laba}"), L("tke", "TOTAL KEWAJIBAN DAN EKUITAS", "total", "={tk}+{te}"),
                    L("cek", "Cek keseimbangan", "formula", "=IF(ROUND({ta}-{tke},0)=0,\"SEIMBANG\",\"TIDAK SEIMBANG: \"&TEXT({ta}-{tke},\"#,##0\"))", "teks"),
                },
            },
            new SheetTemplate
            {
                Id = "kuitansi", Name = "Kuitansi", Category = "Kantor", Title = "KUITANSI",
                Description = "Kuitansi pembayaran dengan terbilang otomatis.",
                Lines =
                {
                    L("no", "No.", "input", null, "teks"), L("dari", "Sudah terima dari", "input", null, "teks"), L("jumlah", "Uang sejumlah"),
                    L("terbilang", "Terbilang", "formula", "=TERBILANG.RUPIAH({jumlah})", "teks"), L("untuk", "Untuk pembayaran", "input", null, "teks"),
                    L("b1", "", "blank"), L("tgl", "Tempat, tanggal", "formula", "=\"[Kota], \"&TANGGAL.INDO(TODAY())", "teks"), L("penerima", "Penerima", "input", null, "teks"),
                },
            },
        };

        private static List<TemplateColumn> BuildAttendanceColumns()
        {
            var cols = new List<TemplateColumn> { C("No", 5), C("Nama", 26) };
            for (int d = 1; d <= 31; d++) cols.Add(new TemplateColumn { Header = d.ToString(), Width = 3.5, Choices = new[] { "H", "S", "I", "A", "C" } });
            // kolom C..AG berisi hari 1..31
            foreach (var code in new[] { "H", "S", "I", "A", "C" })
                cols.Add(C(code, 5, "angka", "=COUNTIF(C{r}:AG{r},\"" + code + "\")", "=SUM({col}{first}:{col}{last})"));
            cols.Add(C("% Hadir", 9, "persen", "=IF(COUNTA(C{r}:AG{r})=0,\"\",AH{r}/COUNTA(C{r}:AG{r}))"));
            return cols;
        }

        private static List<TemplateColumn> BuildLikertColumns(int items)
        {
            var cols = new List<TemplateColumn> { C("Responden", 11) };
            for (int i = 1; i <= items; i++) cols.Add(new TemplateColumn { Header = "P" + i, Width = 5, Format = "angka", Number = true, Choices = new[] { "1", "2", "3", "4", "5" } });
            var last = FormulaTools.ColumnLetter(items + 1);
            cols.Add(C("Total", 8, "angka", "=IF(COUNT(B{r}:" + last + "{r})=0,\"\",SUM(B{r}:" + last + "{r}))"));
            return cols;
        }

        public static SheetTemplate Find(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return null;
            var key = id.Trim().ToLowerInvariant().Replace(" ", "_").Replace("-", "_");
            return All.FirstOrDefault(t => t.Id == key)
                ?? All.FirstOrDefault(t => t.Name.ToLowerInvariant().Contains(id.ToLowerInvariant()))
                ?? All.FirstOrDefault(t => key.Contains(t.Id) || t.Id.Contains(key));
        }

        /// <summary>Ganti placeholder rumus kolom: {r}, {p}, {first}, {last}, {last+N}, {col}.</summary>
        public static string Expand(string formula, int row, int first, int last, string colLetter)
        {
            if (formula == null) return null;
            var f = formula.Replace("{r}", row.ToString()).Replace("{p}", (row - 1).ToString())
                           .Replace("{first}", first.ToString()).Replace("{col}", colLetter ?? "");
            for (int k = 9; k >= 1; k--) f = f.Replace("{last+" + k + "}", (last + k).ToString());
            return f.Replace("{last}", last.ToString());
        }
    }
}
