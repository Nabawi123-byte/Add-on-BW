using System.Collections.Generic;

namespace MasterBW.Core.Tools
{
    /// <summary>Katalog aksi Microsoft Excel 2016.</summary>
    internal static class ExcelTools
    {
        private static ToolDef T(string name, string label, string category, string desc)
            => new ToolDef { Name = name, App = OfficeApp.Excel, Label = label, Category = category, Description = desc };

        private const string Out = "lokasi hasil: 'new_sheet' atau sel kiri-atas mis. 'H2'";

        public static List<ToolDef> Build()
        {
            return new List<ToolDef>
            {
                // ===================== Rumus & data =====================
                T("write_formula", "Tulis rumus", "Rumus", "Tulis rumus Excel 2016 (nama fungsi Inggris, pemisah koma) di sel pertama range lalu salin ke seluruh range.")
                    .P("range", "string", "sel/range tujuan, mis. G2:G120", true)
                    .P("formula", "string", "rumus untuk sel pertama, diawali =", true)
                    .P("header", "string", "judul kolom ditulis di atas range (opsional)")
                    .K("rumus", "formula", "hitung", "jumlahkan", "total", "rata-rata", "sumif", "vlookup", "if", "persentase", "selisih", "ranking", "kolom baru"),

                T("write_values", "Isi nilai", "Data", "Tulis nilai/teks ke sel mulai dari sel kiri-atas. values = array baris.")
                    .P("start_cell", "string", "sel kiri-atas", true)
                    .P("values", "array", "[[..],[..]]", true)
                    .P("bold_header", "boolean", "baris pertama tebal")
                    .K("isi", "ketik", "tulis data", "masukkan"),

                T("explain_formula", "Jelaskan rumus", "Rumus", "Jelaskan rumus di sel terpilih langkah demi langkah.")
                    .P("cell", "string", "sel (kosong = sel aktif)")
                    .K("jelaskan", "arti rumus", "maksud rumus", "explain").Ro(),

                T("fix_formula", "Perbaiki rumus", "Rumus", "Diagnosis error (#N/A, #REF!, #VALUE!, #DIV/0!, #NAME?) dan usulkan rumus perbaikan.")
                    .P("cell", "string", "sel (kosong = sel aktif)")
                    .K("error", "perbaiki rumus", "#n/a", "#ref", "#value", "#div/0", "#name", "salah").Ai(),

                T("terbilang", "Terbilang", "Rumus", "Tulis angka sebagai kata (terbilang) bahasa Indonesia.")
                    .P("source", "string", "range angka", true).P("destination", "string", "range/sel hasil (kiri-atas)", true)
                    .P("rupiah", "boolean", "tambahkan 'rupiah'", false, null, "true")
                    .K("terbilang", "angka ke huruf", "huruf", "kuitansi"),

                T("ai_fill", "Isi kolom dengan AI", "AI", "AI memproses tiap baris (kategorisasi, ekstraksi, sentimen, ringkas, terjemah) dan menulis hasil ke kolom tujuan.")
                    .P("source", "string", "range sumber (satu kolom)", true)
                    .P("destination", "string", "sel kiri-atas hasil", true)
                    .P("instruction", "string", "instruksi per baris", true)
                    .P("categories", "array", "pilihan kategori (opsional)")
                    .K("kategorikan", "klasifikasi", "sentimen", "ekstrak", "ai per baris", "label otomatis").Ai(),

                // ===================== Format =====================
                T("format_number", "Format angka", "Format", "Format angka: rupiah, ribuan, persen, tanggal Indonesia, akuntansi.")
                    .P("range", "string", "range", true)
                    .P("format", "string", "jenis", true, new[] { "rupiah", "rupiah_desimal", "ribuan", "desimal2", "persen", "tanggal_panjang", "tanggal_pendek", "akuntansi", "teks", "custom" })
                    .P("custom", "string", "kode format jika custom")
                    .K("format rupiah", "rp", "mata uang", "persen", "ribuan", "format tanggal", "akuntansi"),

                T("format_cells", "Format sel", "Format", "Format tampilan sel.")
                    .P("range", "string", "range", true)
                    .P("bold", "boolean", "").P("italic", "boolean", "").P("font_size", "number", "")
                    .P("font_color", "string", "warna huruf").P("fill_color", "string", "warna isi")
                    .P("h_align", "string", "", false, new[] { "left", "center", "right" })
                    .P("v_align", "string", "", false, new[] { "top", "middle", "bottom" })
                    .P("wrap", "boolean", "bungkus teks").P("borders", "string", "", false, new[] { "all", "outline", "none" })
                    .P("merge", "boolean", "gabung sel")
                    .K("warna", "tebal", "border", "garis", "rata tengah", "merge", "gabung sel", "wrap"),

                T("format_as_table", "Jadikan tabel", "Format", "Format range sebagai Tabel Excel bergaya (filter, baris berselang).")
                    .P("range", "string", "range (kosong = area data aktif)")
                    .P("style", "string", "gaya, mis. TableStyleMedium2", false, null, "TableStyleMedium2")
                    .P("name", "string", "nama tabel")
                    .K("format as table", "jadikan tabel", "tabel excel", "rapikan tabel"),

                T("autofit", "Lebar otomatis", "Format", "Sesuaikan lebar kolom/tinggi baris otomatis.")
                    .P("range", "string", "range (kosong = sheet aktif)")
                    .K("autofit", "lebar kolom", "rapikan kolom"),

                T("conditional_format", "Format bersyarat", "Format", "Tambahkan format bersyarat.")
                    .P("range", "string", "range", true)
                    .P("rule", "string", "aturan", true, new[] { "greater_than", "less_than", "between", "equal", "text_contains", "duplicates", "unique", "top_n", "bottom_n", "above_average", "below_average", "color_scale", "data_bar", "icon_set", "formula" })
                    .P("value1", "string", "nilai/rumus 1").P("value2", "string", "nilai 2")
                    .P("color", "string", "warna sorot", false, null, "merah muda")
                    .K("format bersyarat", "conditional", "warnai jika", "sorot jika", "tandai", "data bar", "color scale"),

                T("freeze_panes", "Bekukan panel", "Format", "Bekukan baris/kolom di atas-kiri sel.")
                    .P("cell", "string", "sel, mis. B2", true)
                    .K("freeze", "bekukan", "kunci header"),

                T("print_setup", "Pengaturan cetak", "Format", "Atur cetak: area, orientasi, muat satu halaman, ulangi judul.")
                    .P("range", "string", "area cetak")
                    .P("orientation", "string", "", false, new[] { "portrait", "landscape" })
                    .P("paper", "string", "", false, new[] { "A4", "F4", "Letter", "Legal" })
                    .P("fit_width", "boolean", "muat lebar 1 halaman")
                    .P("title_rows", "string", "baris judul diulang, mis. $1:$1")
                    .K("cetak", "print", "area cetak", "satu halaman"),

                // ===================== Olah data =====================
                T("sort_range", "Urutkan", "Data", "Urutkan data berdasarkan kolom.")
                    .P("range", "string", "range (kosong = area data aktif)")
                    .P("key", "string", "kolom kunci (huruf atau nama header)", true)
                    .P("order", "string", "", false, new[] { "asc", "desc" }, "asc")
                    .P("key2", "string", "kolom kunci kedua").P("order2", "string", "", false, new[] { "asc", "desc" })
                    .K("urutkan", "sort", "terbesar", "terkecil", "abjad"),

                T("filter_range", "Filter", "Data", "Pasang AutoFilter dengan kriteria.")
                    .P("range", "string", "range (kosong = area data aktif)")
                    .P("column", "string", "kolom (huruf/nama header)", true)
                    .P("criteria", "string", "kriteria, mis. '>1000' atau 'Jawa'", true)
                    .K("filter", "saring", "tampilkan hanya"),

                T("remove_duplicates", "Hapus duplikat", "Data", "Hapus baris duplikat.")
                    .P("range", "string", "range (kosong = area data aktif)")
                    .P("columns", "array", "kolom pembanding (kosong = semua)")
                    .K("duplikat", "data ganda", "dobel", "duplicate").Danger(),

                T("clean_data", "Bersihkan data", "Data", "Bersihkan data teks/angka/tanggal secara massal.")
                    .P("range", "string", "range", true)
                    .P("operations", "array", "trim, proper, upper, lower, text_to_number, normalize_date, normalize_phone, remove_nonprintable, remove_extra_spaces, fix_rupiah", true)
                    .K("bersihkan", "rapikan data", "spasi", "huruf kapital", "angka teks", "tanggal berantakan", "nomor hp"),

                T("split_column", "Pisah kolom", "Data", "Pisahkan satu kolom menjadi beberapa kolom.")
                    .P("range", "string", "range kolom sumber", true)
                    .P("delimiter", "string", "pemisah: koma, spasi, titik koma, '-', '/'", true)
                    .P("destination", "string", "sel kiri-atas hasil")
                    .K("pisah", "split", "pisahkan nama", "text to columns"),

                T("combine_columns", "Gabung kolom", "Data", "Gabungkan beberapa kolom menjadi satu.")
                    .P("columns", "array", "kolom yang digabung, mis. [\"A\",\"B\"]", true)
                    .P("separator", "string", "pemisah", false, null, " ")
                    .P("destination", "string", "kolom/sel hasil", true)
                    .K("gabung", "combine", "concatenate", "satukan kolom"),

                T("data_validation", "Validasi data", "Data", "Batasi isian sel (dropdown, angka, tanggal).")
                    .P("range", "string", "range", true)
                    .P("type", "string", "", true, new[] { "list", "whole", "decimal", "date", "text_length" })
                    .P("values", "array", "pilihan dropdown")
                    .P("min", "string", "").P("max", "string", "")
                    .P("message", "string", "pesan error")
                    .K("dropdown", "validasi", "pilihan", "daftar pilihan", "batasi"),

                T("find_replace", "Cari & ganti", "Data", "Cari dan ganti nilai.")
                    .P("find", "string", "", true).P("replace", "string", "", true)
                    .P("range", "string", "range (kosong = sheet aktif)")
                    .P("whole_cell", "boolean", "").P("match_case", "boolean", "")
                    .K("ganti", "replace", "cari"),

                T("subtotal_group", "Subtotal per grup", "Data", "Buat ringkasan total per kelompok (mis. total penjualan per wilayah) ke lokasi baru.")
                    .P("range", "string", "range data (kosong = area aktif)")
                    .P("group_by", "string", "kolom kelompok", true)
                    .P("value_columns", "array", "kolom yang dijumlah", true)
                    .P("function", "string", "", false, new[] { "sum", "average", "count", "max", "min" }, "sum")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("per wilayah", "per bulan", "per kategori", "rekap", "kelompokkan", "group by", "subtotal"),

                // ===================== Pivot, grafik, dashboard =====================
                T("create_pivot", "PivotTable", "Analisis", "Buat PivotTable.")
                    .P("source", "string", "range sumber (kosong = area data aktif)")
                    .P("rows", "array", "field baris", true)
                    .P("columns", "array", "field kolom")
                    .P("values", "array", "[{field, function: sum|count|average|max|min}]", true)
                    .P("filters", "array", "field filter")
                    .P("destination", "string", Out, false, null, "new_sheet")
                    .K("pivot", "pivottable", "tabel pivot", "rekapitulasi", "ringkasan"),

                T("create_chart", "Grafik", "Analisis", "Buat grafik dari data.")
                    .P("source", "string", "range data termasuk header", true)
                    .P("type", "string", "jenis", true, new[] { "column", "bar", "line", "pie", "doughnut", "area", "scatter", "column_stacked", "combo" })
                    .P("title", "string", "judul")
                    .P("x_title", "string", "judul sumbu X").P("y_title", "string", "judul sumbu Y")
                    .P("position", "string", "sel letak grafik")
                    .K("grafik", "chart", "diagram", "kurva", "diagram batang", "pie", "lingkaran", "garis"),

                T("create_dashboard", "Dashboard otomatis", "Analisis", "Buat sheet dashboard (kartu KPI, pivot, dan grafik) dari tabel data.")
                    .P("source", "string", "range data (kosong = area aktif)")
                    .P("title", "string", "judul dashboard")
                    .P("measure", "string", "kolom nilai utama (mis. Total)")
                    .P("dimensions", "array", "kolom kategori untuk dirinci (mis. Wilayah, Produk)")
                    .P("date_column", "string", "kolom tanggal untuk tren")
                    .K("dashboard", "laporan visual", "kpi", "ringkasan visual"),

                T("analyze_data", "Analisis & insight", "Analisis", "Hitung statistik data lalu AI menulis temuan/insight dalam bahasa Indonesia.")
                    .P("range", "string", "range (kosong = area aktif)")
                    .P("question", "string", "pertanyaan fokus (opsional)")
                    .P("output", "string", "", false, new[] { "chat", "new_sheet" }, "chat")
                    .K("analisis", "insight", "temuan", "kesimpulan data", "tren", "analisa").Ai(),

                // ===================== Statistik (akademik/riset) =====================
                T("descriptive_stats", "Statistik deskriptif", "Statistik", "N, mean, median, modus, SD, varians, min, max, kuartil, skewness, kurtosis per kolom.")
                    .P("range", "string", "range data", true)
                    .P("has_headers", "boolean", "", false, null, "true")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("deskriptif", "rata-rata", "standar deviasi", "simpangan baku", "mean", "median", "modus"),

                T("correlation", "Korelasi", "Statistik", "Matriks korelasi Pearson/Spearman + nilai p.")
                    .P("range", "string", "range (kolom = variabel)", true)
                    .P("method", "string", "", false, new[] { "pearson", "spearman" }, "pearson")
                    .P("has_headers", "boolean", "", false, null, "true")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("korelasi", "hubungan", "correlation", "pearson", "spearman"),

                T("regression", "Regresi linear", "Statistik", "Regresi linear sederhana/berganda: koefisien, t, p, R², F, uji asumsi ringkas.")
                    .P("y_range", "string", "range Y (dependen)", true)
                    .P("x_range", "string", "range X (independen, boleh beberapa kolom)", true)
                    .P("has_headers", "boolean", "", false, null, "true")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("regresi", "pengaruh", "regression", "berganda"),

                T("t_test", "Uji t", "Statistik", "Uji t satu sampel, independen, atau berpasangan.")
                    .P("range1", "string", "range kelompok 1", true).P("range2", "string", "range kelompok 2")
                    .P("type", "string", "", true, new[] { "one_sample", "independent", "paired" })
                    .P("mu", "number", "nilai uji (one_sample)")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("uji t", "t-test", "beda rata-rata", "pretest posttest", "berpasangan"),

                T("anova", "ANOVA satu jalur", "Statistik", "ANOVA satu jalur antar kelompok (tiap kolom satu kelompok).")
                    .P("range", "string", "range (kolom = kelompok)", true)
                    .P("has_headers", "boolean", "", false, null, "true")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("anova", "anava", "beda lebih dari dua"),

                T("chi_square", "Uji chi-square", "Statistik", "Uji independensi chi-square dari tabel kontingensi.")
                    .P("range", "string", "tabel frekuensi (tanpa total)", true)
                    .P("has_headers", "boolean", "", false, null, "true")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("chi square", "chi kuadrat", "kai kuadrat", "kontingensi"),

                T("normality_test", "Uji normalitas", "Statistik", "Uji normalitas (Kolmogorov-Smirnov/Lilliefors, Shapiro-Wilk, Jarque-Bera) per kolom.")
                    .P("range", "string", "range", true)
                    .P("has_headers", "boolean", "", false, null, "true")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("normalitas", "normal", "kolmogorov", "shapiro"),

                T("validity_reliability", "Uji validitas & reliabilitas", "Statistik", "Validitas item (r hitung vs r tabel) dan reliabilitas Cronbach's Alpha untuk data kuesioner (baris = responden, kolom = item).")
                    .P("range", "string", "range skor item", true)
                    .P("has_headers", "boolean", "", false, null, "true")
                    .P("alpha", "number", "taraf signifikansi", false, null, "0.05")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("validitas", "reliabilitas", "cronbach", "kuesioner", "angket", "r tabel"),

                T("detect_outliers", "Deteksi outlier", "Statistik", "Tandai pencilan dengan metode IQR atau z-score.")
                    .P("range", "string", "range satu kolom", true)
                    .P("method", "string", "", false, new[] { "iqr", "zscore" }, "iqr")
                    .P("highlight", "boolean", "warnai sel", false, null, "true")
                    .K("outlier", "pencilan", "data aneh", "anomali"),

                T("forecast", "Peramalan tren", "Statistik", "Ramalkan periode berikutnya (tren linear, rata-rata bergerak, atau eksponensial).")
                    .P("y_range", "string", "range nilai historis", true)
                    .P("periods", "integer", "jumlah periode ke depan", false, null, "3")
                    .P("method", "string", "", false, new[] { "linear", "moving_average", "exponential" }, "linear")
                    .P("window", "integer", "jendela rata-rata bergerak", false, null, "3")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("ramal", "prediksi", "forecast", "proyeksi", "tren"),

                // ===================== Keuangan =====================
                T("loan_schedule", "Tabel angsuran", "Keuangan", "Jadwal angsuran pinjaman (anuitas, flat, efektif/menurun).")
                    .P("principal", "number", "pokok pinjaman", true)
                    .P("annual_rate", "number", "bunga per tahun (%)", true)
                    .P("months", "integer", "tenor bulan", true)
                    .P("method", "string", "", false, new[] { "annuity", "flat", "effective" }, "annuity")
                    .P("start_date", "string", "tanggal mulai")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("angsuran", "cicilan", "kredit", "pinjaman", "kpr", "amortisasi", "bunga"),

                T("depreciation_schedule", "Penyusutan aset", "Keuangan", "Jadwal penyusutan (garis lurus, saldo menurun, jumlah angka tahun).")
                    .P("cost", "number", "harga perolehan", true).P("salvage", "number", "nilai sisa", true)
                    .P("life_years", "integer", "umur ekonomis", true)
                    .P("method", "string", "", false, new[] { "straight_line", "declining_balance", "double_declining", "sum_of_years" }, "straight_line")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("penyusutan", "depresiasi", "aset tetap"),

                T("investment_analysis", "Kelayakan investasi", "Keuangan", "NPV, IRR, Profitability Index, Payback Period dari arus kas.")
                    .P("cashflows", "string", "range arus kas (periode 0 = investasi awal negatif) atau daftar angka", true)
                    .P("rate", "number", "tingkat diskonto (%)", true)
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("npv", "irr", "payback", "kelayakan", "investasi", "nilai sekarang"),

                T("tax_calculation", "Hitung pajak", "Keuangan", "Hitung PPh 21 (tahunan, tarif Pasal 17 + PTKP), PPN, PPh 23, PPh Final UMKM.")
                    .P("tax", "string", "", true, new[] { "pph21", "ppn", "pph23", "pph_final_umkm" })
                    .P("amount", "number", "penghasilan bruto per bulan / DPP / omzet", true)
                    .P("ptkp", "string", "status PTKP", false, new[] { "TK/0", "TK/1", "TK/2", "TK/3", "K/0", "K/1", "K/2", "K/3" }, "TK/0")
                    .P("has_npwp", "boolean", "", false, null, "true")
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("pajak", "pph", "pph 21", "ppn", "ptkp", "gaji bersih"),

                T("break_even", "Titik impas (BEP)", "Keuangan", "Hitung BEP unit & rupiah.")
                    .P("fixed_cost", "number", "biaya tetap", true).P("price", "number", "harga jual/unit", true)
                    .P("variable_cost", "number", "biaya variabel/unit", true)
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("bep", "break even", "titik impas", "balik modal"),

                T("financial_ratios", "Rasio keuangan", "Keuangan", "Hitung rasio likuiditas, solvabilitas, profitabilitas dari angka laporan keuangan.")
                    .P("values", "object", "{aset_lancar, utang_lancar, persediaan, kas, total_aset, total_utang, ekuitas, penjualan, laba_kotor, laba_bersih, laba_operasi}", true)
                    .P("output", "string", Out, false, null, "new_sheet")
                    .K("rasio", "current ratio", "roa", "roe", "der", "likuiditas", "profitabilitas"),

                // ===================== Templat =====================
                T("create_template", "Templat lembar kerja", "Templat", "Buat lembar kerja siap pakai lengkap dengan rumus & format.")
                    .P("template", "string", "id templat, lihat daftar templat", true)
                    .P("title", "string", "judul")
                    .P("rows", "integer", "jumlah baris isian", false, null, "20")
                    .K("templat", "template", "anggaran", "rab", "arus kas", "laba rugi", "neraca", "buku kas", "invoice", "faktur", "kuitansi", "daftar hadir", "absensi", "nilai siswa", "rapor", "inventaris", "gaji", "payroll", "jadwal", "stok"),

                // ===================== Sheet & struktur =====================
                T("add_sheet", "Tambah sheet", "Sheet", "Tambah worksheet baru.")
                    .P("name", "string", "nama", true)
                    .K("sheet baru", "tambah sheet", "lembar baru"),

                T("rename_sheet", "Ganti nama sheet", "Sheet", "Ganti nama worksheet.")
                    .P("old_name", "string", "nama lama (kosong = aktif)").P("new_name", "string", "nama baru", true)
                    .K("ganti nama sheet", "rename"),

                T("delete_sheet", "Hapus sheet", "Sheet", "Hapus worksheet.")
                    .P("name", "string", "nama", true)
                    .K("hapus sheet").Danger(),

                T("insert_cells", "Sisip baris/kolom", "Sheet", "Sisipkan baris atau kolom.")
                    .P("type", "string", "", true, new[] { "rows", "columns" })
                    .P("at", "string", "posisi: nomor baris (5) atau huruf kolom (C)", true)
                    .P("count", "integer", "jumlah", false, null, "1")
                    .K("sisip baris", "tambah kolom", "insert row"),

                T("delete_cells", "Hapus baris/kolom", "Sheet", "Hapus baris atau kolom.")
                    .P("type", "string", "", true, new[] { "rows", "columns", "blank_rows" })
                    .P("at", "string", "posisi/rentang: 5:7 atau C:D (blank_rows = range data)")
                    .K("hapus baris", "hapus kolom", "baris kosong").Danger(),

                T("define_name", "Nama range", "Sheet", "Beri nama pada range.")
                    .P("name", "string", "", true).P("range", "string", "", true)
                    .K("named range", "nama range"),

                T("protect_sheet", "Proteksi sheet", "Sheet", "Kunci/buka kunci sheet (sel yang di-unlock tetap bisa diisi).")
                    .P("protect", "boolean", "", true).P("password", "string", "")
                    .P("unlocked_range", "string", "range yang tetap bisa diedit")
                    .K("proteksi", "kunci sheet", "password", "protect"),

                T("hide_show", "Sembunyikan/tampilkan", "Sheet", "Sembunyikan atau tampilkan baris/kolom/sheet.")
                    .P("target", "string", "", true, new[] { "rows", "columns", "sheet" })
                    .P("ref", "string", "5:7, C:D, atau nama sheet", true)
                    .P("hide", "boolean", "", true)
                    .K("sembunyikan", "hide", "tampilkan", "unhide"),

                T("vba_macro", "Buat makro VBA", "Otomasi", "Buat makro VBA untuk tugas yang tidak ada di daftar aksi. Kode diperiksa keamanannya dan ditampilkan sebelum dipasang.")
                    .P("description", "string", "apa yang dilakukan makro", true)
                    .P("code", "string", "kode VBA lengkap (Sub ... End Sub)", true)
                    .P("run_now", "boolean", "jalankan setelah dipasang", false, null, "false")
                    .K("makro", "macro", "vba", "otomatis", "tombol")
                    .Ai(),
            };
        }
    }
}
