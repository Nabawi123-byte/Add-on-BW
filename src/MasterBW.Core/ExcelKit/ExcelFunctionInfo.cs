using System;
using System.Collections.Generic;
using System.Linq;

namespace MasterBW.Core.ExcelKit
{
    /// <summary>Kamus fungsi Excel 2016 dengan penjelasan bahasa Indonesia (untuk penjelas rumus offline tanpa AI).</summary>
    public static class ExcelFunctionInfo
    {
        /// <summary>Fungsi yang TIDAK ada di Excel 2016 perpetual (baru di 2019/2021/365).</summary>
        public static readonly string[] NotIn2016 =
        {
            "XLOOKUP", "XMATCH", "FILTER", "UNIQUE", "SORT", "SORTBY", "SEQUENCE", "RANDARRAY", "LET", "LAMBDA",
            "TEXTJOIN", "CONCAT", "IFS", "SWITCH", "MAXIFS", "MINIFS", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER",
            "VSTACK", "HSTACK", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS",
            "EXPAND", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY", "ARRAYTOTEXT", "VALUETOTEXT", "IMAGE", "GROUPBY", "PIVOTBY"
        };

        /// <summary>Pengganti yang kompatibel Excel 2016.</summary>
        public static readonly Dictionary<string, string> Alternatives = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["XLOOKUP"] = "INDEX(kolom_hasil, MATCH(nilai, kolom_cari, 0))",
            ["XMATCH"] = "MATCH(nilai, range, 0)",
            ["FILTER"] = "AutoFilter, PivotTable, atau rumus INDEX/SMALL/IF (array)",
            ["UNIQUE"] = "Hapus Duplikat atau PivotTable",
            ["SORT"] = "fitur Urutkan (sort_range)",
            ["TEXTJOIN"] = "gabungan & atau CONCATENATE",
            ["CONCAT"] = "CONCATENATE atau operator &",
            ["IFS"] = "IF bertingkat: IF(k1,h1,IF(k2,h2,h3))",
            ["SWITCH"] = "IF bertingkat atau CHOOSE/MATCH",
            ["MAXIFS"] = "AGGREGATE(14,6,range/(kriteria),1) atau MAX(IF(...)) sebagai rumus array",
            ["MINIFS"] = "AGGREGATE(15,6,range/(kriteria),1) atau MIN(IF(...)) sebagai rumus array",
        };

        public sealed class Fn
        {
            public string Name;
            public string Syntax;
            public string Description;
            public string Category;
        }

        private static readonly List<Fn> Functions = new List<Fn>
        {
            F("SUM", "SUM(angka1, [angka2], ...)", "Menjumlahkan semua angka.", "Matematika"),
            F("SUMIF", "SUMIF(range_kriteria, kriteria, [range_jumlah])", "Menjumlahkan nilai yang memenuhi satu kriteria.", "Matematika"),
            F("SUMIFS", "SUMIFS(range_jumlah, range_kriteria1, kriteria1, ...)", "Menjumlahkan nilai yang memenuhi banyak kriteria sekaligus.", "Matematika"),
            F("SUMPRODUCT", "SUMPRODUCT(array1, [array2], ...)", "Mengalikan elemen array yang bersesuaian lalu menjumlahkannya; sering dipakai untuk penjumlahan bersyarat kompleks.", "Matematika"),
            F("PRODUCT", "PRODUCT(angka1, ...)", "Mengalikan semua angka.", "Matematika"),
            F("ROUND", "ROUND(angka, jumlah_digit)", "Membulatkan angka ke jumlah digit tertentu.", "Matematika"),
            F("ROUNDUP", "ROUNDUP(angka, digit)", "Membulatkan ke atas (menjauhi nol).", "Matematika"),
            F("ROUNDDOWN", "ROUNDDOWN(angka, digit)", "Membulatkan ke bawah (mendekati nol).", "Matematika"),
            F("MROUND", "MROUND(angka, kelipatan)", "Membulatkan ke kelipatan terdekat (mis. ke 500 rupiah).", "Matematika"),
            F("CEILING", "CEILING(angka, kelipatan)", "Membulatkan ke atas ke kelipatan tertentu.", "Matematika"),
            F("FLOOR", "FLOOR(angka, kelipatan)", "Membulatkan ke bawah ke kelipatan tertentu.", "Matematika"),
            F("INT", "INT(angka)", "Mengambil bilangan bulat (membulatkan ke bawah).", "Matematika"),
            F("MOD", "MOD(angka, pembagi)", "Sisa hasil bagi.", "Matematika"),
            F("ABS", "ABS(angka)", "Nilai mutlak.", "Matematika"),
            F("POWER", "POWER(angka, pangkat)", "Memangkatkan angka.", "Matematika"),
            F("SQRT", "SQRT(angka)", "Akar kuadrat.", "Matematika"),
            F("SUBTOTAL", "SUBTOTAL(kode_fungsi, ref1, ...)", "Agregasi yang mengabaikan baris tersembunyi/terfilter (9=SUM, 1=AVERAGE, 3=COUNTA).", "Matematika"),
            F("AGGREGATE", "AGGREGATE(fungsi, opsi, array, [k])", "Agregasi yang bisa mengabaikan error dan baris tersembunyi.", "Matematika"),
            F("RAND", "RAND()", "Angka acak 0–1.", "Matematika"),
            F("RANDBETWEEN", "RANDBETWEEN(bawah, atas)", "Bilangan bulat acak di antara dua nilai.", "Matematika"),

            F("AVERAGE", "AVERAGE(angka1, ...)", "Rata-rata aritmetika.", "Statistik"),
            F("AVERAGEIF", "AVERAGEIF(range, kriteria, [range_rata])", "Rata-rata nilai yang memenuhi satu kriteria.", "Statistik"),
            F("AVERAGEIFS", "AVERAGEIFS(range_rata, range_kriteria1, kriteria1, ...)", "Rata-rata dengan banyak kriteria.", "Statistik"),
            F("COUNT", "COUNT(nilai1, ...)", "Menghitung sel berisi angka.", "Statistik"),
            F("COUNTA", "COUNTA(nilai1, ...)", "Menghitung sel yang tidak kosong.", "Statistik"),
            F("COUNTBLANK", "COUNTBLANK(range)", "Menghitung sel kosong.", "Statistik"),
            F("COUNTIF", "COUNTIF(range, kriteria)", "Menghitung sel yang memenuhi satu kriteria.", "Statistik"),
            F("COUNTIFS", "COUNTIFS(range1, kriteria1, ...)", "Menghitung sel yang memenuhi banyak kriteria.", "Statistik"),
            F("MAX", "MAX(angka1, ...)", "Nilai terbesar.", "Statistik"),
            F("MIN", "MIN(angka1, ...)", "Nilai terkecil.", "Statistik"),
            F("LARGE", "LARGE(array, k)", "Nilai terbesar ke-k.", "Statistik"),
            F("SMALL", "SMALL(array, k)", "Nilai terkecil ke-k.", "Statistik"),
            F("MEDIAN", "MEDIAN(angka1, ...)", "Nilai tengah.", "Statistik"),
            F("MODE", "MODE(angka1, ...)", "Nilai yang paling sering muncul.", "Statistik"),
            F("MODE.SNGL", "MODE.SNGL(angka1, ...)", "Nilai yang paling sering muncul.", "Statistik"),
            F("STDEV", "STDEV(angka1, ...)", "Simpangan baku sampel.", "Statistik"),
            F("STDEV.S", "STDEV.S(angka1, ...)", "Simpangan baku sampel.", "Statistik"),
            F("STDEV.P", "STDEV.P(angka1, ...)", "Simpangan baku populasi.", "Statistik"),
            F("VAR.S", "VAR.S(angka1, ...)", "Varians sampel.", "Statistik"),
            F("RANK", "RANK(angka, ref, [urutan])", "Peringkat angka dalam daftar (0 = terbesar peringkat 1).", "Statistik"),
            F("RANK.EQ", "RANK.EQ(angka, ref, [urutan])", "Peringkat; nilai kembar mendapat peringkat sama.", "Statistik"),
            F("RANK.AVG", "RANK.AVG(angka, ref, [urutan])", "Peringkat; nilai kembar mendapat rata-rata peringkat.", "Statistik"),
            F("PERCENTILE.INC", "PERCENTILE.INC(array, k)", "Persentil ke-k (0–1).", "Statistik"),
            F("QUARTILE.INC", "QUARTILE.INC(array, kuartil)", "Kuartil (0–4).", "Statistik"),
            F("CORREL", "CORREL(array1, array2)", "Koefisien korelasi Pearson.", "Statistik"),
            F("FORECAST", "FORECAST(x, y_diketahui, x_diketahui)", "Meramal nilai dengan tren linear.", "Statistik"),
            F("FORECAST.LINEAR", "FORECAST.LINEAR(x, y_diketahui, x_diketahui)", "Meramal nilai dengan tren linear.", "Statistik"),
            F("TREND", "TREND(y_diketahui, [x_diketahui], [x_baru])", "Nilai tren linear.", "Statistik"),
            F("SLOPE", "SLOPE(y, x)", "Kemiringan garis regresi.", "Statistik"),
            F("INTERCEPT", "INTERCEPT(y, x)", "Titik potong garis regresi.", "Statistik"),
            F("FREQUENCY", "FREQUENCY(data, bin)", "Distribusi frekuensi (rumus array).", "Statistik"),
            F("T.TEST", "T.TEST(array1, array2, ekor, tipe)", "Nilai p uji t.", "Statistik"),

            F("IF", "IF(uji_logika, nilai_jika_benar, [nilai_jika_salah])", "Mengembalikan satu nilai jika kondisi benar dan nilai lain jika salah.", "Logika"),
            F("AND", "AND(logika1, ...)", "BENAR jika semua kondisi benar.", "Logika"),
            F("OR", "OR(logika1, ...)", "BENAR jika salah satu kondisi benar.", "Logika"),
            F("NOT", "NOT(logika)", "Membalik BENAR/SALAH.", "Logika"),
            F("IFERROR", "IFERROR(nilai, nilai_jika_error)", "Mengganti hasil error dengan nilai lain.", "Logika"),
            F("IFNA", "IFNA(nilai, nilai_jika_na)", "Mengganti error #N/A dengan nilai lain.", "Logika"),
            F("ISBLANK", "ISBLANK(nilai)", "BENAR jika sel kosong.", "Logika"),
            F("ISNUMBER", "ISNUMBER(nilai)", "BENAR jika nilai berupa angka.", "Logika"),
            F("ISERROR", "ISERROR(nilai)", "BENAR jika nilai error.", "Logika"),
            F("ISTEXT", "ISTEXT(nilai)", "BENAR jika nilai berupa teks.", "Logika"),

            F("VLOOKUP", "VLOOKUP(nilai_cari, tabel, no_kolom, [range_lookup])", "Mencari nilai di kolom pertama tabel lalu mengambil nilai dari kolom ke-n. Isi FALSE/0 untuk pencarian persis.", "Pencarian"),
            F("HLOOKUP", "HLOOKUP(nilai_cari, tabel, no_baris, [range_lookup])", "Seperti VLOOKUP tetapi mencari mendatar di baris pertama.", "Pencarian"),
            F("LOOKUP", "LOOKUP(nilai_cari, vektor_cari, [vektor_hasil])", "Pencarian sederhana pada data terurut.", "Pencarian"),
            F("INDEX", "INDEX(array, no_baris, [no_kolom])", "Mengambil nilai pada posisi baris/kolom tertentu.", "Pencarian"),
            F("MATCH", "MATCH(nilai_cari, array, [tipe])", "Mencari posisi nilai dalam daftar (0 = persis).", "Pencarian"),
            F("CHOOSE", "CHOOSE(indeks, nilai1, nilai2, ...)", "Memilih nilai berdasarkan nomor urut.", "Pencarian"),
            F("OFFSET", "OFFSET(ref, baris, kolom, [tinggi], [lebar])", "Referensi yang digeser sejumlah baris/kolom.", "Pencarian"),
            F("INDIRECT", "INDIRECT(teks_ref)", "Mengubah teks menjadi referensi sel.", "Pencarian"),
            F("ROW", "ROW([ref])", "Nomor baris.", "Pencarian"),
            F("COLUMN", "COLUMN([ref])", "Nomor kolom.", "Pencarian"),
            F("ROWS", "ROWS(array)", "Jumlah baris.", "Pencarian"),
            F("TRANSPOSE", "TRANSPOSE(array)", "Menukar baris menjadi kolom.", "Pencarian"),
            F("HYPERLINK", "HYPERLINK(alamat, [teks])", "Membuat tautan.", "Pencarian"),

            F("CONCATENATE", "CONCATENATE(teks1, teks2, ...)", "Menggabungkan teks.", "Teks"),
            F("LEFT", "LEFT(teks, [jumlah])", "Mengambil karakter dari kiri.", "Teks"),
            F("RIGHT", "RIGHT(teks, [jumlah])", "Mengambil karakter dari kanan.", "Teks"),
            F("MID", "MID(teks, mulai, jumlah)", "Mengambil karakter dari tengah.", "Teks"),
            F("LEN", "LEN(teks)", "Jumlah karakter.", "Teks"),
            F("FIND", "FIND(cari, dalam_teks, [mulai])", "Posisi teks (peka huruf besar/kecil).", "Teks"),
            F("SEARCH", "SEARCH(cari, dalam_teks, [mulai])", "Posisi teks (tidak peka huruf besar/kecil, mendukung wildcard).", "Teks"),
            F("SUBSTITUTE", "SUBSTITUTE(teks, lama, baru, [ke-])", "Mengganti teks tertentu.", "Teks"),
            F("REPLACE", "REPLACE(teks, mulai, jumlah, baru)", "Mengganti karakter pada posisi tertentu.", "Teks"),
            F("TRIM", "TRIM(teks)", "Menghapus spasi berlebih.", "Teks"),
            F("CLEAN", "CLEAN(teks)", "Menghapus karakter tak tercetak.", "Teks"),
            F("UPPER", "UPPER(teks)", "Huruf besar semua.", "Teks"),
            F("LOWER", "LOWER(teks)", "Huruf kecil semua.", "Teks"),
            F("PROPER", "PROPER(teks)", "Huruf besar di awal setiap kata.", "Teks"),
            F("TEXT", "TEXT(nilai, format)", "Mengubah angka/tanggal menjadi teks berformat, mis. TEXT(A1,\"dd mmmm yyyy\").", "Teks"),
            F("VALUE", "VALUE(teks)", "Mengubah teks angka menjadi angka.", "Teks"),
            F("REPT", "REPT(teks, jumlah)", "Mengulang teks.", "Teks"),
            F("EXACT", "EXACT(teks1, teks2)", "Membandingkan dua teks persis.", "Teks"),

            F("TODAY", "TODAY()", "Tanggal hari ini.", "Tanggal"),
            F("NOW", "NOW()", "Tanggal dan jam sekarang.", "Tanggal"),
            F("DATE", "DATE(tahun, bulan, hari)", "Membuat tanggal.", "Tanggal"),
            F("YEAR", "YEAR(tanggal)", "Tahun dari tanggal.", "Tanggal"),
            F("MONTH", "MONTH(tanggal)", "Bulan (1–12) dari tanggal.", "Tanggal"),
            F("DAY", "DAY(tanggal)", "Hari (1–31) dari tanggal.", "Tanggal"),
            F("WEEKDAY", "WEEKDAY(tanggal, [tipe])", "Hari dalam minggu (angka).", "Tanggal"),
            F("WEEKNUM", "WEEKNUM(tanggal, [tipe])", "Minggu ke- dalam tahun.", "Tanggal"),
            F("EOMONTH", "EOMONTH(tanggal, bulan)", "Tanggal akhir bulan.", "Tanggal"),
            F("EDATE", "EDATE(tanggal, bulan)", "Tanggal yang digeser sejumlah bulan (jatuh tempo).", "Tanggal"),
            F("DATEDIF", "DATEDIF(awal, akhir, satuan)", "Selisih tanggal dalam \"y\" tahun, \"m\" bulan, \"d\" hari (untuk usia/masa kerja).", "Tanggal"),
            F("NETWORKDAYS", "NETWORKDAYS(awal, akhir, [libur])", "Jumlah hari kerja.", "Tanggal"),
            F("WORKDAY", "WORKDAY(awal, hari, [libur])", "Tanggal setelah sejumlah hari kerja.", "Tanggal"),
            F("DATEVALUE", "DATEVALUE(teks)", "Mengubah teks tanggal menjadi tanggal.", "Tanggal"),
            F("HOUR", "HOUR(waktu)", "Jam.", "Tanggal"),
            F("MINUTE", "MINUTE(waktu)", "Menit.", "Tanggal"),

            F("PMT", "PMT(bunga, nper, pv, [fv], [tipe])", "Angsuran per periode pinjaman anuitas (hasil negatif = uang keluar).", "Keuangan"),
            F("IPMT", "IPMT(bunga, per, nper, pv)", "Porsi bunga angsuran periode tertentu.", "Keuangan"),
            F("PPMT", "PPMT(bunga, per, nper, pv)", "Porsi pokok angsuran periode tertentu.", "Keuangan"),
            F("PV", "PV(bunga, nper, pmt, [fv])", "Nilai sekarang.", "Keuangan"),
            F("FV", "FV(bunga, nper, pmt, [pv])", "Nilai masa depan (tabungan/investasi).", "Keuangan"),
            F("NPV", "NPV(bunga, nilai1, ...)", "Net present value; arus kas dimulai periode 1 (investasi awal ditambahkan terpisah).", "Keuangan"),
            F("IRR", "IRR(nilai, [tebakan])", "Tingkat pengembalian internal.", "Keuangan"),
            F("XNPV", "XNPV(bunga, nilai, tanggal)", "NPV dengan tanggal tidak beraturan.", "Keuangan"),
            F("XIRR", "XIRR(nilai, tanggal)", "IRR dengan tanggal tidak beraturan.", "Keuangan"),
            F("RATE", "RATE(nper, pmt, pv)", "Suku bunga per periode.", "Keuangan"),
            F("NPER", "NPER(bunga, pmt, pv)", "Jumlah periode.", "Keuangan"),
            F("SLN", "SLN(harga, sisa, umur)", "Penyusutan garis lurus.", "Keuangan"),
            F("DB", "DB(harga, sisa, umur, periode)", "Penyusutan saldo menurun.", "Keuangan"),
            F("DDB", "DDB(harga, sisa, umur, periode)", "Penyusutan saldo menurun ganda.", "Keuangan"),
            F("SYD", "SYD(harga, sisa, umur, periode)", "Penyusutan jumlah angka tahun.", "Keuangan"),
        };

        private static Fn F(string n, string s, string d, string c) => new Fn { Name = n, Syntax = s, Description = d, Category = c };

        private static readonly Dictionary<string, Fn> Index = Functions.ToDictionary(f => f.Name, StringComparer.OrdinalIgnoreCase);

        public static Fn Get(string name) => name != null && Index.TryGetValue(name, out var f) ? f : null;
        public static IReadOnlyList<Fn> All => Functions;

        public static bool IsUnavailableIn2016(string name) => NotIn2016.Contains((name ?? "").ToUpperInvariant().Replace("_XLFN.", ""));
    }
}
