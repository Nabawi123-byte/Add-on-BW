using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using ExcelDna.Integration;
using MasterBW.Core.Ai;
using MasterBW.Core.Finance;
using MasterBW.Core.Indonesia;
using MasterBW.Core.Util;

namespace MasterBW.Functions
{
    /// <summary>
    /// Fungsi lembar kerja MASTER BW. Fungsi AI berjalan asinkron (sel menampilkan #N/A sesaat),
    /// hasilnya disimpan di cache disk agar tidak dihitung ulang setiap Excel kalkulasi.
    /// </summary>
    public static class CellFunctions
    {
        private static readonly EngineClient Client = new EngineClient();
        private static readonly ConcurrentDictionary<string, string> Memory = new ConcurrentDictionary<string, string>();

        // ================================================================== AI
        [ExcelFunction(Name = "AI", Category = "MASTER BW AI", Description = "Tanya AI offline. Contoh: =AI(\"Buat slogan untuk\"; A2)")]
        public static object Ai(
            [ExcelArgument(Name = "perintah", Description = "Instruksi untuk AI")] string prompt,
            [ExcelArgument(Name = "data1", Description = "Sel/teks tambahan (opsional)")] object data1,
            [ExcelArgument(Name = "data2", Description = "Sel/teks tambahan (opsional)")] object data2)
        {
            var extra = string.Join("\n", new[] { data1, data2 }.Select(Text).Where(s => s.Length > 0));
            var user = prompt + (extra.Length > 0 ? "\nData:\n" + extra : "");
            return RunAi("AI", user, "Jawab sangat singkat (maksimal 2 kalimat) dalam bahasa Indonesia, tanpa pembuka.", 160);
        }

        [ExcelFunction(Name = "AI.RINGKAS", Category = "MASTER BW AI", Description = "Ringkas teks menjadi satu kalimat.")]
        public static object Ringkas(string teks, object jumlahKata)
        {
            int n = jumlahKata is double d ? (int)d : 20;
            return RunAi("AI.RINGKAS", "Ringkas menjadi maksimal " + n + " kata:\n" + teks, "Tulis ringkasannya saja.", 120);
        }

        [ExcelFunction(Name = "AI.KATEGORI", Category = "MASTER BW AI", Description = "Klasifikasikan teks ke salah satu kategori. Contoh: =AI.KATEGORI(B2; \"Makanan,Minuman,Lainnya\")")]
        public static object Kategori(string teks, object kategori)
        {
            var cats = Text(kategori);
            return RunAi("AI.KATEGORI", "Kategorikan teks berikut" + (cats.Length > 0 ? " ke tepat satu dari: " + cats : "") + ".\nTeks: " + teks, "Jawab hanya dengan nama kategorinya.", 20);
        }

        [ExcelFunction(Name = "AI.SENTIMEN", Category = "MASTER BW AI", Description = "Sentimen teks: Positif/Negatif/Netral.")]
        public static object Sentimen(string teks)
            => RunAi("AI.SENTIMEN", "Tentukan sentimen (Positif, Negatif, atau Netral) dari: " + teks, "Jawab satu kata saja.", 8);

        [ExcelFunction(Name = "AI.TERJEMAH", Category = "MASTER BW AI", Description = "Terjemahkan teks. bahasa: \"Inggris\" (default) atau bahasa lain.")]
        public static object Terjemah(string teks, object bahasa)
        {
            var lang = Text(bahasa);
            return RunAi("AI.TERJEMAH", "Terjemahkan ke bahasa " + (lang.Length > 0 ? lang : "Inggris") + ":\n" + teks, "Tulis terjemahannya saja.", 200);
        }

        [ExcelFunction(Name = "AI.EKSTRAK", Category = "MASTER BW AI", Description = "Ambil informasi tertentu dari teks. Contoh: =AI.EKSTRAK(A2; \"nomor HP\")")]
        public static object Ekstrak(string teks, string apa)
            => RunAi("AI.EKSTRAK", "Ambil " + apa + " dari teks berikut. Jika tidak ada, jawab '-'.\nTeks: " + teks, "Jawab hanya nilainya.", 40);

        [ExcelFunction(Name = "AI.PERBAIKI", Category = "MASTER BW AI", Description = "Perbaiki ejaan dan tata bahasa (EYD).")]
        public static object Perbaiki(string teks)
            => RunAi("AI.PERBAIKI", "Perbaiki ejaan dan tata bahasa sesuai EYD tanpa mengubah makna:\n" + teks, "Tulis hasil perbaikannya saja.", 200);

        private static object RunAi(string name, string user, string system, int maxTokens)
        {
            if (string.IsNullOrWhiteSpace(user)) return "";
            var key = Hash(name + "\u0001" + system + "\u0001" + user);
            if (Memory.TryGetValue(key, out var cached)) return cached;
            var disk = Path.Combine(AppPaths.CacheDir, "cell", key + ".txt");
            if (File.Exists(disk)) { var t = File.ReadAllText(disk, Encoding.UTF8); Memory[key] = t; return t; }

            return ExcelAsyncUtil.Run(name, new object[] { key }, () =>
            {
                try
                {
                    var st = EngineLauncher.EnsureRunningAsync(Client, null, CancellationToken.None).GetAwaiter().GetResult();
                    if (!st.Ok) return "#AI: " + st.Error;
                    var req = new GenerateRequest
                    {
                        Messages = { ChatMessage.System("Kamu MASTER BW. " + system), ChatMessage.User(user) },
                        MaxTokens = maxTokens,
                        Temperature = 0.2f,
                        Reasoning = ReasoningMode.Cepat,
                    };
                    var r = Client.GenerateAsync(req, CancellationToken.None).GetAwaiter().GetResult();
                    var text = ThinkFilter.StripThinking(r.Text).Trim().Trim('"').Split('\n')[0].Trim();
                    Memory[key] = text;
                    try { Directory.CreateDirectory(Path.GetDirectoryName(disk)); File.WriteAllText(disk, text, Encoding.UTF8); } catch { }
                    return text;
                }
                catch (Exception ex) { return "#AI: " + ex.Message; }
            });
        }

        // ================================================================== Bahasa Indonesia
        [ExcelFunction(Name = "TERBILANG", Category = "MASTER BW Indonesia", Description = "Angka menjadi kata. =TERBILANG(1250) → seribu dua ratus lima puluh")]
        public static object TerbilangFn(double angka, object hurufBesar)
            => Bool(hurufBesar) ? Terbilang.ToTitle(Terbilang.Convert(angka)) : Terbilang.Convert(angka);

        [ExcelFunction(Name = "TERBILANG.RUPIAH", Category = "MASTER BW Indonesia", Description = "Angka menjadi terbilang rupiah untuk kuitansi. =TERBILANG.RUPIAH(1500000)")]
        public static object TerbilangRupiah(double angka) => Terbilang.Rupiah((decimal)Math.Round(angka, 2));

        [ExcelFunction(Name = "TANGGAL.INDO", Category = "MASTER BW Indonesia", Description = "Tanggal format Indonesia. format: 0 = 1 Januari 2025, 1 = Senin, 1 Januari 2025, 2 = 01/01/2025")]
        public static object TanggalIndo(double tanggal, object format)
        {
            var d = DateTime.FromOADate(tanggal);
            int f = format is double x ? (int)x : 0;
            return f == 1 ? IndoDate.WithDay(d) : f == 2 ? IndoDate.Short(d) : IndoDate.Long(d);
        }

        [ExcelFunction(Name = "HARI.INDO", Category = "MASTER BW Indonesia", Description = "Nama hari bahasa Indonesia dari tanggal.")]
        public static object HariIndo(double tanggal) => IndoDate.Hari[(int)DateTime.FromOADate(tanggal).DayOfWeek];

        [ExcelFunction(Name = "BULAN.INDO", Category = "MASTER BW Indonesia", Description = "Nama bulan bahasa Indonesia dari tanggal atau angka 1–12.")]
        public static object BulanIndo(double nilai)
        {
            int m = nilai >= 1 && nilai <= 12 ? (int)nilai : DateTime.FromOADate(nilai).Month;
            return IndoDate.Bulan[m - 1];
        }

        [ExcelFunction(Name = "ROMAWI", Category = "MASTER BW Indonesia", Description = "Angka Romawi (untuk nomor surat/BAB).")]
        public static object Romawi(double angka) => Roman.ToRoman((int)angka);

        [ExcelFunction(Name = "ANGKA.INDO", Category = "MASTER BW Indonesia", Description = "Ubah teks seperti \"Rp1.250.000,50\" / \"1,5 jt\" menjadi angka.")]
        public static object AngkaIndo(string teks) => IndoNumber.TryParse(teks, out var v) ? (object)v : ExcelError.ExcelErrorValue;

        [ExcelFunction(Name = "TANGGAL.TEKS", Category = "MASTER BW Indonesia", Description = "Ubah teks tanggal (\"12 Januari 2024\", \"12-01-24\") menjadi tanggal Excel.")]
        public static object TanggalTeks(string teks) => IndoDate.TryParse(teks, out var d) ? (object)d.ToOADate() : ExcelError.ExcelErrorValue;

        [ExcelFunction(Name = "NAMA.RAPI", Category = "MASTER BW Indonesia", Description = "Rapikan kapitalisasi nama orang (gelar dipertahankan).")]
        public static object NamaRapi(string nama) => IndoText.ProperName(nama);

        [ExcelFunction(Name = "HP.RAPI", Category = "MASTER BW Indonesia", Description = "Normalisasi nomor HP ke format +62.")]
        public static object HpRapi(object nomor) => IndoText.NormalizePhone(Text(nomor));

        [ExcelFunction(Name = "USIA", Category = "MASTER BW Indonesia", Description = "Usia/masa kerja: \"x tahun y bulan\" dari tanggal awal sampai hari ini (atau tanggal akhir).")]
        public static object Usia(double tanggalAwal, object tanggalAkhir)
        {
            var a = DateTime.FromOADate(tanggalAwal);
            var b = tanggalAkhir is double x ? DateTime.FromOADate(x) : DateTime.Today;
            int months = (b.Year - a.Year) * 12 + b.Month - a.Month - (b.Day < a.Day ? 1 : 0);
            return (months / 12) + " tahun " + (months % 12) + " bulan";
        }

        [ExcelFunction(Name = "PREDIKAT", Category = "MASTER BW Indonesia", Description = "Predikat nilai: A ≥ 90, B ≥ 80, C ≥ 70, selain itu D (batas bisa diubah).")]
        public static object Predikat(double nilai, object batasA, object batasB, object batasC)
        {
            double a = batasA is double x ? x : 90, b = batasB is double y ? y : 80, c = batasC is double z ? z : 70;
            return nilai >= a ? "A" : nilai >= b ? "B" : nilai >= c ? "C" : "D";
        }

        // ================================================================== Pajak & keuangan
        [ExcelFunction(Name = "PPH21.TAHUNAN", Category = "MASTER BW Keuangan", Description = "Estimasi PPh 21 setahun (tarif Pasal 17) dari gaji bruto bulanan. status: \"TK/0\", \"K/1\", …")]
        public static object Pph21(double brutoBulanan, object status)
        {
            var lines = TaxCalc.Pph21Annual(brutoBulanan, Text(status).Length > 0 ? Text(status) : "TK/0");
            return lines.First(l => l.Label.StartsWith("Total PPh 21")).Amount;
        }

        [ExcelFunction(Name = "PTKP", Category = "MASTER BW Keuangan", Description = "Nilai PTKP setahun untuk status \"TK/0\" s.d. \"K/3\".")]
        public static object Ptkp(string status) => TaxCalc.Ptkp(status);

        [ExcelFunction(Name = "PPN.HITUNG", Category = "MASTER BW Keuangan", Description = "PPN 12% atas DPP nilai lain 11/12 (efektif 11%). barang_mewah = TRUE untuk 12% penuh.")]
        public static object Ppn(double harga, object barangMewah) => TaxCalc.Ppn(harga, Bool(barangMewah))[2].Amount;

        [ExcelFunction(Name = "ANGSURAN", Category = "MASTER BW Keuangan", Description = "Angsuran per bulan. metode: \"anuitas\" (default), \"flat\", \"efektif\" (bulan pertama).")]
        public static object Angsuran(double pokok, double bungaPersenTahun, double tenorBulan, object metode)
        {
            var m = Text(metode).ToLowerInvariant();
            var method = m.StartsWith("flat") ? "flat" : m.StartsWith("efek") ? "effective" : "annuity";
            return FinanceCalc.LoanSchedule(pokok, bungaPersenTahun, (int)tenorBulan, method)[0].Payment;
        }

        // ================================================================== util
        private static string Text(object v)
        {
            if (v == null || v is ExcelMissing || v is ExcelEmpty) return "";
            if (v is object[,] arr) return string.Join(", ", arr.Cast<object>().Select(Text).Where(s => s.Length > 0));
            return Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture)?.Trim() ?? "";
        }

        private static bool Bool(object v) => v is bool b ? b : v is double d && d != 0;

        private static string Hash(string s)
        {
            using (var sha = SHA1.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(s))).Replace("-", "").Substring(0, 24);
        }
    }
}
