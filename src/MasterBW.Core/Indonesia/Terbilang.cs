using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace MasterBW.Core.Indonesia
{
    /// <summary>Mengubah angka menjadi kata dalam bahasa Indonesia (terbilang), untuk kuitansi, cek, dan surat resmi.</summary>
    public static class Terbilang
    {
        private static readonly string[] Satuan = { "", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas" };
        private static readonly string[] Skala = { "", "ribu", "juta", "miliar", "triliun", "kuadriliun" };

        /// <summary>Contoh: 1250000 → "satu juta dua ratus lima puluh ribu".</summary>
        public static string Convert(decimal value)
        {
            if (value == 0) return "nol";
            var sb = new StringBuilder();
            if (value < 0)
            {
                sb.Append("minus ");
                value = -value;
            }
            var whole = decimal.Truncate(value);
            sb.Append(Whole(whole));
            var frac = value - whole;
            if (frac > 0)
            {
                var digits = frac.ToString(CultureInfo.InvariantCulture).TrimStart('0').TrimStart('.').TrimEnd('0');
                if (digits.Length > 0)
                {
                    sb.Append(" koma");
                    foreach (var d in digits) sb.Append(' ').Append(d == '0' ? "nol" : Satuan[d - '0']);
                }
            }
            return sb.ToString().Trim();
        }

        public static string Convert(double value) => Convert((decimal)Math.Round(value, 6));

        /// <summary>Contoh: 1500000.50 → "Satu Juta Lima Ratus Ribu Rupiah Lima Puluh Sen".</summary>
        public static string Rupiah(decimal value, bool titleCase = true, bool withSen = true)
        {
            var whole = decimal.Truncate(Math.Abs(value));
            var sen = (int)Math.Round((Math.Abs(value) - whole) * 100);
            if (sen == 100) { whole += 1; sen = 0; }
            var text = (value < 0 ? "minus " : "") + (whole == 0 ? "nol" : Whole(whole)) + " rupiah";
            if (withSen && sen > 0) text += " " + Whole(sen) + " sen";
            return titleCase ? ToTitle(text) : text;
        }

        private static string Whole(decimal n)
        {
            if (n == 0) return "";
            var groups = new List<int>();
            while (n > 0)
            {
                groups.Add((int)(n % 1000));
                n = decimal.Truncate(n / 1000);
            }
            if (groups.Count > Skala.Length) throw new ArgumentOutOfRangeException(nameof(n), "Angka terlalu besar untuk terbilang.");
            var parts = new List<string>();
            for (int i = groups.Count - 1; i >= 0; i--)
            {
                var g = groups[i];
                if (g == 0) continue;
                if (i == 1 && g == 1) parts.Add("seribu");
                else
                {
                    var w = Hundreds(g);
                    parts.Add(Skala[i].Length > 0 ? w + " " + Skala[i] : w);
                }
            }
            return string.Join(" ", parts);
        }

        private static string Hundreds(int n)
        {
            var parts = new List<string>();
            int h = n / 100, rest = n % 100;
            if (h == 1) parts.Add("seratus");
            else if (h > 1) parts.Add(Satuan[h] + " ratus");
            if (rest > 0)
            {
                if (rest < 12) parts.Add(Satuan[rest]);
                else if (rest < 20) parts.Add(Satuan[rest - 10] + " belas");
                else
                {
                    parts.Add(Satuan[rest / 10] + " puluh");
                    if (rest % 10 > 0) parts.Add(Satuan[rest % 10]);
                }
            }
            return string.Join(" ", parts);
        }

        public static string ToTitle(string s)
        {
            return string.Join(" ", s.Split(' ').Select(w => w.Length == 0 ? w : char.ToUpperInvariant(w[0]) + w.Substring(1)));
        }
    }

    /// <summary>Angka Romawi (untuk BAB, nomor surat, bulan pada nomor surat).</summary>
    public static class Roman
    {
        private static readonly (int v, string s)[] Map =
        {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
        };

        public static string ToRoman(int n)
        {
            if (n <= 0 || n > 3999) return n.ToString(CultureInfo.InvariantCulture);
            var sb = new StringBuilder();
            foreach (var (v, s) in Map)
                while (n >= v) { sb.Append(s); n -= v; }
            return sb.ToString();
        }

        public static int FromRoman(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            s = s.Trim().ToUpperInvariant();
            int total = 0, i = 0;
            foreach (var (v, r) in Map)
                while (i + r.Length <= s.Length && s.Substring(i, r.Length) == r) { total += v; i += r.Length; }
            return i == s.Length ? total : 0;
        }
    }
}
