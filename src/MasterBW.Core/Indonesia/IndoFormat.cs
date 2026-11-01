using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MasterBW.Core.Indonesia
{
    /// <summary>Parsing &amp; format angka gaya Indonesia (titik ribuan, koma desimal, "Rp", "jt", "M").</summary>
    public static class IndoNumber
    {
        public static readonly CultureInfo Id = CreateCulture();

        private static CultureInfo CreateCulture()
        {
            try { return CultureInfo.GetCultureInfo("id-ID"); }
            catch (CultureNotFoundException) { return CultureInfo.InvariantCulture; }
        }

        /// <summary>
        /// Terima "Rp 1.250.000,50", "1,250,000.50", "1.5 jt", "2,3 M", "15%", "(1.000)" (negatif akuntansi).
        /// </summary>
        public static bool TryParse(string input, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(input)) return false;
            var s = input.Trim().ToLowerInvariant();
            bool negative = false;
            if (s.StartsWith("(") && s.EndsWith(")")) { negative = true; s = s.Substring(1, s.Length - 2); }
            s = s.Replace("rp.", "").Replace("rp", "").Replace("idr", "").Replace(" ", "").Replace(" ", "");
            if (s.StartsWith("-")) { negative = !negative; s = s.Substring(1); }

            double multiplier = 1;
            bool percent = false;
            if (s.EndsWith("%")) { percent = true; s = s.TrimEnd('%'); }
            var suffixes = new[] { ("triliun", 1e12), ("miliar", 1e9), ("milyar", 1e9), ("juta", 1e6), ("jt", 1e6), ("ribu", 1e3), ("rb", 1e3), ("k", 1e3), ("m", 1e9), ("t", 1e12) };
            foreach (var (suf, mul) in suffixes)
            {
                if (s.EndsWith(suf) && s.Length > suf.Length && char.IsDigit(s[s.Length - suf.Length - 1]))
                {
                    multiplier = mul;
                    s = s.Substring(0, s.Length - suf.Length);
                    break;
                }
            }
            s = s.TrimEnd(',', '-').Replace(",-", "");
            if (s.Length == 0 || !Regex.IsMatch(s, @"^[0-9.,]+$")) return false;

            int dots = s.Count(c => c == '.'), commas = s.Count(c => c == ',');
            string normalized;
            if (dots > 0 && commas > 0)
            {
                // pemisah yang muncul terakhir adalah desimal
                normalized = s.LastIndexOf(',') > s.LastIndexOf('.')
                    ? s.Replace(".", "").Replace(',', '.')
                    : s.Replace(",", "");
            }
            else if (commas > 0)
            {
                // "1,5" → desimal; "1,250,000" → ribuan
                normalized = commas == 1 && s.Length - s.IndexOf(',') - 1 != 3 ? s.Replace(',', '.') : s.Replace(",", "");
            }
            else if (dots > 0)
            {
                // "1.250.000" → ribuan (gaya Indonesia); "1.5" → desimal
                normalized = dots == 1 && s.Length - s.IndexOf('.') - 1 != 3 ? s : s.Replace(".", "");
            }
            else normalized = s;

            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
            value *= multiplier;
            if (percent) value /= 100.0;
            if (negative) value = -value;
            return true;
        }

        public static string Rupiah(double value, int decimals = 0)
        {
            var abs = Math.Abs(value).ToString("N" + decimals, Id);
            return (value < 0 ? "-Rp" : "Rp") + abs;
        }

        public static string Format(double value, int decimals = 0) => value.ToString("N" + decimals, Id);
    }

    /// <summary>Tanggal gaya Indonesia: "Senin, 1 Januari 2024", nomor surat bulan Romawi, parsing tanggal berantakan.</summary>
    public static class IndoDate
    {
        public static readonly string[] Bulan = { "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember" };
        public static readonly string[] Hari = { "Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu" };

        public static string Long(DateTime d) => d.Day + " " + Bulan[d.Month - 1] + " " + d.Year;
        public static string WithDay(DateTime d) => Hari[(int)d.DayOfWeek] + ", " + Long(d);
        public static string Short(DateTime d) => d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

        /// <summary>Format bulan Romawi untuk nomor surat, mis. 005/SDN-1/IV/2025.</summary>
        public static string RomanMonth(DateTime d) => Roman.ToRoman(d.Month);

        private static readonly (string key, int month)[] MonthNames =
        {
            ("januari", 1), ("jan", 1), ("februari", 2), ("pebruari", 2), ("feb", 2), ("peb", 2), ("maret", 3), ("mar", 3),
            ("april", 4), ("apr", 4), ("mei", 5), ("may", 5), ("juni", 6), ("jun", 6), ("juli", 7), ("jul", 7),
            ("agustus", 8), ("agu", 8), ("ags", 8), ("aug", 8), ("agt", 8), ("september", 9), ("sep", 9), ("sept", 9),
            ("oktober", 10), ("okt", 10), ("oct", 10), ("november", 11), ("nopember", 11), ("nov", 11), ("nop", 11),
            ("desember", 12), ("des", 12), ("dec", 12)
        };

        /// <summary>
        /// Parsing tanggal dari banyak format umum di data Indonesia:
        /// "12 Januari 2024", "12-01-2024", "12/1/24", "2024-01-12", "Senin, 12 Jan 2024", serial Excel (45303).
        /// Format angka diasumsikan hari/bulan/tahun (gaya Indonesia).
        /// </summary>
        public static bool TryParse(string input, out DateTime date)
        {
            date = default;
            if (string.IsNullOrWhiteSpace(input)) return false;
            var s = input.Trim().ToLowerInvariant();
            s = Regex.Replace(s, @"^(senin|selasa|rabu|kamis|jumat|jum'at|sabtu|minggu|ahad)\s*,?\s*", "");

            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial) && serial > 20000 && serial < 80000 && !s.Contains("."))
            {
                date = DateTime.FromOADate(serial);
                return true;
            }

            var iso = Regex.Match(s, @"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})");
            if (iso.Success) return Make(int.Parse(iso.Groups[1].Value), int.Parse(iso.Groups[2].Value), int.Parse(iso.Groups[3].Value), out date);

            var dmy = Regex.Match(s, @"^(\d{1,2})[-/.\s](\d{1,2})[-/.\s](\d{2,4})$");
            if (dmy.Success) return Make(Year(dmy.Groups[3].Value), int.Parse(dmy.Groups[2].Value), int.Parse(dmy.Groups[1].Value), out date);

            var named = Regex.Match(s, @"^(\d{1,2})\s*[-/ ]?\s*([a-z]+)\.?\s*[-/ ,]?\s*(\d{2,4})$");
            if (named.Success)
            {
                var m = MonthFromName(named.Groups[2].Value);
                if (m > 0) return Make(Year(named.Groups[3].Value), m, int.Parse(named.Groups[1].Value), out date);
            }

            var monthFirst = Regex.Match(s, @"^([a-z]+)\.?\s+(\d{1,2}),?\s+(\d{2,4})$");
            if (monthFirst.Success)
            {
                var m = MonthFromName(monthFirst.Groups[1].Value);
                if (m > 0) return Make(Year(monthFirst.Groups[3].Value), m, int.Parse(monthFirst.Groups[2].Value), out date);
            }
            return false;
        }

        public static int MonthFromName(string name)
        {
            name = (name ?? "").Trim().ToLowerInvariant().TrimEnd('.');
            foreach (var (key, month) in MonthNames)
                if (name == key) return month;
            foreach (var (key, month) in MonthNames)
                if (key.Length >= 3 && name.StartsWith(key)) return month;
            return 0;
        }

        private static int Year(string y)
        {
            var v = int.Parse(y);
            if (y.Length == 2) v += v < 70 ? 2000 : 1900;
            return v;
        }

        private static bool Make(int y, int m, int d, out DateTime date)
        {
            date = default;
            if (m < 1 || m > 12 || d < 1 || d > DateTime.DaysInMonth(Math.Max(1, Math.Min(9999, y)), m) || y < 1900 || y > 2200) return false;
            date = new DateTime(y, m, d);
            return true;
        }
    }

    /// <summary>Normalisasi data Indonesia lain: nomor HP, nama, NIK.</summary>
    public static class IndoText
    {
        /// <summary>"0812-3456-789" / "+62 812 3456 789" / "62812..." → "+628123456789".</summary>
        public static string NormalizePhone(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return input;
            var digits = new string(input.Where(char.IsDigit).ToArray());
            if (digits.Length < 8) return input.Trim();
            if (digits.StartsWith("62")) return "+" + digits;
            if (digits.StartsWith("0")) return "+62" + digits.Substring(1);
            if (digits.StartsWith("8")) return "+62" + digits;
            return digits;
        }

        private static readonly string[] LowerParticles = { "bin", "binti", "van", "der", "de", "al", "el" };
        private static readonly string[] Degrees = { "S.Pd.", "S.Pd", "M.Pd.", "M.Pd", "S.E.", "S.E", "S.H.", "S.Kom.", "S.T.", "M.M.", "M.Si.", "Dr.", "Prof.", "Ir.", "S.Sos.", "M.Kes.", "S.Kep.", "A.Md." };

        /// <summary>Kapitalisasi nama orang yang benar (gelar akademik dipertahankan).</summary>
        public static string ProperName(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return input;
            var parts = Regex.Split(input.Trim(), @"\s+");
            var sb = new StringBuilder();
            for (int i = 0; i < parts.Length; i++)
            {
                var p = parts[i];
                var degree = Degrees.FirstOrDefault(d => string.Equals(d.TrimEnd('.'), p.TrimEnd('.', ','), StringComparison.OrdinalIgnoreCase));
                string word;
                if (degree != null) word = degree + (p.EndsWith(",") ? "," : "");
                else if (i > 0 && LowerParticles.Contains(p.ToLowerInvariant())) word = p.ToLowerInvariant();
                else word = string.Join("-", p.Split('-').Select(x => x.Length == 0 ? x : char.ToUpperInvariant(x[0]) + x.Substring(1).ToLowerInvariant()));
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(word);
            }
            return sb.ToString();
        }

        /// <summary>Hapus spasi ganda, spasi sebelum tanda baca, karakter tak tercetak.</summary>
        public static string CleanSpaces(string input)
        {
            if (input == null) return null;
            var s = new string(input.Where(c => !char.IsControl(c) || c == '\n' || c == '\t').ToArray());
            s = s.Replace(' ', ' ');
            s = Regex.Replace(s, @"[ \t]{2,}", " ");
            s = Regex.Replace(s, @" +([,.;:!?%)])", "$1");
            s = Regex.Replace(s, @"\( +", "(");
            return s.Trim();
        }
    }
}
