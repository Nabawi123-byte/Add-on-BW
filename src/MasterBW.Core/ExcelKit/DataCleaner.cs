using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using MasterBW.Core.Indonesia;

namespace MasterBW.Core.ExcelKit
{
    /// <summary>Pembersih data massal yang bekerja pada array nilai Excel (object[,]) — cepat karena tanpa COM per sel.</summary>
    public static class DataCleaner
    {
        public static readonly string[] Operations =
        {
            "trim", "proper", "proper_name", "upper", "lower", "text_to_number", "normalize_date", "normalize_phone",
            "remove_nonprintable", "remove_extra_spaces", "fix_rupiah"
        };

        /// <summary>Bersihkan satu nilai sesuai daftar operasi. Mengembalikan nilai baru (string/double/DateTime).</summary>
        public static object CleanValue(object value, IEnumerable<string> operations, out bool changed)
        {
            changed = false;
            if (value == null) return null;
            object current = value;
            foreach (var raw in operations)
            {
                var op = (raw ?? "").Trim().ToLowerInvariant();
                if (!(current is string s)) continue;
                object next = current;
                switch (op)
                {
                    case "trim":
                    case "remove_extra_spaces":
                        next = Regex.Replace(s.Replace(' ', ' '), @"\s{2,}", " ").Trim();
                        break;
                    case "remove_nonprintable":
                        next = new string(s.Where(c => !char.IsControl(c)).ToArray()).Replace("​", "");
                        break;
                    case "upper": next = s.ToUpperInvariant(); break;
                    case "lower": next = s.ToLowerInvariant(); break;
                    case "proper":
                        next = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(s.ToLowerInvariant());
                        break;
                    case "proper_name":
                        next = IndoText.ProperName(s);
                        break;
                    case "text_to_number":
                    case "fix_rupiah":
                        if (IndoNumber.TryParse(s, out var d)) next = d;
                        break;
                    case "normalize_date":
                        if (IndoDate.TryParse(s, out var dt)) next = dt;
                        break;
                    case "normalize_phone":
                        if (Regex.IsMatch(s, @"^[\s\d+\-().]{8,}$")) next = IndoText.NormalizePhone(s);
                        break;
                }
                if (!Equals(next, current)) { changed = true; current = next; }
            }
            return current;
        }

        /// <summary>Bersihkan seluruh array 2D. Mengembalikan jumlah sel yang berubah.</summary>
        public static int CleanGrid(object[,] grid, IList<string> operations)
        {
            int changedCount = 0;
            int r0 = grid.GetLowerBound(0), r1 = grid.GetUpperBound(0), c0 = grid.GetLowerBound(1), c1 = grid.GetUpperBound(1);
            for (int r = r0; r <= r1; r++)
                for (int c = c0; c <= c1; c++)
                {
                    var v = CleanValue(grid[r, c], operations, out var ch);
                    if (!ch) continue;
                    grid[r, c] = v is DateTime dt ? (object)dt.ToOADate() : v;
                    changedCount++;
                }
            return changedCount;
        }

        /// <summary>Pisahkan teks berdasarkan pemisah (mendukung nama pemisah dalam bahasa Indonesia).</summary>
        public static string[] Split(string text, string delimiter)
        {
            if (text == null) return new string[0];
            var d = (delimiter ?? ",").Trim().ToLowerInvariant();
            string[] seps;
            switch (d)
            {
                case "koma": case "comma": seps = new[] { "," }; break;
                case "titik koma": case "semicolon": seps = new[] { ";" }; break;
                case "spasi": case "space": case "": seps = new[] { " " }; break;
                case "tab": seps = new[] { "\t" }; break;
                case "strip": case "dash": seps = new[] { "-" }; break;
                case "garis miring": case "slash": seps = new[] { "/" }; break;
                case "pipe": seps = new[] { "|" }; break;
                default: seps = new[] { delimiter }; break;
            }
            return text.Split(seps, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
        }

        /// <summary>Konversi nilai sel menjadi angka (mendukung teks Rupiah). NaN jika bukan angka.</summary>
        public static double ToNumber(object v)
        {
            switch (v)
            {
                case null: return double.NaN;
                case double d: return d;
                case int i: return i;
                case decimal m: return (double)m;
                case bool b: return b ? 1 : 0;
                case DateTime dt: return dt.ToOADate();
                case string s: return IndoNumber.TryParse(s, out var x) ? x : double.NaN;
                default:
                    return double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out var y) ? y : double.NaN;
            }
        }
    }

    /// <summary>Pemeriksa keamanan kode VBA buatan AI sebelum dipasang ke workbook.</summary>
    public static class VbaSafety
    {
        private static readonly (string pattern, string risk)[] Rules =
        {
            (@"\bKill\b", "menghapus file (Kill)"),
            (@"\bRmDir\b", "menghapus folder (RmDir)"),
            (@"\bShell\b", "menjalankan program lain (Shell)"),
            (@"CreateObject\s*\(\s*""WScript\.Shell""", "akses WScript.Shell"),
            (@"CreateObject\s*\(\s*""Scripting\.FileSystemObject""", "akses sistem file (FileSystemObject)"),
            (@"MSXML2\.|WinHttp|XMLHTTP|URLDownloadToFile", "akses internet"),
            (@"RegWrite|RegDelete|SaveSetting|DeleteSetting", "mengubah registry/pengaturan Windows"),
            (@"\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\b", "memanggil API Windows (Declare)"),
            (@"VBProject|VBComponents", "mengubah kode VBA lain"),
            (@"\.Delete\b", "menghapus objek (sheet/range/file)"),
            (@"Application\.DisplayAlerts\s*=\s*False", "mematikan peringatan Excel"),
            (@"\bSendKeys\b", "mengirim ketukan keyboard"),
            (@"\.SaveAs\b|\.SaveCopyAs\b", "menyimpan file ke lokasi lain"),
            (@"Environ\s*\(", "membaca variabel lingkungan Windows"),
        };

        public static List<string> Scan(string code)
        {
            var risks = new List<string>();
            if (string.IsNullOrEmpty(code)) return risks;
            foreach (var (pattern, risk) in Rules)
                if (Regex.IsMatch(code, pattern, RegexOptions.IgnoreCase)) risks.Add(risk);
            return risks;
        }

        /// <summary>Ambil kode dari blok ```vba ... ``` bila ada.</summary>
        public static string ExtractCode(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var m = Regex.Match(text, @"```(?:vba|vb|basic)?\s*\n(.*?)```", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            return (m.Success ? m.Groups[1].Value : text).Trim();
        }

        /// <summary>Nama Sub pertama dalam kode.</summary>
        public static string FirstSubName(string code)
        {
            var m = Regex.Match(code ?? "", @"^\s*(?:Public\s+|Private\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(", RegexOptions.Multiline | RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value : null;
        }
    }
}
