using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MasterBW.Core.ExcelKit
{
    public sealed class FormulaToken
    {
        /// <summary>function | ref | number | string | operator | paren | comma | name | error | bool</summary>
        public string Kind { get; set; }
        public string Text { get; set; }
        public override string ToString() => Kind + ":" + Text;
    }

    /// <summary>Tokenizer, penjelas, pemeriksa kompatibilitas, dan diagnosis error rumus Excel — offline tanpa AI.</summary>
    public static class FormulaTools
    {
        private static readonly Regex TokenRegex = new Regex(
            @"(?<string>""(?:[^""]|"""")*"")" +
            @"|(?<error>#(?:N/A|REF!|VALUE!|DIV/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!))" +
            @"|(?<ref>(?:(?:'[^']+'|[A-Za-z_][\w\.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(])|(?:(?:'[^']+'|[A-Za-z_][\w\.]*)!)?\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w(])|(?:(?:'[^']+'|[A-Za-z_][\w\.]*)!)?\$?\d+:\$?\d+)" +
            @"|(?<function>[A-Za-z_][A-Za-z0-9_\.]*)(?=\()" +
            @"|(?<bool>\b(?:TRUE|FALSE)\b)" +
            @"|(?<number>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?%?)" +
            @"|(?<name>[A-Za-z_\\][A-Za-z0-9_\.]*)" +
            @"|(?<operator><>|>=|<=|[+\-*/^&=<>%])" +
            @"|(?<paren>[()])" +
            @"|(?<comma>[,;])",
            RegexOptions.Compiled);

        public static List<FormulaToken> Tokenize(string formula)
        {
            var list = new List<FormulaToken>();
            if (string.IsNullOrEmpty(formula)) return list;
            var f = formula.TrimStart('=');
            foreach (Match m in TokenRegex.Matches(f))
            {
                foreach (var kind in new[] { "string", "error", "ref", "function", "bool", "number", "name", "operator", "paren", "comma" })
                {
                    if (m.Groups[kind].Success)
                    {
                        list.Add(new FormulaToken { Kind = kind, Text = m.Value });
                        break;
                    }
                }
            }
            return list;
        }

        public static List<string> FunctionsUsed(string formula)
            => Tokenize(formula).Where(t => t.Kind == "function").Select(t => t.Text.ToUpperInvariant().Replace("_XLFN.", "")).Distinct().ToList();

        /// <summary>Daftar fungsi di rumus yang tidak didukung Excel 2016 + sarannya.</summary>
        public static List<string> CompatibilityIssues(string formula)
        {
            var result = new List<string>();
            foreach (var fn in FunctionsUsed(formula))
            {
                if (!ExcelFunctionInfo.IsUnavailableIn2016(fn)) continue;
                var alt = ExcelFunctionInfo.Alternatives.TryGetValue(fn, out var a) ? " Ganti dengan: " + a + "." : "";
                result.Add("Fungsi " + fn + " tidak tersedia di Excel 2016." + alt);
            }
            return result;
        }

        /// <summary>Periksa kurung seimbang &amp; tanda kutip. null jika aman.</summary>
        public static string SyntaxProblem(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) return "Rumus kosong.";
            if (!formula.TrimStart().StartsWith("=")) return "Rumus harus diawali tanda =.";
            int depth = 0; bool inStr = false;
            foreach (var c in formula)
            {
                if (c == '"') inStr = !inStr;
                else if (!inStr && c == '(') depth++;
                else if (!inStr && c == ')') { depth--; if (depth < 0) return "Kurung tutup berlebih."; }
            }
            if (inStr) return "Tanda kutip tidak ditutup.";
            if (depth > 0) return "Kurang " + depth + " kurung tutup.";
            return null;
        }

        /// <summary>Normalisasi rumus dari AI: pemisah ; → , (di luar string), hapus spasi pinggir, pastikan diawali =.</summary>
        public static string Normalize(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) return formula;
            var f = formula.Trim().Trim('`');
            if (!f.StartsWith("=")) f = "=" + f;
            var sb = new StringBuilder();
            bool inStr = false;
            foreach (var c in f)
            {
                if (c == '"') inStr = !inStr;
                if (!inStr && c == ';') sb.Append(',');
                else if (!inStr && (c == '“' || c == '”')) sb.Append('"');
                else sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>Penjelasan rumus langkah demi langkah dalam bahasa Indonesia.</summary>
        public static string Explain(string formula)
        {
            var tokens = Tokenize(formula);
            if (tokens.Count == 0) return "Sel ini tidak berisi rumus.";
            var sb = new StringBuilder();
            sb.AppendLine("Rumus: " + formula);
            var fns = FunctionsUsed(formula);
            if (fns.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Fungsi yang dipakai:");
                foreach (var fn in fns)
                {
                    var info = ExcelFunctionInfo.Get(fn);
                    sb.AppendLine(info != null
                        ? "• " + info.Name + " — " + info.Description + " Sintaks: " + info.Syntax
                        : "• " + fn + (ExcelFunctionInfo.IsUnavailableIn2016(fn) ? " — TIDAK tersedia di Excel 2016." : " — fungsi Excel."));
                }
            }
            var refs = tokens.Where(t => t.Kind == "ref").Select(t => t.Text).Distinct().ToList();
            if (refs.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Sel/range yang dirujuk: " + string.Join(", ", refs));
                if (refs.Any(r => r.Contains("$"))) sb.AppendLine("Tanda $ berarti referensi absolut (tidak bergeser saat rumus disalin).");
            }
            var crit = tokens.Where(t => t.Kind == "string").Select(t => t.Text).Distinct().ToList();
            if (crit.Count > 0) sb.AppendLine("Teks/kriteria: " + string.Join(", ", crit));
            var ops = tokens.Where(t => t.Kind == "operator").Select(t => t.Text).Distinct().ToList();
            if (ops.Count > 0)
            {
                var names = new Dictionary<string, string> { ["+"] = "tambah", ["-"] = "kurang", ["*"] = "kali", ["/"] = "bagi", ["^"] = "pangkat", ["&"] = "gabung teks", ["="] = "sama dengan", ["<>"] = "tidak sama dengan", [">"] = "lebih dari", ["<"] = "kurang dari", [">="] = "lebih dari atau sama dengan", ["<="] = "kurang dari atau sama dengan", ["%"] = "persen" };
                sb.AppendLine("Operator: " + string.Join(", ", ops.Select(o => o + " (" + (names.TryGetValue(o, out var n) ? n : o) + ")")));
            }
            var outer = tokens.FirstOrDefault();
            if (outer != null && outer.Kind == "function")
            {
                var info = ExcelFunctionInfo.Get(outer.Text);
                if (info != null) { sb.AppendLine(); sb.AppendLine("Inti: hasil akhir ditentukan oleh " + info.Name + " — " + info.Description); }
            }
            foreach (var issue in CompatibilityIssues(formula)) sb.AppendLine("⚠ " + issue);
            return sb.ToString().TrimEnd();
        }

        /// <summary>Diagnosis penyebab umum error Excel.</summary>
        public static string DiagnoseError(string errorText, string formula)
        {
            var fns = FunctionsUsed(formula ?? "");
            switch ((errorText ?? "").ToUpperInvariant())
            {
                case "#N/A":
                    return fns.Any(f => f == "VLOOKUP" || f == "HLOOKUP" || f == "MATCH" || f == "LOOKUP")
                        ? "#N/A: nilai yang dicari tidak ditemukan. Penyebab umum: beda spasi (gunakan TRIM), angka tersimpan sebagai teks, range tabel tidak dikunci dengan $, atau argumen terakhir VLOOKUP bukan FALSE. Bungkus dengan IFERROR(...,\"Tidak ditemukan\") bila wajar tidak ada."
                        : "#N/A: data tidak tersedia untuk perhitungan.";
                case "#REF!":
                    return "#REF!: rumus merujuk sel/baris/kolom/sheet yang sudah dihapus, atau nomor kolom VLOOKUP melebihi lebar tabel.";
                case "#VALUE!":
                    return "#VALUE!: tipe data salah, misalnya menjumlahkan teks dengan angka atau tanggal berupa teks. Bersihkan data (text_to_number/normalize_date) atau gunakan VALUE().";
                case "#DIV/0!":
                    return "#DIV/0!: pembagian dengan nol atau sel kosong. Gunakan IFERROR(rumus,0) atau IF(pembagi=0,0,rumus).";
                case "#NAME?":
                    var bad = fns.Where(f => ExcelFunctionInfo.IsUnavailableIn2016(f)).ToList();
                    return bad.Count > 0
                        ? "#NAME?: fungsi " + string.Join(", ", bad) + " tidak dikenal Excel 2016. " + string.Join(" ", CompatibilityIssues(formula))
                        : "#NAME?: nama fungsi salah ketik, nama range belum dibuat, atau teks tanpa tanda kutip.";
                case "#NUM!":
                    return "#NUM!: angka tidak valid (mis. akar bilangan negatif, IRR tidak konvergen, angka terlalu besar).";
                case "#NULL!":
                    return "#NULL!: dua range tidak berpotongan; mungkin kurang tanda koma atau titik dua.";
                default:
                    var syn = SyntaxProblem(formula);
                    return syn ?? "Tidak ada error yang terdeteksi secara otomatis.";
            }
        }

        /// <summary>Konversi nomor kolom (1-based) ↔ huruf.</summary>
        public static string ColumnLetter(int col)
        {
            var s = "";
            while (col > 0) { int m = (col - 1) % 26; s = (char)('A' + m) + s; col = (col - m) / 26; }
            return s;
        }

        public static int ColumnNumber(string letters)
        {
            int n = 0;
            foreach (var c in (letters ?? "").ToUpperInvariant().Where(char.IsLetter)) n = n * 26 + (c - 'A' + 1);
            return n;
        }
    }
}
