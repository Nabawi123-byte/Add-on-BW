using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace MasterBW.Core.Tools
{
    public static class ToolCatalog
    {
        private static readonly Lazy<List<ToolDef>> WordList = new Lazy<List<ToolDef>>(WordTools.Build);
        private static readonly Lazy<List<ToolDef>> ExcelList = new Lazy<List<ToolDef>>(ExcelTools.Build);

        public static IReadOnlyList<ToolDef> Word => WordList.Value;
        public static IReadOnlyList<ToolDef> Excel => ExcelList.Value;

        public static IReadOnlyList<ToolDef> For(OfficeApp app) => app == OfficeApp.Word ? Word : Excel;

        public static ToolDef Find(OfficeApp app, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return For(app).FirstOrDefault(t => string.Equals(t.Name, name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Pilih aksi paling relevan untuk sebuah perintah agar prompt tetap pendek
        /// (prompt pendek = jauh lebih cepat di CPU).
        /// </summary>
        public static List<ToolDef> Select(OfficeApp app, string request, int max = 10)
        {
            var text = Normalize(request);
            var words = new HashSet<string>(text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
            var scored = new List<(ToolDef tool, double score)>();
            foreach (var tool in For(app))
            {
                double score = 0;
                foreach (var kw in tool.Keywords)
                {
                    var k = Normalize(kw);
                    if (k.Length == 0) continue;
                    if (k.Contains(" "))
                    {
                        if (text.Contains(k)) score += 3;
                    }
                    else if (words.Contains(k)) score += 2;
                    else if (k.Length >= 4 && words.Any(w => w.StartsWith(k, StringComparison.Ordinal) || (w.Length >= 5 && k.StartsWith(w, StringComparison.Ordinal)))) score += 1;
                }
                if (text.Contains(tool.Name.Replace('_', ' '))) score += 3;
                if (Normalize(tool.Label).Length > 0 && text.Contains(Normalize(tool.Label))) score += 2;
                scored.Add((tool, score));
            }

            var picked = scored.Where(s => s.score > 0).OrderByDescending(s => s.score).Take(max).Select(s => s.tool).ToList();

            // Aksi umum yang selalu disertakan sebagai cadangan.
            var fallback = app == OfficeApp.Word
                ? new[] { "insert_text", "generate_content", "transform_text", "format_font", "format_paragraph" }
                : new[] { "write_formula", "write_values", "format_cells", "format_number", "vba_macro" };
            foreach (var name in fallback)
            {
                if (picked.Count >= max + 3) break;
                var t = Find(app, name);
                if (t != null && !picked.Contains(t)) picked.Add(t);
            }
            return picked;
        }

        internal static string Normalize(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (var ch in s.ToLowerInvariant().Normalize(NormalizationForm.FormD))
            {
                var cat = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (cat == UnicodeCategory.NonSpacingMark) continue;
                sb.Append(char.IsLetterOrDigit(ch) || ch == '#' || ch == '/' ? ch : ' ');
            }
            return string.Join(" ", sb.ToString().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
        }

        /// <summary>Validasi rencana: aksi dikenal, argumen wajib ada, nilai enum valid. Mengembalikan daftar masalah.</summary>
        public static List<string> Validate(OfficeApp app, ActionPlan plan)
        {
            var problems = new List<string>();
            if (plan?.Steps == null) return problems;
            for (int i = 0; i < plan.Steps.Count; i++)
            {
                var step = plan.Steps[i];
                var def = Find(app, step.Action);
                if (def == null)
                {
                    var msg = "Langkah " + (i + 1) + ": aksi '" + step.Action + "' tidak dikenal.";
                    step.Warnings.Add(msg);
                    step.Enabled = false;
                    problems.Add(msg);
                    continue;
                }
                step.Action = def.Name;
                if (step.Args == null) step.Args = new JObject();
                foreach (var p in def.Params)
                {
                    var val = step.Args[p.Name];
                    var missing = val == null || val.Type == JTokenType.Null || (val.Type == JTokenType.String && string.IsNullOrWhiteSpace((string)val));
                    if (missing)
                    {
                        if (p.Required)
                        {
                            var msg = "Langkah " + (i + 1) + " (" + def.Label + "): argumen '" + p.Name + "' belum diisi.";
                            step.Warnings.Add(msg);
                            problems.Add(msg);
                        }
                        continue;
                    }
                    if (p.Enum != null && p.Enum.Length > 0 && val.Type == JTokenType.String)
                    {
                        var s = (string)val;
                        var match = p.Enum.FirstOrDefault(e => string.Equals(e, s, StringComparison.OrdinalIgnoreCase));
                        if (match != null) step.Args[p.Name] = match;
                        else
                        {
                            var msg = "Langkah " + (i + 1) + " (" + def.Label + "): nilai '" + s + "' untuk '" + p.Name + "' tidak dikenal; dipakai '" + (p.Default ?? p.Enum[0]) + "'.";
                            step.Warnings.Add(msg);
                            step.Args[p.Name] = p.Default ?? p.Enum[0];
                        }
                    }
                }
                if (def.Destructive) step.Warnings.Add("Aksi ini menghapus data. Periksa sebelum menjalankan.");
            }
            return problems;
        }
    }
}
