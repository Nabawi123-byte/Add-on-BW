using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace MasterBW.Office.ForWord
{
    /// <summary>Utilitas Word: warna, satuan, sisip teks bergaya markdown ringan.</summary>
    internal static class WordHelpers
    {
        public static float Cm(double cm) => (float)(cm * 28.3464567);

        /// <summary>Warna Word (BGR) dari nama Indonesia/Inggris atau #RRGGBB.</summary>
        public static int? ParseColor(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Trim().ToLowerInvariant();
            var named = new Dictionary<string, string>
            {
                ["hitam"] = "#000000", ["black"] = "#000000", ["putih"] = "#FFFFFF", ["white"] = "#FFFFFF",
                ["merah"] = "#C00000", ["red"] = "#C00000", ["biru"] = "#1F4E79", ["blue"] = "#1F4E79",
                ["hijau"] = "#00B050", ["green"] = "#00B050", ["kuning"] = "#FFFF00", ["yellow"] = "#FFFF00",
                ["oranye"] = "#ED7D31", ["orange"] = "#ED7D31", ["ungu"] = "#7030A0", ["purple"] = "#7030A0",
                ["abu-abu"] = "#808080", ["abu"] = "#808080", ["gray"] = "#808080", ["biru muda"] = "#BDD7EE",
                ["hijau muda"] = "#C6EFCE", ["merah muda"] = "#FFC7CE", ["pink"] = "#FFC7CE", ["abu muda"] = "#D9D9D9",
                ["biru tua"] = "#1F3864", ["emas"] = "#BF9000", ["coklat"] = "#843C0C",
            };
            if (named.TryGetValue(s, out var hex)) s = hex;
            if (!s.StartsWith("#") || s.Length != 7) return null;
            int r = int.Parse(s.Substring(1, 2), NumberStyles.HexNumber);
            int g = int.Parse(s.Substring(3, 2), NumberStyles.HexNumber);
            int b = int.Parse(s.Substring(5, 2), NumberStyles.HexNumber);
            return r + g * 256 + b * 65536;
        }

        public static Word.WdColorIndex Highlight(string s)
        {
            switch ((s ?? "").Trim().ToLowerInvariant())
            {
                case "hijau": case "green": return Word.WdColorIndex.wdBrightGreen;
                case "biru": case "blue": return Word.WdColorIndex.wdTurquoise;
                case "merah": case "red": return Word.WdColorIndex.wdRed;
                case "pink": case "merah muda": return Word.WdColorIndex.wdPink;
                case "none": case "tidak": case "hapus": return Word.WdColorIndex.wdNoHighlight;
                default: return Word.WdColorIndex.wdYellow;
            }
        }

        public static void ResizeKeepRatio(Word.InlineShape pic, float? width = null, float? height = null)
        {
            float ratio = pic.Height > 0 ? pic.Width / pic.Height : 1;
            if (height.HasValue) { pic.Height = height.Value; pic.Width = height.Value * ratio; }
            else if (width.HasValue) { pic.Width = width.Value; pic.Height = width.Value / ratio; }
        }

        public static Word.Style BuiltIn(Word.Document doc, Word.WdBuiltinStyle style) => doc.Styles[style];

        public static Word.WdBuiltinStyle HeadingStyle(int level)
        {
            switch (Math.Max(1, Math.Min(level, 4)))
            {
                case 1: return Word.WdBuiltinStyle.wdStyleHeading1;
                case 2: return Word.WdBuiltinStyle.wdStyleHeading2;
                case 3: return Word.WdBuiltinStyle.wdStyleHeading3;
                default: return Word.WdBuiltinStyle.wdStyleHeading4;
            }
        }

        private struct Span { public int Start, Length; public bool Bold, Italic; }

        /// <summary>Pisahkan **tebal** dan *miring* menjadi teks polos + daftar rentang format.</summary>
        private static string ParseInline(string line, List<Span> spans)
        {
            var result = new System.Text.StringBuilder();
            var rx = new Regex(@"\*\*(.+?)\*\*|\*(?!\s)(.+?)\*|__(.+?)__");
            int last = 0;
            foreach (Match m in rx.Matches(line))
            {
                result.Append(line, last, m.Index - last);
                var inner = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value;
                spans.Add(new Span { Start = result.Length, Length = inner.Length, Bold = m.Groups[1].Success || m.Groups[3].Success, Italic = m.Groups[2].Success });
                result.Append(inner);
                last = m.Index + m.Length;
            }
            result.Append(line, last, line.Length - last);
            return result.ToString();
        }

        /// <summary>
        /// Sisipkan satu paragraf di posisi <paramref name="pos"/> dengan style tertentu, mendukung **tebal**/*miring*.
        /// Mengembalikan range paragraf yang disisipkan (tanpa tanda paragraf) dan memajukan pos.
        /// </summary>
        public static Word.Range InsertParagraph(Word.Document doc, ref int pos, string text, object style = null,
                                                 Word.WdParagraphAlignment? align = null, bool bold = false)
        {
            var spans = new List<Span>();
            var plain = ParseInline(text ?? "", spans);
            var r = doc.Range(pos, pos);
            r.InsertAfter(plain);
            int start = r.Start, end = r.End;
            r.InsertParagraphAfter();
            var para = doc.Range(start, end);
            var full = doc.Range(start, end + 1);
            if (style != null) full.set_Style(ref style);
            para.Font.Bold = bold ? 1 : 0;
            para.Font.Italic = 0;
            if (align.HasValue) full.ParagraphFormat.Alignment = align.Value;
            foreach (var s in spans)
            {
                var sub = doc.Range(start + s.Start, start + s.Start + s.Length);
                if (s.Bold) sub.Font.Bold = 1;
                if (s.Italic) sub.Font.Italic = 1;
            }
            pos = end + 1;
            return para;
        }

        /// <summary>Sisipkan teks markdown ringan (judul, poin, nomor, paragraf). pos dimajukan ke akhir sisipan.</summary>
        public static void InsertMarkdown(Word.Document doc, ref int pos, string markdown)
        {
            var lines = (markdown ?? "").Replace("\r\n", "\n").Split('\n');
            object normal = Word.WdBuiltinStyle.wdStyleNormal;
            object bullet = Word.WdBuiltinStyle.wdStyleListBullet;
            object number = Word.WdBuiltinStyle.wdStyleListNumber;
            bool lastBlank = false;
            foreach (var raw in lines)
            {
                var line = raw.TrimEnd();
                if (line.Trim().Length == 0) { lastBlank = true; continue; }
                lastBlank = false;
                var t = line.TrimStart();
                Match m;
                if ((m = Regex.Match(t, @"^(#{1,4})\s+(.*)$")).Success)
                {
                    object h = HeadingStyle(m.Groups[1].Value.Length);
                    InsertParagraph(doc, ref pos, m.Groups[2].Value.Trim().Trim('*'), h);
                }
                else if ((m = Regex.Match(t, @"^[-*•]\s+(.*)$")).Success) InsertParagraph(doc, ref pos, m.Groups[1].Value, bullet);
                else if ((m = Regex.Match(t, @"^\d{1,3}[.)]\s+(.*)$")).Success) InsertParagraph(doc, ref pos, m.Groups[1].Value, number);
                else InsertParagraph(doc, ref pos, t, normal);
            }
            _ = lastBlank;
        }

        public static Word.Range InsertionRange(Word.Application app, Word.Document doc, string position)
        {
            switch ((position ?? "cursor").ToLowerInvariant())
            {
                case "start": return doc.Range(0, 0);
                case "end":
                {
                    int end = doc.Content.End - 1;
                    var r = doc.Range(end, end);
                    // pastikan mulai di paragraf baru
                    if (end > 0 && doc.Range(end - 1, end).Text != "\r") { r.InsertParagraphAfter(); end = doc.Content.End - 1; r = doc.Range(end, end); }
                    return r;
                }
                case "replace_selection":
                {
                    var sel = app.Selection.Range;
                    if (sel.Start != sel.End) sel.Text = "";
                    return doc.Range(sel.Start, sel.Start);
                }
                default:
                {
                    var sel = app.Selection.Range;
                    int p = sel.End;
                    return doc.Range(p, p);
                }
            }
        }

        public static Word.Range ScopeRange(Word.Application app, Word.Document doc, string scope)
        {
            switch ((scope ?? "selection").ToLowerInvariant())
            {
                case "document": return doc.Content;
                case "paragraph": return app.Selection.Paragraphs.First.Range;
                default:
                    var r = app.Selection.Range;
                    return r.Start == r.End ? app.Selection.Paragraphs.First.Range : r;
            }
        }
    }
}
