using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using MasterBW.Core.Academic;
using MasterBW.Core.Ai;
using MasterBW.Core.Indonesia;
using MasterBW.Office.Common;
using Word = Microsoft.Office.Interop.Word;

namespace MasterBW.Office.ForWord
{
    /// <summary>Merender blok dokumen (templat/AI) ke Word dengan style bawaan agar daftar isi otomatis berfungsi.</summary>
    internal sealed class WordRenderer
    {
        private readonly Word.Application _app;
        private readonly Word.Document _doc;
        private readonly ExecutionSession _session;
        public int Pos;

        public WordRenderer(Word.Application app, Word.Document doc, int pos, ExecutionSession session)
        {
            _app = app; _doc = doc; Pos = pos; _session = session;
        }

        private static readonly object Normal = Word.WdBuiltinStyle.wdStyleNormal;
        private static readonly object TitleStyle = Word.WdBuiltinStyle.wdStyleTitle;

        public async Task RenderAsync(IList<DocBlock> blocks, bool aiFill, string context)
        {
            int aiIndex = 0, aiTotal = blocks.Count(b => b.Type == "ai");
            foreach (var b in blocks)
            {
                _session.Ct.ThrowIfCancellationRequested();
                switch (b.Type)
                {
                    case "title":
                        WordHelpers.InsertParagraph(_doc, ref Pos, b.Text, TitleStyle, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                        break;
                    case "heading":
                        object h = WordHelpers.HeadingStyle(b.Level);
                        WordHelpers.InsertParagraph(_doc, ref Pos, b.Text, h);
                        break;
                    case "center":
                        NoIndent(WordHelpers.InsertParagraph(_doc, ref Pos, b.Text, Normal, Word.WdParagraphAlignment.wdAlignParagraphCenter, b.Bold));
                        break;
                    case "right":
                        NoIndent(WordHelpers.InsertParagraph(_doc, ref Pos, b.Text, Normal, Word.WdParagraphAlignment.wdAlignParagraphRight));
                        break;
                    case "bullets":
                    case "numbered":
                        object ls = b.Type == "bullets" ? Word.WdBuiltinStyle.wdStyleListBullet : Word.WdBuiltinStyle.wdStyleListNumber;
                        foreach (var item in b.Items ?? new List<string>()) WordHelpers.InsertParagraph(_doc, ref Pos, item, ls);
                        break;
                    case "table":
                        InsertTable(b.Rows, true, "grid", null);
                        break;
                    case "keyvalue":
                        InsertKeyValue(b.Rows);
                        break;
                    case "signature":
                        InsertSignature(b.Data ?? new Dictionary<string, string>());
                        break;
                    case "letterhead":
                        InsertLetterhead(_session.Settings.Institution, "", _session.Settings.InstitutionAddress, "", _session.Settings.LogoPath);
                        break;
                    case "toc":
                        InsertToc("DAFTAR ISI", 3);
                        break;
                    case "page_break":
                        _doc.Range(Pos, Pos).InsertBreak(Word.WdBreakType.wdPageBreak);
                        Pos = _doc.Range(Pos, Pos).End + 1;
                        Pos = Math.Min(Pos, _doc.Content.End - 1);
                        break;
                    case "spacer":
                        WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal);
                        break;
                    case "cover":
                        InsertCover(b.Data ?? new Dictionary<string, string>());
                        break;
                    case "ai":
                        aiIndex++;
                        if (aiFill)
                        {
                            _session.Report($"AI menulis bagian {aiIndex}/{aiTotal}…");
                            var text = await _session.Ai.CompleteAsync(Prompts.BuildContent(b.Text, "sedang", "formal", context), _session.Mode, 900, _session.Ct);
                            WordHelpers.InsertMarkdown(_doc, ref Pos, text);
                        }
                        else
                        {
                            var r = WordHelpers.InsertParagraph(_doc, ref Pos, "[Isi: " + b.Text + "]", Normal);
                            r.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                        }
                        break;
                    default:
                        WordHelpers.InsertMarkdown(_doc, ref Pos, b.Text);
                        break;
                }
            }
        }

        private static void NoIndent(Word.Range r)
        {
            r.ParagraphFormat.FirstLineIndent = 0;
            r.ParagraphFormat.LeftIndent = 0;
        }

        public Word.Table InsertTable(List<List<string>> rows, bool header, string style, string caption)
        {
            if (rows == null || rows.Count == 0) return null;
            int cols = rows.Max(r => r.Count);
            if (!string.IsNullOrWhiteSpace(caption))
            {
                // Konvensi akademik Indonesia: judul tabel di ATAS tabel.
                var capPara = WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                EnsureCaptionLabel("Tabel");
                capPara.InsertCaption("Tabel", " " + caption, "", Word.WdCaptionPosition.wdCaptionPositionAbove, false);
                Pos = capPara.Paragraphs.Last.Range.End;
            }
            var anchor = WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal);
            var table = _doc.Tables.Add(anchor, rows.Count, cols);
            table.Range.ParagraphFormat.FirstLineIndent = 0;
            table.Range.ParagraphFormat.LeftIndent = 0;
            table.Range.ParagraphFormat.SpaceAfter = 0;
            table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            for (int i = 0; i < rows.Count; i++)
                for (int j = 0; j < cols; j++)
                    table.Cell(i + 1, j + 1).Range.Text = j < rows[i].Count ? rows[i][j] ?? "" : "";
            WordTableStyler.Apply(table, style, header, null);
            Pos = table.Range.End;
            // setelah tabel pastikan ada paragraf untuk lanjut menulis
            var after = _doc.Range(Pos, Pos);
            if (Pos >= _doc.Content.End - 1) { after.InsertParagraphAfter(); }
            Pos = Math.Min(table.Range.End + 1, _doc.Content.End - 1);
            return table;
        }

        private void EnsureCaptionLabel(string label)
        {
            try { var _ = _app.CaptionLabels[label]; }
            catch { _app.CaptionLabels.Add(label); }
        }

        private void InsertKeyValue(List<List<string>> rows)
        {
            if (rows == null || rows.Count == 0) return;
            var anchor = WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal);
            var t = _doc.Tables.Add(anchor, rows.Count, 2);
            t.Borders.Enable = 0;
            t.Range.ParagraphFormat.FirstLineIndent = 0;
            t.Range.ParagraphFormat.SpaceAfter = 0;
            t.Range.ParagraphFormat.SpaceBefore = 0;
            t.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            t.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            t.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
            t.Columns[1].PreferredWidth = WordHelpers.Cm(4.2);
            for (int i = 0; i < rows.Count; i++)
            {
                t.Cell(i + 1, 1).Range.Text = rows[i].Count > 0 ? rows[i][0] : "";
                t.Cell(i + 1, 2).Range.Text = rows[i].Count > 1 ? rows[i][1] : "";
            }
            t.AllowAutoFit = true;
            Pos = Math.Min(t.Range.End + 1, _doc.Content.End - 1);
        }

        public void InsertSignature(IDictionary<string, string> d)
        {
            string G(string k, string def = "") => d.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v) ? v : def;
            var settings = _session.Settings;
            string city = G("city", settings.City), date = G("date", IndoDate.Long(DateTime.Now));
            string place = d.ContainsKey("raw_place") ? city + "\n" + date : (string.IsNullOrWhiteSpace(city) ? "[Kota]" : city) + ", " + date;
            string title = G("title", settings.UserTitle), name = G("name", settings.UserName);
            string idLabel = d.ContainsKey("id_label") ? d["id_label"] : "NIP";
            string idNum = G("id_number", settings.UserNip);
            string align = G("align", "right");

            List<string> Block(string ttl, string nm, bool withPlace)
            {
                var lines = new List<string>();
                if (withPlace) lines.AddRange(place.Split('\n'));
                lines.Add(ttl);
                lines.AddRange(new[] { "", "", "" });
                lines.Add("**" + (string.IsNullOrWhiteSpace(nm) ? "[Nama]" : nm) + "**");
                if (!string.IsNullOrWhiteSpace(idLabel)) lines.Add(idLabel + ". " + idNum);
                return lines;
            }

            WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal);
            if (align == "two_columns")
            {
                var left = Block(G("second_title", "Mengetahui,"), G("second_name"), false);
                var right = Block(title, name, true);
                left.Insert(0, "");
                var anchor = WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal);
                var t = _doc.Tables.Add(anchor, 1, 2);
                t.Borders.Enable = 0;
                t.Range.ParagraphFormat.FirstLineIndent = 0;
                t.Range.ParagraphFormat.SpaceAfter = 0;
                t.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                t.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                FillCell(t.Cell(1, 1).Range, left);
                FillCell(t.Cell(1, 2).Range, right);
                Pos = Math.Min(t.Range.End + 1, _doc.Content.End - 1);
                return;
            }
            float indent = align == "left" ? 0 : WordHelpers.Cm(align == "center" ? 5 : 9.5);
            foreach (var line in Block(title, name, true))
            {
                var r = WordHelpers.InsertParagraph(_doc, ref Pos, line, Normal);
                r.ParagraphFormat.LeftIndent = indent;
                r.ParagraphFormat.FirstLineIndent = 0;
                r.ParagraphFormat.SpaceAfter = 0;
                r.ParagraphFormat.SpaceBefore = 0;
                r.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                if (line.StartsWith("**")) r.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            }
        }

        private void FillCell(Word.Range cell, List<string> lines)
        {
            cell.Text = string.Join("\r", lines.Select(l => l.Replace("**", "")));
            var nameLine = lines.FindIndex(l => l.StartsWith("**"));
            if (nameLine >= 0 && nameLine < cell.Paragraphs.Count)
            {
                var p = cell.Paragraphs[nameLine + 1].Range;
                p.Font.Bold = 1;
                p.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            }
        }

        public void InsertLetterhead(string institution, string subtitle, string address, string contact, string logoPath)
        {
            institution = string.IsNullOrWhiteSpace(institution) ? "[NAMA INSTANSI]" : institution;
            var lines = new List<(string text, float size, bool bold)>();
            if (!string.IsNullOrWhiteSpace(subtitle)) lines.Add((subtitle.ToUpperInvariant(), 12, true));
            lines.Add((institution.ToUpperInvariant(), 16, true));
            lines.Add((string.IsNullOrWhiteSpace(address) ? "[Alamat lengkap instansi]" : address, 10, false));
            if (!string.IsNullOrWhiteSpace(contact)) lines.Add((contact, 10, false));

            bool hasLogo = !string.IsNullOrWhiteSpace(logoPath) && File.Exists(logoPath);
            var anchor = WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal);
            var t = _doc.Tables.Add(anchor, 1, hasLogo ? 2 : 1);
            t.Borders.Enable = 0;
            t.Range.ParagraphFormat.FirstLineIndent = 0;
            t.Range.ParagraphFormat.SpaceAfter = 0;
            t.Range.ParagraphFormat.SpaceBefore = 0;
            t.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            var textCell = t.Cell(1, hasLogo ? 2 : 1);
            if (hasLogo)
            {
                t.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                t.Columns[1].PreferredWidth = WordHelpers.Cm(2.8);
                var pic = _doc.InlineShapes.AddPicture(logoPath, false, true, t.Cell(1, 1).Range);
                WordHelpers.ResizeKeepRatio(pic, height: WordHelpers.Cm(2.4));
            }
            textCell.Range.Text = string.Join("\r", lines.Select(l => l.text));
            textCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            for (int i = 0; i < lines.Count; i++)
            {
                var p = textCell.Range.Paragraphs[i + 1].Range;
                p.Font.Size = lines[i].size;
                p.Font.Bold = lines[i].bold ? 1 : 0;
            }
            Pos = Math.Min(t.Range.End + 1, _doc.Content.End - 1);
            // garis kop tebal-tipis
            var line = WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal);
            var border = _doc.Range(line.Start, line.End + 1).ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom];
            border.LineStyle = Word.WdLineStyle.wdLineStyleThickThinLargeGap;
            border.LineWidth = Word.WdLineWidth.wdLineWidth300pt;
        }

        public void InsertToc(string title, int levels)
        {
            if (!string.IsNullOrWhiteSpace(title))
            {
                var t = WordHelpers.InsertParagraph(_doc, ref Pos, title, Normal, Word.WdParagraphAlignment.wdAlignParagraphCenter, true);
                NoIndent(t);
            }
            var anchor = WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal);
            NoIndent(anchor);
            var toc = _doc.TablesOfContents.Add(anchor, true, 1, Math.Max(1, Math.Min(levels, 9)));
            Pos = Math.Min(toc.Range.End + 1, _doc.Content.End - 1);
        }

        public void InsertCover(IDictionary<string, string> d)
        {
            string G(string k) => d.TryGetValue(k, out var v) ? v ?? "" : "";
            void C(string text, float size, bool bold, int spaceAfter = 6)
            {
                foreach (var line in (text ?? "").Split('\n'))
                {
                    var r = WordHelpers.InsertParagraph(_doc, ref Pos, line, Normal, Word.WdParagraphAlignment.wdAlignParagraphCenter, bold);
                    NoIndent(r);
                    r.Font.Size = size;
                    r.ParagraphFormat.SpaceAfter = spaceAfter;
                    r.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
            }
            C(G("kind").ToUpperInvariant(), 14, true, 12);
            C(G("title").ToUpperInvariant(), 14, true, 18);
            if (G("subtitle").Length > 0) C(G("subtitle"), 12, false, 24);
            var logo = G("logo_path");
            if (string.IsNullOrWhiteSpace(logo)) logo = _session.Settings.LogoPath;
            if (!string.IsNullOrWhiteSpace(logo) && File.Exists(logo))
            {
                var p = WordHelpers.InsertParagraph(_doc, ref Pos, "", Normal, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                NoIndent(p);
                var pic = _doc.InlineShapes.AddPicture(logo, false, true, p);
                WordHelpers.ResizeKeepRatio(pic, height: WordHelpers.Cm(4.5));
                Pos = p.Paragraphs.Last.Range.End;
            }
            else C("[LOGO]", 12, false, 24);
            C("Disusun oleh:", 12, false, 0);
            var authors = G("author").Split(',').Select(a => a.Trim()).Where(a => a.Length > 0).ToList();
            var ids = G("id_number").Split(',').Select(a => a.Trim()).ToList();
            if (authors.Count == 0) authors.Add("[Nama]");
            for (int i = 0; i < authors.Count; i++) C(authors[i] + (i < ids.Count && ids[i].Length > 0 ? " (" + ids[i] + ")" : ""), 12, true, 0);
            C("", 12, false, 24);
            if (G("program").Length > 0) C(G("program").ToUpperInvariant(), 12, true, 0);
            C(G("institution").ToUpperInvariant(), 14, true, 0);
            C(G("city").ToUpperInvariant(), 12, true, 0);
            C(G("year"), 12, true, 0);
            _doc.Range(Pos, Pos).InsertBreak(Word.WdBreakType.wdPageBreak);
            Pos = Math.Min(Pos + 1, _doc.Content.End - 1);
        }
    }

    internal static class WordTableStyler
    {
        public static void Apply(Word.Table t, string style, bool header, double? fontSize)
        {
            style = (style ?? "grid").ToLowerInvariant();
            t.Borders.Enable = 0;
            if (fontSize.HasValue) t.Range.Font.Size = (float)fontSize.Value;
            switch (style)
            {
                case "akademik":
                    // gaya APA: garis atas, bawah header, dan bawah tabel saja
                    t.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    if (header) t.Rows[1].Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    break;
                case "modern":
                    t.Borders[Word.WdBorderType.wdBorderHorizontal].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    t.Borders[Word.WdBorderType.wdBorderHorizontal].Color = Word.WdColor.wdColorGray25;
                    t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    if (header)
                    {
                        t.Rows[1].Shading.BackgroundPatternColor = (Word.WdColor)(0x79 * 65536 + 0x4E * 256 + 0x1F);
                        t.Rows[1].Range.Font.Color = Word.WdColor.wdColorWhite;
                    }
                    break;
                case "polos":
                    break;
                default:
                    t.Borders.Enable = 1;
                    if (header) t.Rows[1].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                    break;
            }
            if (header)
            {
                t.Rows[1].Range.Font.Bold = 1;
                t.Rows[1].HeadingFormat = -1;
                t.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            t.Rows.AllowBreakAcrossPages = 0;
        }
    }
}
