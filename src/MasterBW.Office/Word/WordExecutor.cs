using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MasterBW.Core.Academic;
using MasterBW.Core.Ai;
using MasterBW.Core.Indonesia;
using MasterBW.Core.Tools;
using MasterBW.Core.Util;
using MasterBW.Office.Common;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace MasterBW.Office.ForWord
{
    /// <summary>Menjalankan aksi dari katalog WordTools terhadap dokumen aktif.</summary>
    public sealed class WordExecutor : IActionExecutor
    {
        private readonly Word.Application _app;
        private bool _recording;
        private bool _hasUndo;

        public WordExecutor(object application) { _app = (Word.Application)application; }

        public OfficeApp App => OfficeApp.Word;
        private Word.Document Doc => _app.Documents.Count == 0 ? _app.Documents.Add() : _app.ActiveDocument;
        private static readonly object Normal = Word.WdBuiltinStyle.wdStyleNormal;

        // ------------------------------------------------------------------ konteks & undo
        public OfficeContext GetContext()
        {
            var ctx = new OfficeContext { App = OfficeApp.Word };
            try
            {
                var doc = Doc;
                ctx.DocumentName = doc.Name;
                ctx.WordCount = doc.Content.ComputeStatistics(Word.WdStatistic.wdStatisticWords);
                ctx.PageCount = doc.Content.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                var sel = _app.Selection.Range;
                if (sel.Start != sel.End) ctx.SelectionText = sel.Text;
                try { ctx.CurrentParagraphStyle = ((Word.Style)_app.Selection.Paragraphs.First.get_Style()).NameLocal; } catch { }
                int n = 0;
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    if (p.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    {
                        ctx.Headings.Add(p.Range.Text.Trim());
                        if (++n >= 20) break;
                    }
                    if (doc.Paragraphs.Count > 2000) break;
                }
            }
            catch (Exception ex) { Log.Warn("GetContext Word: " + ex.Message); }
            return ctx;
        }

        public void BeginPlan(string title)
        {
            try { _app.UndoRecord.StartCustomRecord("MASTER BW: " + (title.Length > 50 ? title.Substring(0, 50) : title)); _recording = true; }
            catch { _recording = false; }
        }

        public void EndPlan()
        {
            if (_recording) { try { _app.UndoRecord.EndCustomRecord(); } catch { } }
            _recording = false;
            _hasUndo = true;
        }

        public bool CanUndoLastPlan => _hasUndo;

        public string UndoLastPlan()
        {
            try { Doc.Undo(1); _hasUndo = false; return "Aksi terakhir MASTER BW dibatalkan (Ctrl+Z juga bisa)."; }
            catch (Exception ex) { return "Gagal membatalkan: " + ex.Message; }
        }

        // ------------------------------------------------------------------ eksekusi
        public async Task<string> ExecuteAsync(ActionStep s, ExecutionSession session)
        {
            var doc = Doc;
            switch (s.Action)
            {
                case "insert_text":
                {
                    int pos = WordHelpers.InsertionRange(_app, doc, s.Str("position")).Start;
                    WordHelpers.InsertMarkdown(doc, ref pos, s.Str("text"));
                    _app.Selection.SetRange(pos, pos);
                    return "Teks disisipkan.";
                }
                case "generate_content":
                {
                    session.Report("AI sedang menulis…");
                    var ctxText = Truncate(doc.Content.Text, 1500);
                    var text = await session.Ai.CompleteAsync(Prompts.BuildContent(s.Str("instruction"), s.Str("length"), s.Str("style"), ctxText),
                                                              session.Mode, s.Str("length") == "panjang" ? 1600 : 900, session.Ct, null);
                    int pos = WordHelpers.InsertionRange(_app, doc, s.Str("position")).Start;
                    WordHelpers.InsertMarkdown(doc, ref pos, text);
                    _app.Selection.SetRange(pos, pos);
                    return "Konten ditulis (" + CountWords(text) + " kata).";
                }
                case "transform_text": return await TransformText(s, session, doc);
                case "build_document":
                {
                    var blocks = DocBlock.ParseMany(s.Args["blocks"]);
                    var target = PrepareTarget(s.Str("position"), out int pos);
                    var r = new WordRenderer(_app, target, pos, session);
                    await r.RenderAsync(blocks, true, "");
                    return blocks.Count + " blok dokumen disusun.";
                }
                case "create_from_template": return await CreateFromTemplate(s, session);
                case "insert_cover":
                {
                    var data = new Dictionary<string, string>();
                    foreach (var k in new[] { "title", "kind", "author", "id_number", "program", "institution", "city", "year", "logo_path" }) data[k] = s.Str(k, "");
                    if (string.IsNullOrEmpty(data["kind"])) data["kind"] = "MAKALAH";
                    if (string.IsNullOrEmpty(data["year"])) data["year"] = DateTime.Now.Year.ToString();
                    if (string.IsNullOrEmpty(data["institution"])) data["institution"] = session.Settings.Institution;
                    if (string.IsNullOrEmpty(data["city"])) data["city"] = session.Settings.City;
                    if (string.IsNullOrEmpty(data["author"])) data["author"] = session.Settings.UserName;
                    var r = new WordRenderer(_app, doc, 0, session);
                    r.InsertCover(data);
                    return "Halaman sampul ditambahkan di awal dokumen.";
                }

                // ---------------- format
                case "insert_heading":
                {
                    int pos = WordHelpers.InsertionRange(_app, doc, "cursor").Start;
                    if (pos > 0 && doc.Range(pos - 1, pos).Text != "\r") { doc.Range(pos, pos).InsertParagraphAfter(); pos++; }
                    object h = WordHelpers.HeadingStyle(s.IntOr("level", 1));
                    WordHelpers.InsertParagraph(doc, ref pos, s.Str("text"), h);
                    _app.Selection.SetRange(pos, pos);
                    return "Judul ditambahkan.";
                }
                case "apply_style":
                {
                    var r = WordHelpers.ScopeRange(_app, doc, s.Str("scope", "paragraph"));
                    object style = ResolveStyle(doc, s.Str("style"));
                    r.set_Style(ref style);
                    return "Style diterapkan.";
                }
                case "format_font":
                {
                    var r = WordHelpers.ScopeRange(_app, doc, s.Str("scope", "selection"));
                    if (s.Str("name") != null) r.Font.Name = s.Str("name");
                    if (s.Num("size").HasValue) r.Font.Size = (float)s.Num("size").Value;
                    if (s.Args["bold"] != null) r.Font.Bold = s.Bool("bold") ? 1 : 0;
                    if (s.Args["italic"] != null) r.Font.Italic = s.Bool("italic") ? 1 : 0;
                    if (s.Args["underline"] != null) r.Font.Underline = s.Bool("underline") ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone;
                    var c = WordHelpers.ParseColor(s.Str("color"));
                    if (c.HasValue) r.Font.Color = (Word.WdColor)c.Value;
                    if (s.Str("highlight") != null) r.HighlightColorIndex = WordHelpers.Highlight(s.Str("highlight"));
                    if (s.Str("scope") == "document")
                    {
                        var normal = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
                        if (s.Str("name") != null) normal.Font.Name = s.Str("name");
                        if (s.Num("size").HasValue) normal.Font.Size = (float)s.Num("size").Value;
                    }
                    return "Format huruf diubah.";
                }
                case "format_paragraph":
                {
                    var r = WordHelpers.ScopeRange(_app, doc, s.Str("scope", "selection"));
                    ApplyParagraph(r.ParagraphFormat, s);
                    if (s.Str("scope") == "document") ApplyParagraph(doc.Styles[Word.WdBuiltinStyle.wdStyleNormal].ParagraphFormat, s);
                    return "Format paragraf diubah.";
                }
                case "page_setup":
                    ApplyPageSetup(doc, s.Str("paper"), s.Str("orientation"), s.Num("margin_top_cm"), s.Num("margin_bottom_cm"), s.Num("margin_left_cm"), s.Num("margin_right_cm"));
                    return "Pengaturan halaman diterapkan.";
                case "apply_academic_format":
                {
                    var preset = FormatPreset.All.TryGetValue(s.Str("preset", "Skripsi"), out var p) ? p : FormatPreset.All["Skripsi"];
                    ApplyPreset(doc, preset, s.Str("font"), s.Num("size"), s.Num("line_spacing"));
                    return "Format " + preset.Name + " diterapkan: " + preset.Font + " " + preset.Size + " pt, spasi " + preset.LineSpacing + ", margin " + preset.MarginTop + "-" + preset.MarginLeft + "-" + preset.MarginBottom + "-" + preset.MarginRight + " cm.";
                }
                case "change_case":
                {
                    var r = WordHelpers.ScopeRange(_app, doc, s.Str("scope", "selection"));
                    switch (s.Str("case"))
                    {
                        case "upper": r.Case = Word.WdCharacterCase.wdUpperCase; break;
                        case "lower": r.Case = Word.WdCharacterCase.wdLowerCase; break;
                        case "title": r.Case = Word.WdCharacterCase.wdTitleWord; break;
                        default: r.Case = Word.WdCharacterCase.wdTitleSentence; break;
                    }
                    return "Kapitalisasi diubah.";
                }
                case "set_columns":
                {
                    var r = s.Str("scope") == "selection" ? _app.Selection.Range : doc.Content;
                    r.PageSetup.TextColumns.SetCount(Math.Max(1, s.IntOr("count", 2)));
                    if (s.Num("spacing_cm").HasValue) r.PageSetup.TextColumns.Spacing = WordHelpers.Cm(s.Num("spacing_cm").Value);
                    return "Teks dibagi menjadi " + s.IntOr("count", 2) + " kolom.";
                }
                case "clean_document": return CleanDocument(doc, s);

                // ---------------- struktur
                case "heading_numbering": return HeadingNumbering(doc, s.Str("scheme", "bab_romawi"));
                case "insert_toc":
                {
                    int pos = WordHelpers.InsertionRange(_app, doc, "cursor").Start;
                    new WordRenderer(_app, doc, pos, session).InsertToc(s.Str("title", "DAFTAR ISI"), s.IntOr("levels", 3));
                    return "Daftar isi disisipkan. Gunakan 'Perbarui field' setelah dokumen berubah.";
                }
                case "insert_table_of_figures":
                {
                    var label = s.Str("label", "Gambar");
                    EnsureCaptionLabel(label);
                    int pos = WordHelpers.InsertionRange(_app, doc, "cursor").Start;
                    var title = s.Str("title", "DAFTAR " + label.ToUpperInvariant());
                    WordHelpers.InsertParagraph(doc, ref pos, title, Normal, Word.WdParagraphAlignment.wdAlignParagraphCenter, true);
                    var anchor = WordHelpers.InsertParagraph(doc, ref pos, "", Normal);
                    doc.TablesOfFigures.Add(anchor, label);
                    return "Daftar " + label.ToLowerInvariant() + " disisipkan.";
                }
                case "update_fields":
                    foreach (Word.TableOfContents t in doc.TablesOfContents) t.Update();
                    foreach (Word.TableOfFigures t in doc.TablesOfFigures) t.Update();
                    doc.Fields.Update();
                    return "Daftar isi, caption, dan field diperbarui.";
                case "insert_page_numbers": return PageNumbers(doc, s);
                case "set_header_footer":
                    foreach (Word.Section sec in doc.Sections)
                    {
                        if (s.Args["different_first_page"] != null) sec.PageSetup.DifferentFirstPageHeaderFooter = s.Bool("different_first_page") ? -1 : 0;
                        if (s.Str("header") != null) sec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = s.Str("header");
                        if (s.Str("footer") != null) sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = s.Str("footer");
                    }
                    return "Header/footer diisi.";
                case "insert_break":
                {
                    var r = WordHelpers.InsertionRange(_app, doc, "cursor");
                    var type = s.Str("type", "page");
                    r.InsertBreak(type == "section_next" ? Word.WdBreakType.wdSectionBreakNextPage : type == "section_continuous" ? Word.WdBreakType.wdSectionBreakContinuous : type == "column" ? Word.WdBreakType.wdColumnBreak : Word.WdBreakType.wdPageBreak);
                    return "Pemisah disisipkan.";
                }
                case "insert_list": return InsertList(doc, s);

                // ---------------- tabel & objek
                case "insert_table":
                {
                    int pos = WordHelpers.InsertionRange(_app, doc, "cursor").Start;
                    var rows = s.Grid("rows");
                    var t = new WordRenderer(_app, doc, pos, session).InsertTable(rows, s.Bool("header", true), s.Str("style", "grid"), s.Str("caption"));
                    if (t != null && s.Str("autofit") == "content") t.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
                    else t?.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                    return "Tabel " + rows.Count + " baris disisipkan.";
                }
                case "format_table":
                {
                    var tables = s.Str("scope") == "all" ? doc.Tables.Cast<Word.Table>().ToList()
                               : _app.Selection.Tables.Count > 0 ? new List<Word.Table> { _app.Selection.Tables[1] } : new List<Word.Table>();
                    if (tables.Count == 0) return "Tidak ada tabel terpilih. Letakkan kursor di dalam tabel atau gunakan scope 'all'.";
                    foreach (var t in tables)
                    {
                        WordTableStyler.Apply(t, s.Str("style", "akademik"), true, s.Num("font_size"));
                        var fill = WordHelpers.ParseColor(s.Str("header_fill"));
                        if (fill.HasValue) t.Rows[1].Shading.BackgroundPatternColor = (Word.WdColor)fill.Value;
                        if (s.Str("autofit") == "content") t.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
                        if (s.Str("autofit") == "window") t.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                        if (s.Bool("repeat_header", true)) t.Rows[1].HeadingFormat = -1;
                    }
                    return tables.Count + " tabel diformat.";
                }
                case "convert_text_to_table":
                {
                    var r = _app.Selection.Range;
                    if (r.Start == r.End) return "Pilih dulu teks yang akan dijadikan tabel.";
                    var sep = s.Str("separator", "tab");
                    object sepObj = sep == "comma" ? (object)"," : sep == "semicolon" ? (object)";" : sep == "pipe" ? (object)"|" : Word.WdTableFieldSeparator.wdSeparateByTabs;
                    var t = r.ConvertToTable(sepObj);
                    WordTableStyler.Apply(t, "grid", true, null);
                    return "Teks diubah menjadi tabel " + t.Rows.Count + " × " + t.Columns.Count + ".";
                }
                case "insert_caption":
                {
                    var label = s.Str("label", "Gambar");
                    EnsureCaptionLabel(label);
                    var pos = s.Str("position", label == "Tabel" ? "above" : "below") == "above" ? Word.WdCaptionPosition.wdCaptionPositionAbove : Word.WdCaptionPosition.wdCaptionPositionBelow;
                    _app.Selection.Range.InsertCaption(label, " " + s.Str("text"), "", pos, false);
                    return "Caption " + label + " ditambahkan (nomor otomatis).";
                }
                case "insert_image":
                {
                    var path = s.Str("path");
                    if (!File.Exists(path)) return "File gambar tidak ditemukan: " + path;
                    var r = WordHelpers.InsertionRange(_app, doc, "cursor");
                    var pic = doc.InlineShapes.AddPicture(path, false, true, r);
                    if (s.Num("width_cm").HasValue) WordHelpers.ResizeKeepRatio(pic, width: WordHelpers.Cm(s.Num("width_cm").Value));
                    if (!string.IsNullOrWhiteSpace(s.Str("caption")))
                    {
                        EnsureCaptionLabel("Gambar");
                        pic.Range.InsertCaption("Gambar", " " + s.Str("caption"), "", Word.WdCaptionPosition.wdCaptionPositionBelow, false);
                    }
                    return "Gambar disisipkan.";
                }
                case "insert_equation":
                {
                    var r = WordHelpers.InsertionRange(_app, doc, "cursor");
                    r.Text = s.Str("linear");
                    r.OMaths.Add(r);
                    r.OMaths[1].BuildUp();
                    return "Persamaan disisipkan.";
                }
                case "insert_footnote":
                    doc.Footnotes.Add(_app.Selection.Range, Type.Missing, s.Str("text"));
                    return "Catatan kaki ditambahkan.";

                // ---------------- akademik
                case "insert_citation":
                {
                    var reference = ToReference(s.Args["reference"]);
                    var style = s.Str("style", session.Settings.CitationStyle);
                    var store = ReferenceStore.Load();
                    int number = store.AddOrGet(reference);
                    store.Save();
                    var text = CitationFormatter.InText(reference, style, s.Str("page"), number);
                    int pos = _app.Selection.Range.End;
                    InsertInline(doc, ref pos, " " + text);
                    _app.Selection.SetRange(pos, pos);
                    return "Sitasi " + text + " disisipkan dan pustaka disimpan untuk daftar pustaka.";
                }
                case "insert_bibliography": return InsertBibliography(doc, s, session);
                case "format_bibliography":
                {
                    var r = _app.Selection.Range;
                    if (r.Start == r.End) return "Pilih dulu paragraf daftar pustaka.";
                    r.Sort(false, Word.WdSortFieldType.wdSortFieldAlphanumeric);
                    r.ParagraphFormat.LeftIndent = WordHelpers.Cm(1.27);
                    r.ParagraphFormat.FirstLineIndent = -WordHelpers.Cm(1.27);
                    r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    return "Daftar pustaka diurutkan abjad dan diberi indentasi gantung 1,27 cm.";
                }
                case "check_spelling": return CheckSpelling(doc, s);

                // ---------------- kantor
                case "insert_letterhead":
                {
                    var r = new WordRenderer(_app, doc, 0, session);
                    r.InsertLetterhead(s.Str("institution", session.Settings.Institution), s.Str("subtitle"), s.Str("address", session.Settings.InstitutionAddress), s.Str("contact"), s.Str("logo_path", session.Settings.LogoPath));
                    return "Kop surat ditambahkan di awal dokumen.";
                }
                case "insert_signature":
                {
                    int pos = WordHelpers.InsertionRange(_app, doc, "cursor").Start;
                    var d = new Dictionary<string, string>
                    {
                        ["city"] = s.Str("city", ""), ["date"] = s.Str("date", ""), ["title"] = s.Str("position_title", ""), ["name"] = s.Str("name", ""),
                        ["id_label"] = s.Str("id_label", "NIP"), ["id_number"] = s.Str("id_number", ""), ["align"] = s.Str("align", "right"),
                        ["second_title"] = s.Str("second_title", ""), ["second_name"] = s.Str("second_name", ""),
                    };
                    new WordRenderer(_app, doc, pos, session).InsertSignature(d);
                    return "Blok tanda tangan disisipkan.";
                }
                case "insert_date":
                {
                    var now = DateTime.Now;
                    var f = s.Str("format", "panjang");
                    var text = f == "pendek" ? IndoDate.Short(now) : f == "hari_tanggal" ? IndoDate.WithDay(now) : IndoDate.Long(now);
                    _app.Selection.Range.InsertAfter(text);
                    return "Tanggal disisipkan: " + text;
                }
                case "insert_watermark": return Watermark(doc, s.Str("text", "DRAFT"));
                case "mail_merge": return MailMerge(doc, s);
                case "find_replace":
                {
                    var find = doc.Content.Find;
                    find.ClearFormatting();
                    find.Replacement.ClearFormatting();
                    bool ok = find.Execute(s.Str("find"), s.Bool("match_case"), s.Bool("whole_word"), s.Bool("wildcards"), false, false, true,
                                           Word.WdFindWrap.wdFindContinue, false, s.Str("replace") ?? "", Word.WdReplace.wdReplaceAll);
                    return ok ? "Penggantian selesai." : "Teks '" + s.Str("find") + "' tidak ditemukan.";
                }
                case "highlight_text":
                {
                    _app.Options.DefaultHighlightColorIndex = WordHelpers.Highlight(s.Str("color"));
                    var find = doc.Content.Find;
                    find.ClearFormatting();
                    find.Replacement.ClearFormatting();
                    find.Replacement.Highlight = 1;
                    find.Execute(s.Str("find"), false, false, false, false, false, true, Word.WdFindWrap.wdFindContinue, true, "^&", Word.WdReplace.wdReplaceAll);
                    return "Semua '" + s.Str("find") + "' disorot.";
                }
                case "add_comment":
                    doc.Comments.Add(_app.Selection.Range, s.Str("text"));
                    return "Komentar ditambahkan.";
                case "track_changes":
                    doc.TrackRevisions = s.Bool("on", true);
                    return "Lacak perubahan " + (doc.TrackRevisions ? "aktif." : "nonaktif.");
                case "insert_hyperlink":
                    doc.Hyperlinks.Add(_app.Selection.Range, s.Str("url"), Type.Missing, Type.Missing, s.Str("text"));
                    return "Tautan disisipkan.";
                case "export_pdf":
                {
                    var path = s.Str("path");
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        var baseName = string.IsNullOrEmpty(doc.Path) ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Path.GetFileNameWithoutExtension(doc.Name)) : Path.Combine(doc.Path, Path.GetFileNameWithoutExtension(doc.Name));
                        path = baseName + ".pdf";
                    }
                    doc.ExportAsFixedFormat(path, Word.WdExportFormat.wdExportFormatPDF);
                    return "PDF disimpan: " + path;
                }
                case "document_stats":
                {
                    int words = doc.Content.ComputeStatistics(Word.WdStatistic.wdStatisticWords);
                    var text = "Statistik dokumen:\n" +
                               "• Halaman: " + doc.Content.ComputeStatistics(Word.WdStatistic.wdStatisticPages) + "\n" +
                               "• Kata: " + words + "\n" +
                               "• Karakter (dengan spasi): " + doc.Content.ComputeStatistics(Word.WdStatistic.wdStatisticCharactersWithSpaces) + "\n" +
                               "• Paragraf: " + doc.Content.ComputeStatistics(Word.WdStatistic.wdStatisticParagraphs) + "\n" +
                               "• Tabel: " + doc.Tables.Count + ", gambar: " + doc.InlineShapes.Count + "\n" +
                               "• Perkiraan waktu baca: " + Math.Max(1, words / 200) + " menit";
                    session.ShowText("Statistik dokumen", text);
                    return text;
                }
                default:
                    return "Aksi '" + s.Action + "' belum didukung di Word.";
            }
        }

        // ================================================================== implementasi
        private Word.Document PrepareTarget(string position, out int pos)
        {
            if (position == "new_document" || Doc.Content.Text.Trim().Length == 0 && position != "cursor")
            {
                var d = Doc.Content.Text.Trim().Length == 0 ? Doc : _app.Documents.Add();
                d.Activate();
                pos = 0;
                return d;
            }
            var doc = Doc;
            pos = WordHelpers.InsertionRange(_app, doc, position ?? "cursor").Start;
            return doc;
        }

        private async Task<string> CreateFromTemplate(ActionStep s, ExecutionSession session)
        {
            var tpl = DocumentTemplates.Find(s.Str("template"));
            if (tpl == null) return "Templat '" + s.Str("template") + "' tidak ditemukan. Tersedia: " + string.Join(", ", DocumentTemplates.All.Select(t => t.Id));
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["instansi"] = session.Settings.Institution, ["alamat_instansi"] = session.Settings.InstitutionAddress, ["kota"] = session.Settings.City,
                ["jabatan"] = session.Settings.UserTitle, ["nama_penandatangan"] = session.Settings.UserName, ["nip"] = session.Settings.UserNip,
                ["penulis"] = session.Settings.UserName, ["institusi"] = session.Settings.Institution,
            };
            if (s.Args["fields"] is JObject f) foreach (var p in f.Properties()) values[p.Name] = p.Value.ToString();
            var blocks = tpl.Render(values);
            var doc = Doc.Content.Text.Trim().Length == 0 ? Doc : _app.Documents.Add();
            doc.Activate();
            if (FormatPreset.All.TryGetValue(tpl.Preset, out var preset)) ApplyPreset(doc, preset, null, null, null, numbering: false);
            var summary = string.Join("; ", values.Where(v => !string.IsNullOrWhiteSpace(v.Value)).Select(v => v.Key + ": " + v.Value));
            var r = new WordRenderer(_app, doc, 0, session);
            await r.RenderAsync(blocks, s.Bool("ai_fill", true), "Templat: " + tpl.Name + ". Data: " + summary);
            foreach (Word.TableOfContents t in doc.TablesOfContents) t.Update();
            return tpl.Name + " dibuat di dokumen baru. Periksa bagian bertanda [ ] dan sesuaikan isinya.";
        }

        private async Task<string> TransformText(ActionStep s, ExecutionSession session, Word.Document doc)
        {
            var scope = s.Str("scope", "selection");
            Word.Range r = scope == "document" ? doc.Content : _app.Selection.Range;
            if (r.Start == r.End) r = _app.Selection.Paragraphs.First.Range;
            var source = r.Text ?? "";
            if (source.Trim().Length == 0) return "Tidak ada teks untuk diolah. Pilih teksnya dulu.";
            var task = s.Str("task", "perbaiki");
            session.Report("AI mengolah teks (" + task + ")…");
            var result = await session.Ai.TransformAsync(task, source.Trim(), session.Mode, session.Ct, session.Report);
            var output = s.Str("output", "replace");
            if (task == "judul" || task == "pertanyaan") output = output == "replace" ? "below" : output;
            switch (output)
            {
                case "comment":
                    doc.Comments.Add(r, result);
                    return "Hasil ditambahkan sebagai komentar.";
                case "new_document":
                {
                    var nd = _app.Documents.Add();
                    int p0 = 0;
                    WordHelpers.InsertMarkdown(nd, ref p0, result);
                    return "Hasil ditulis di dokumen baru.";
                }
                case "below":
                {
                    int pos = r.End;
                    if (doc.Range(pos - 1, pos).Text != "\r") { doc.Range(pos, pos).InsertParagraphAfter(); pos++; }
                    WordHelpers.InsertMarkdown(doc, ref pos, result);
                    return "Hasil ditulis di bawah teks asli.";
                }
                default:
                {
                    bool endsWithPara = source.EndsWith("\r");
                    int start = r.Start;
                    var target = endsWithPara ? doc.Range(r.Start, r.End - 1) : r;
                    target.Text = "";
                    int pos = start;
                    if (result.Contains("\n") || result.StartsWith("#") || result.StartsWith("- "))
                    {
                        WordHelpers.InsertMarkdown(doc, ref pos, result);
                        // hapus paragraf kosong sisa
                        var leftover = doc.Range(pos, Math.Min(pos + 1, doc.Content.End));
                        if (leftover.Text == "\r" && pos < doc.Content.End - 1) leftover.Delete();
                    }
                    else doc.Range(start, start).InsertAfter(result);
                    return "Teks diganti dengan hasil " + task + " (Ctrl+Z untuk membatalkan).";
                }
            }
        }

        private static void ApplyParagraph(Word.ParagraphFormat pf, ActionStep s)
        {
            switch (s.Str("alignment"))
            {
                case "left": pf.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                case "center": pf.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                case "right": pf.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                case "justify": pf.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify; break;
            }
            var ls = s.Num("line_spacing");
            if (ls.HasValue) SetLineSpacing(pf, ls.Value);
            if (s.Num("space_before").HasValue) pf.SpaceBefore = (float)s.Num("space_before").Value;
            if (s.Num("space_after").HasValue) pf.SpaceAfter = (float)s.Num("space_after").Value;
            if (s.Num("first_line_indent_cm").HasValue) pf.FirstLineIndent = WordHelpers.Cm(s.Num("first_line_indent_cm").Value);
            if (s.Num("left_indent_cm").HasValue) pf.LeftIndent = WordHelpers.Cm(s.Num("left_indent_cm").Value);
            if (s.Num("hanging_indent_cm").HasValue)
            {
                pf.LeftIndent = WordHelpers.Cm(s.Num("hanging_indent_cm").Value);
                pf.FirstLineIndent = -WordHelpers.Cm(s.Num("hanging_indent_cm").Value);
            }
        }

        private static void SetLineSpacing(Word.ParagraphFormat pf, double v)
        {
            if (Math.Abs(v - 1) < 0.01) pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            else if (Math.Abs(v - 1.5) < 0.01) pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
            else if (Math.Abs(v - 2) < 0.01) pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
            else { pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple; pf.LineSpacing = (float)(12 * v); }
        }

        private static void ApplyPageSetup(Word.Document doc, string paper, string orientation, double? top, double? bottom, double? left, double? right)
        {
            var ps = doc.PageSetup;
            switch ((paper ?? "").ToUpperInvariant())
            {
                case "A4": ps.PaperSize = Word.WdPaperSize.wdPaperA4; break;
                case "LETTER": ps.PaperSize = Word.WdPaperSize.wdPaperLetter; break;
                case "LEGAL": ps.PaperSize = Word.WdPaperSize.wdPaperLegal; break;
                case "A5": ps.PaperSize = Word.WdPaperSize.wdPaperA5; break;
                case "F4": ps.PageWidth = WordHelpers.Cm(21.5); ps.PageHeight = WordHelpers.Cm(33); break;
            }
            if (orientation == "landscape") ps.Orientation = Word.WdOrientation.wdOrientLandscape;
            if (orientation == "portrait") ps.Orientation = Word.WdOrientation.wdOrientPortrait;
            if (top.HasValue) ps.TopMargin = WordHelpers.Cm(top.Value);
            if (bottom.HasValue) ps.BottomMargin = WordHelpers.Cm(bottom.Value);
            if (left.HasValue) ps.LeftMargin = WordHelpers.Cm(left.Value);
            if (right.HasValue) ps.RightMargin = WordHelpers.Cm(right.Value);
        }

        private void ApplyPreset(Word.Document doc, FormatPreset p, string font, double? size, double? spacing, bool numbering = false)
        {
            font = string.IsNullOrWhiteSpace(font) ? p.Font : font;
            float sz = (float)(size ?? p.Size);
            double sp = spacing ?? p.LineSpacing;
            ApplyPageSetup(doc, p.Paper, "portrait", p.MarginTop, p.MarginBottom, p.MarginLeft, p.MarginRight);
            var normal = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
            normal.Font.Name = font;
            normal.Font.Size = sz;
            normal.Font.Color = Word.WdColor.wdColorAutomatic;
            SetLineSpacing(normal.ParagraphFormat, sp);
            normal.ParagraphFormat.SpaceAfter = (float)p.SpaceAfter;
            normal.ParagraphFormat.SpaceBefore = 0;
            normal.ParagraphFormat.FirstLineIndent = WordHelpers.Cm(p.FirstLineIndentCm);
            normal.ParagraphFormat.Alignment = p.Justify ? Word.WdParagraphAlignment.wdAlignParagraphJustify : Word.WdParagraphAlignment.wdAlignParagraphLeft;
            for (int lvl = 1; lvl <= 3; lvl++)
            {
                var h = doc.Styles[WordHelpers.HeadingStyle(lvl)];
                h.Font.Name = font;
                h.Font.Size = sz;
                h.Font.Bold = 1;
                h.Font.Italic = 0;
                h.Font.Color = Word.WdColor.wdColorAutomatic;
                h.ParagraphFormat.FirstLineIndent = 0;
                h.ParagraphFormat.LeftIndent = 0;
                h.ParagraphFormat.SpaceBefore = lvl == 1 ? 0 : 12;
                h.ParagraphFormat.SpaceAfter = lvl == 1 ? 24 : 6;
                h.ParagraphFormat.Alignment = lvl == 1 && p.Name != "Jurnal" ? Word.WdParagraphAlignment.wdAlignParagraphCenter : Word.WdParagraphAlignment.wdAlignParagraphLeft;
                if (lvl == 1 && p.Name != "Jurnal") h.ParagraphFormat.PageBreakBefore = 0;
            }
            var title = doc.Styles[Word.WdBuiltinStyle.wdStyleTitle];
            title.Font.Name = font;
            title.Font.Size = sz + 2;
            title.Font.Bold = 1;
            title.Font.Color = Word.WdColor.wdColorAutomatic;
            title.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            foreach (var ls in new[] { Word.WdBuiltinStyle.wdStyleListBullet, Word.WdBuiltinStyle.wdStyleListNumber })
            {
                var st = doc.Styles[ls];
                st.Font.Name = font;
                st.Font.Size = sz;
                st.ParagraphFormat.FirstLineIndent = -WordHelpers.Cm(0.63);
                st.ParagraphFormat.LeftIndent = WordHelpers.Cm(0.63);
            }
            if (numbering && p.HeadingNumbering != "none") HeadingNumbering(doc, p.HeadingNumbering);
        }

        private string CleanDocument(Word.Document doc, ActionStep s)
        {
            int passes = 0;
            void ReplaceAll(string find, string repl, bool wildcards = false, int loops = 1)
            {
                for (int i = 0; i < loops; i++)
                {
                    var f = doc.Content.Find;
                    f.ClearFormatting();
                    f.Replacement.ClearFormatting();
                    if (!f.Execute(find, false, false, wildcards, false, false, true, Word.WdFindWrap.wdFindContinue, false, repl, Word.WdReplace.wdReplaceAll)) break;
                    passes++;
                }
            }
            if (s.Bool("double_spaces", true)) ReplaceAll("  ", " ", false, 8);
            if (s.Bool("manual_line_breaks")) ReplaceAll("^l", " ");
            if (s.Bool("tabs")) ReplaceAll("^t", " ");
            if (s.Bool("empty_paragraphs", true)) ReplaceAll("^p^p", "^p", false, 8);
            if (s.Bool("smart_quotes"))
            {
                ReplaceAll("“", "\""); ReplaceAll("”", "\""); ReplaceAll("‘", "'"); ReplaceAll("’", "'");
            }
            ReplaceAll(" ^p", "^p", false, 3);
            return "Dokumen dirapikan (" + passes + " jenis perbaikan diterapkan).";
        }

        private string HeadingNumbering(Word.Document doc, string scheme)
        {
            if (scheme == "none")
            {
                for (int lvl = 1; lvl <= 3; lvl++) doc.Styles[WordHelpers.HeadingStyle(lvl)].LinkToListTemplate(null);
                foreach (Word.Paragraph p in doc.Paragraphs)
                    if (p.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText) p.Range.ListFormat.RemoveNumbers();
                return "Penomoran judul dihapus.";
            }
            var lt = doc.ListTemplates.Add(true);
            for (int lvl = 1; lvl <= 3; lvl++)
            {
                var L = lt.ListLevels[lvl];
                L.TrailingCharacter = Word.WdTrailingCharacter.wdTrailingSpace;
                L.Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
                L.TextPosition = 0;
                L.NumberPosition = 0;
                L.TabPosition = 0;
                L.ResetOnHigher = lvl - 1;
                L.StartAt = 1;
                L.LinkedStyle = doc.Styles[WordHelpers.HeadingStyle(lvl)].NameLocal;
            }
            switch (scheme)
            {
                case "bab_angka":
                    Set(lt, 1, "BAB %1", Word.WdListNumberStyle.wdListNumberStyleArabic);
                    Set(lt, 2, "%1.%2", Word.WdListNumberStyle.wdListNumberStyleArabic);
                    Set(lt, 3, "%1.%2.%3", Word.WdListNumberStyle.wdListNumberStyleArabic);
                    break;
                case "desimal":
                    Set(lt, 1, "%1.", Word.WdListNumberStyle.wdListNumberStyleArabic);
                    Set(lt, 2, "%1.%2", Word.WdListNumberStyle.wdListNumberStyleArabic);
                    Set(lt, 3, "%1.%2.%3", Word.WdListNumberStyle.wdListNumberStyleArabic);
                    break;
                default: // bab_romawi: BAB I → A. → 1.
                    Set(lt, 1, "BAB %1", Word.WdListNumberStyle.wdListNumberStyleUppercaseRoman);
                    Set(lt, 2, "%2.", Word.WdListNumberStyle.wdListNumberStyleUppercaseLetter);
                    Set(lt, 3, "%3.", Word.WdListNumberStyle.wdListNumberStyleArabic);
                    break;
            }
            for (int lvl = 1; lvl <= 3; lvl++) doc.Styles[WordHelpers.HeadingStyle(lvl)].LinkToListTemplate(lt, lvl);
            return "Penomoran judul otomatis diterapkan (" + scheme + "). Jika judul Anda sudah berisi 'BAB I' manual, hapus teks nomor manualnya.";
        }

        private static void Set(Word.ListTemplate lt, int level, string format, Word.WdListNumberStyle style)
        {
            lt.ListLevels[level].NumberFormat = format;
            lt.ListLevels[level].NumberStyle = style;
        }

        private string PageNumbers(Word.Document doc, ActionStep s)
        {
            var position = s.Str("position", "bottom_center");
            var align = position.EndsWith("right") ? Word.WdPageNumberAlignment.wdAlignPageNumberRight : Word.WdPageNumberAlignment.wdAlignPageNumberCenter;
            var style = s.Str("format") == "roman_lower" ? Word.WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
                      : s.Str("format") == "roman_upper" ? Word.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman : Word.WdPageNumberStyle.wdPageNumberStyleArabic;
            var sections = s.Str("scope") == "current_section" ? new List<Word.Section> { _app.Selection.Sections[1] } : doc.Sections.Cast<Word.Section>().ToList();
            foreach (var sec in sections)
            {
                var hf = position.StartsWith("top") ? sec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary] : sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                var pn = hf.PageNumbers;
                pn.NumberStyle = style;
                if (s.Num("start_at").HasValue) { pn.RestartNumberingAtSection = true; pn.StartingNumber = s.IntOr("start_at", 1); }
                pn.Add(align, true);
            }
            return "Nomor halaman ditambahkan.";
        }

        private string InsertList(Word.Document doc, ActionStep s)
        {
            var items = s.StrList("items");
            int pos = WordHelpers.InsertionRange(_app, doc, "cursor").Start;
            int start = pos;
            var style = s.Str("style", "bullet");
            object ls = style == "bullet" ? Word.WdBuiltinStyle.wdStyleListBullet : Word.WdBuiltinStyle.wdStyleListNumber;
            foreach (var it in items) WordHelpers.InsertParagraph(doc, ref pos, it, ls);
            if (style == "letter" || style == "roman")
            {
                var lt = doc.ListTemplates.Add(false);
                lt.ListLevels[1].NumberFormat = "%1.";
                lt.ListLevels[1].NumberStyle = style == "letter" ? Word.WdListNumberStyle.wdListNumberStyleLowercaseLetter : Word.WdListNumberStyle.wdListNumberStyleLowercaseRoman;
                doc.Range(start, pos).ListFormat.ApplyListTemplate(lt, false);
            }
            return items.Count + " butir daftar disisipkan.";
        }

        private void EnsureCaptionLabel(string label)
        {
            try { var _ = _app.CaptionLabels[label]; }
            catch { _app.CaptionLabels.Add(label); }
        }

        private static object ResolveStyle(Word.Document doc, string name)
        {
            var key = (name ?? "").Trim().ToLowerInvariant().Replace(" ", "");
            switch (key)
            {
                case "heading1": case "judul1": return Word.WdBuiltinStyle.wdStyleHeading1;
                case "heading2": case "judul2": return Word.WdBuiltinStyle.wdStyleHeading2;
                case "heading3": case "judul3": return Word.WdBuiltinStyle.wdStyleHeading3;
                case "normal": return Word.WdBuiltinStyle.wdStyleNormal;
                case "title": case "judul": return Word.WdBuiltinStyle.wdStyleTitle;
                case "subtitle": case "subjudul": return Word.WdBuiltinStyle.wdStyleSubtitle;
                case "quote": case "kutipan": return Word.WdBuiltinStyle.wdStyleQuote;
                case "caption": case "keterangan": return Word.WdBuiltinStyle.wdStyleCaption;
                case "listparagraph": return Word.WdBuiltinStyle.wdStyleListParagraph;
                case "listbullet": return Word.WdBuiltinStyle.wdStyleListBullet;
                case "listnumber": return Word.WdBuiltinStyle.wdStyleListNumber;
                default: return name;
            }
        }

        private static void InsertInline(Word.Document doc, ref int pos, string marked)
        {
            // dukung *miring* (judul buku/jurnal pada sitasi)
            var parts = marked.Split('*');
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i].Length == 0) continue;
                var r = doc.Range(pos, pos);
                r.InsertAfter(parts[i]);
                r.Font.Italic = i % 2 == 1 ? 1 : 0;
                pos = r.End;
            }
        }

        private static Reference ToReference(JToken t)
        {
            if (t is JObject o)
            {
                var r = o.ToObject<Reference>() ?? new Reference();
                if (o["authors"] is JValue v && v.Type == JTokenType.String) r.Authors = ((string)v).Split(';').Select(x => x.Trim()).ToList();
                if (o["author"] != null && r.Authors.Count == 0) r.Authors = ((string)o["author"]).Split(';').Select(x => x.Trim()).ToList();
                if (!string.IsNullOrEmpty(r.Id))
                {
                    var stored = ReferenceStore.Load().Items.FirstOrDefault(x => x.Id == r.Id);
                    if (stored != null && string.IsNullOrEmpty(r.Title)) return stored;
                }
                return r;
            }
            return new Reference { Title = t?.ToString() };
        }

        private string InsertBibliography(Word.Document doc, ActionStep s, ExecutionSession session)
        {
            var style = s.Str("style", session.Settings.CitationStyle);
            var store = ReferenceStore.Load();
            if (s.Args["references"] is JArray arr)
            {
                foreach (var item in arr) store.AddOrGet(ToReference(item));
                store.Save();
            }
            if (store.Items.Count == 0) return "Belum ada pustaka. Tambahkan lewat 'Sitasi' atau sebutkan daftar pustakanya dalam perintah.";
            var entries = CitationFormatter.Bibliography(store.Items, style);
            int pos = WordHelpers.InsertionRange(_app, doc, "cursor").Start;
            var title = s.Str("title", "DAFTAR PUSTAKA");
            var t = WordHelpers.InsertParagraph(doc, ref pos, title, Normal, Word.WdParagraphAlignment.wdAlignParagraphCenter, true);
            t.ParagraphFormat.FirstLineIndent = 0;
            foreach (var e in entries)
            {
                var para = WordHelpers.InsertParagraph(doc, ref pos, "", Normal);
                int p = para.Start;
                InsertInline(doc, ref p, e);
                var pf = doc.Range(para.Start, p + 1).ParagraphFormat;
                double hang = style.StartsWith("IEEE", StringComparison.OrdinalIgnoreCase) || style.StartsWith("Vanc", StringComparison.OrdinalIgnoreCase) ? 0.9 : 1.27;
                pf.LeftIndent = WordHelpers.Cm(hang);
                pf.FirstLineIndent = -WordHelpers.Cm(hang);
                pf.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                pf.SpaceAfter = 6;
                pos = p + 1;
            }
            return entries.Count + " entri daftar pustaka (" + CitationFormatter.Normalize(style) + ") disisipkan.";
        }

        private string CheckSpelling(Word.Document doc, ActionStep s)
        {
            bool fix = s.Bool("auto_fix");
            var range = s.Str("scope") == "selection" && _app.Selection.Range.Start != _app.Selection.Range.End ? _app.Selection.Range : doc.Content;
            int found = 0, fixedCount = 0, comments = 0;
            var paragraphs = range.Paragraphs.Cast<Word.Paragraph>().ToList();
            for (int pi = paragraphs.Count - 1; pi >= 0; pi--)
            {
                var pr = paragraphs[pi].Range;
                var text = pr.Text ?? "";
                if (text.Trim().Length < 2) continue;
                var issues = EydChecker.Check(text);
                foreach (var issue in issues.OrderByDescending(i => i.Start))
                {
                    var sub = doc.Range(pr.Start + issue.Start, pr.Start + issue.Start + issue.Length);
                    if (sub.Text != issue.Original) continue;
                    found++;
                    bool safe = issue.Category == "baku" || issue.Category == "penulisan_kata" || issue.Category == "spasi" || issue.Category == "tanda_baca";
                    if (fix && safe)
                    {
                        sub.Text = issue.Suggestion;
                        fixedCount++;
                    }
                    else if (comments < 300)
                    {
                        doc.Comments.Add(sub, "EYD: " + issue.Message + " Saran: \"" + issue.Suggestion + "\"");
                        comments++;
                    }
                }
            }
            if (found == 0) return "Tidak ditemukan kesalahan EYD yang terdeteksi aturan. 👍";
            return $"Ditemukan {found} catatan EYD: {fixedCount} diperbaiki otomatis, {comments} diberi komentar.";
        }

        private string Watermark(Word.Document doc, string text)
        {
            foreach (Word.Section sec in doc.Sections)
            {
                dynamic header = sec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                dynamic shape = header.Shapes.AddTextEffect(0, text, "Calibri", 80, 0, 0, 0, 0);
                shape.Name = "MasterBW_Watermark_" + Guid.NewGuid().ToString("N").Substring(0, 6);
                shape.TextEffect.NormalizedHeight = 0;
                shape.Line.Visible = 0;
                shape.Fill.Visible = -1;
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = 0xC0C0C0;
                shape.Fill.Transparency = 0.5f;
                shape.Rotation = 315;
                shape.LockAspectRatio = -1;
                shape.Height = WordHelpers.Cm(4);
                shape.Width = WordHelpers.Cm(16);
                shape.WrapFormat.AllowOverlap = -1;
                shape.WrapFormat.Type = 3; // wdWrapNone (di belakang teks)
                shape.RelativeHorizontalPosition = 0; // margin
                shape.RelativeVerticalPosition = 0;
                shape.Left = -999995; // wdShapeCenter
                shape.Top = -999995;
            }
            return "Watermark '" + text + "' ditambahkan.";
        }

        private string MailMerge(Word.Document doc, ActionStep s)
        {
            var path = s.Str("data_file");
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return "File data tidak ditemukan: " + path;
            var rows = TabularReader.Read(path, s.Str("sheet"));
            if (rows.Count < 2) return "Data kosong atau tidak memiliki header.";
            var headers = rows[0];
            var src = doc.Content.FormattedText;
            var output = _app.Documents.Add();
            for (int i = 1; i < rows.Count; i++)
            {
                int start = output.Content.End - 1;
                var target = output.Range(start, start);
                target.FormattedText = src;
                var inserted = output.Range(start, output.Content.End - 1);
                for (int c = 0; c < headers.Count; c++)
                {
                    var f = inserted.Duplicate.Find;
                    f.ClearFormatting();
                    var value = c < rows[i].Count ? rows[i][c] ?? "" : "";
                    if (value.Length > 250) value = value.Substring(0, 250);
                    f.Execute("{{" + headers[c] + "}}", false, false, false, false, false, true, Word.WdFindWrap.wdFindStop, false, value, Word.WdReplace.wdReplaceAll);
                }
                if (i < rows.Count - 1) output.Range(output.Content.End - 1, output.Content.End - 1).InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }
            return (rows.Count - 1) + " dokumen digabung ke dokumen baru (satu per halaman).";
        }

        private static string Truncate(string s, int max) => s == null ? "" : s.Length <= max ? s : s.Substring(0, max);
        private static int CountWords(string s) => (s ?? "").Split(new[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
    }

    /// <summary>Penyimpanan pustaka pengguna (referensi.json) untuk sitasi &amp; daftar pustaka.</summary>
    public sealed class ReferenceStore
    {
        public List<Reference> Items { get; set; } = new List<Reference>();

        public static ReferenceStore Load() => AppPaths.ReadJson<ReferenceStore>(AppPaths.ReferencesFile) ?? new ReferenceStore();

        public void Save() => AppPaths.WriteJson(AppPaths.ReferencesFile, this);

        /// <summary>Tambahkan bila belum ada; kembalikan nomor urut (untuk IEEE/Vancouver).</summary>
        public int AddOrGet(Reference r)
        {
            int idx = Items.FindIndex(x => string.Equals(x.Title?.Trim(), r.Title?.Trim(), StringComparison.OrdinalIgnoreCase) && x.Year == r.Year);
            if (idx >= 0) return idx + 1;
            if (string.IsNullOrEmpty(r.Id)) r.Id = "ref" + (Items.Count + 1);
            Items.Add(r);
            return Items.Count;
        }
    }

    /// <summary>Membaca CSV atau Excel (via Excel tersembunyi) menjadi baris teks.</summary>
    internal static class TabularReader
    {
        public static List<List<string>> Read(string path, string sheet)
        {
            if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
            {
                var lines = File.ReadAllLines(path, Encoding.UTF8);
                char sep = lines.Length > 0 && lines[0].Count(c => c == ';') > lines[0].Count(c => c == ',') ? ';' : ',';
                return lines.Where(l => l.Trim().Length > 0).Select(l => SplitCsv(l, sep)).ToList();
            }
            var excel = new Microsoft.Office.Interop.Excel.Application { Visible = false, DisplayAlerts = false };
            try
            {
                var wb = excel.Workbooks.Open(path, ReadOnly: true);
                var ws = string.IsNullOrWhiteSpace(sheet) ? (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1] : (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[sheet];
                var values = ws.UsedRange.Value2 as object[,];
                var result = new List<List<string>>();
                if (values != null)
                    for (int r = values.GetLowerBound(0); r <= values.GetUpperBound(0); r++)
                    {
                        var row = new List<string>();
                        for (int c = values.GetLowerBound(1); c <= values.GetUpperBound(1); c++) row.Add(Convert.ToString(values[r, c]) ?? "");
                        result.Add(row);
                    }
                wb.Close(false);
                return result;
            }
            finally { excel.Quit(); System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excel); }
        }

        private static List<string> SplitCsv(string line, char sep)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool q = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"') { if (q && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; } else q = !q; }
                else if (c == sep && !q) { result.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result;
        }
    }
}
