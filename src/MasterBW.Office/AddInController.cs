using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using MasterBW.Core.Academic;
using MasterBW.Core.Ai;
using MasterBW.Core.ExcelKit;
using MasterBW.Core.Tools;
using MasterBW.Core.Util;
using MasterBW.Office.Common;
using MasterBW.Office.ForExcel;
using MasterBW.Office.ForWord;
using MasterBW.Office.UI;
using Newtonsoft.Json.Linq;

namespace MasterBW.Office
{
    /// <summary>Satu perintah UI (ribbon, menu klik kanan, palet perintah).</summary>
    public sealed class UiCommand
    {
        public string Group;
        public string Menu;          // null = tombol langsung di grup
        public string Label;
        public string Tag;           // ui:..., tool:name|{json}, prefill:..., xl:...
        public string Image = "HappyFace";
        public bool Large;
        public bool Context;         // tampil juga di menu klik kanan
        public string Hint;
    }

    /// <summary>
    /// Otak add-in yang dipakai bersama oleh proyek VSTO Word dan Excel.
    /// Proyek VSTO hanya meneruskan event ribbon dan membuat CustomTaskPane.
    /// </summary>
    public sealed class AddInController : IDisposable
    {
        public OfficeApp App { get; }
        public AiService Ai { get; } = new AiService();
        public IActionExecutor Executor { get; }
        private readonly object _application;
        private KeyboardHook _hook;
        private List<UiCommand> _commands;

        /// <summary>Diisi proyek VSTO: tampilkan panel (buat bila belum ada) lalu kembalikan ChatPane-nya.</summary>
        public Func<ChatPane> ShowPane { get; set; }
        /// <summary>Diisi proyek VSTO: tampil/sembunyikan panel di jendela aktif.</summary>
        public Action TogglePane { get; set; }

        public AddInController(OfficeApp app, object application)
        {
            App = app;
            _application = application;
            Executor = app == OfficeApp.Word ? (IActionExecutor)new WordExecutor(application) : new ExcelExecutor(application);
        }

        public void Startup()
        {
            if (SynchronizationContext.Current == null) SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            try { _hook = new KeyboardHook(Ai.Settings.PaletteHotkey, ShowPalette); } catch (Exception ex) { Log.Warn("Hotkey: " + ex.Message); }
            // nyalakan mesin AI di latar belakang agar perintah pertama lebih cepat
            Task.Run(async () =>
            {
                try { await EngineLauncher.EnsureRunningAsync(Ai.Client, null, CancellationToken.None); }
                catch (Exception ex) { Log.Warn("Engine belum bisa dinyalakan: " + ex.Message); }
            });
            Log.Info("MASTER BW dimuat di " + App);
        }

        public ChatPane CreateChatPane() => new ChatPane(this);

        public void Dispose() => _hook?.Dispose();

        // ================================================================== perintah
        public List<UiCommand> Commands => _commands ?? (_commands = App == OfficeApp.Word ? WordCommands() : ExcelCommands());

        private static string Tool(string name, object args = null) => "tool:" + name + "|" + (args == null ? "{}" : JObject.FromObject(args).ToString(Newtonsoft.Json.Formatting.None));

        private static List<UiCommand> Common(string appName)
        {
            return new List<UiCommand>
            {
                new UiCommand { Group = "Asisten", Label = "Panel AI", Tag = "ui:chat", Image = "HappyFace", Large = true, Hint = "Buka/tutup panel chat MASTER BW" },
                new UiCommand { Group = "Asisten", Label = "Palet Perintah", Tag = "ui:palette", Image = "FindDialog", Large = true, Hint = "Cari & jalankan semua perintah (Ctrl+Alt+Space)" },
                new UiCommand { Group = "Asisten", Label = "Batalkan Aksi AI", Tag = "ui:undo", Image = "Undo" },
                new UiCommand { Group = "Asisten", Label = "Pengaturan", Tag = "ui:settings", Image = "ControlProperties" },
            };
        }

        private List<UiCommand> WordCommands()
        {
            var list = Common("Word");
            var tasks = new (string key, string label, string output)[]
            {
                ("ringkas", "Ringkas", "replace"), ("parafrase", "Parafrase", "replace"), ("perbaiki", "Perbaiki Ejaan & Tata Bahasa", "replace"),
                ("formal", "Jadikan Bahasa Formal", "replace"), ("akademik", "Gaya Bahasa Akademik", "replace"), ("santai", "Gaya Santai", "replace"),
                ("sederhanakan", "Sederhanakan", "replace"), ("perpanjang", "Kembangkan/Perpanjang", "replace"), ("lanjutkan", "Lanjutkan Tulisan", "below"),
                ("poin", "Jadikan Poin-poin", "replace"), ("terjemah_en", "Terjemahkan ke Inggris", "replace"), ("terjemah_id", "Terjemahkan ke Indonesia", "replace"),
                ("judul", "Usulkan Judul", "below"), ("abstrak", "Buat Abstrak", "below"), ("kesimpulan", "Buat Kesimpulan", "below"),
                ("pertanyaan", "Buat Soal dari Teks", "below"), ("email", "Ubah Jadi Email Resmi", "replace"),
            };
            foreach (var t in tasks)
                list.Add(new UiCommand { Group = "Tulis dengan AI", Menu = "Olah Teks Terpilih", Label = t.label, Tag = Tool("transform_text", new { task = t.key, scope = "selection", output = t.output }), Image = "EditDocument",
                                         Context = t.key == "ringkas" || t.key == "parafrase" || t.key == "perbaiki" || t.key == "formal" || t.key == "terjemah_en" || t.key == "lanjutkan" });
            list.Add(new UiCommand { Group = "Tulis dengan AI", Label = "Tulis Bagian Baru", Tag = "prefill:Tuliskan ", Image = "WordArtInsertGallery", Large = true, Hint = "AI menulis paragraf/bab baru" });
            list.Add(new UiCommand { Group = "Tulis dengan AI", Label = "Ringkas Seluruh Dokumen", Tag = Tool("transform_text", new { task = "ringkas", scope = "document", output = "new_document" }), Image = "ReviewNewComment" });

            list.Add(new UiCommand { Group = "Dokumen", Label = "Galeri Templat", Tag = "ui:templates", Image = "FileNew", Large = true, Hint = "Surat, SK, notulen, proposal, skripsi, modul ajar, CV…" });
            foreach (var cat in DocumentTemplates.All.Select(t => t.Category).Distinct())
                list.Add(new UiCommand { Group = "Dokumen", Menu = "Templat per Kategori", Label = cat, Tag = "ui:templates:" + cat, Image = "FileNew" });
            foreach (var t in DocumentTemplates.All)
                list.Add(new UiCommand { Group = "Palet", Label = "Templat: " + t.Name, Tag = "ui:template:" + t.Id, Hint = t.Description });
            list.Add(new UiCommand { Group = "Dokumen", Label = "Kop Surat", Tag = Tool("insert_letterhead", new { }), Image = "HeaderInsertGallery" });
            list.Add(new UiCommand { Group = "Dokumen", Label = "Tanda Tangan", Tag = Tool("insert_signature", new { align = "right" }), Image = "SignatureLineInsert" });
            list.Add(new UiCommand { Group = "Dokumen", Label = "Halaman Sampul", Tag = "prefill:Buat halaman sampul makalah berjudul \"", Image = "CoverPageInsertGallery" });

            foreach (var p in new[] { "Skripsi", "Proposal", "Makalah", "Laporan", "Jurnal", "SuratDinas" })
                list.Add(new UiCommand { Group = "Akademik", Menu = "Format Baku", Label = "Format " + p, Tag = Tool("apply_academic_format", new { preset = p }), Image = "StylesPane" });
            foreach (var (k, l) in new[] { ("bab_romawi", "BAB I → A. → 1."), ("bab_angka", "BAB 1 → 1.1 → 1.1.1"), ("desimal", "1. → 1.1 → 1.1.1"), ("none", "Hapus penomoran") })
                list.Add(new UiCommand { Group = "Akademik", Menu = "Penomoran Bab", Label = l, Tag = Tool("heading_numbering", new { scheme = k }), Image = "MultilevelListGallery" });
            list.Add(new UiCommand { Group = "Akademik", Label = "Daftar Isi", Tag = Tool("insert_toc", new { levels = 3 }), Image = "TableOfContentsGallery" });
            list.Add(new UiCommand { Group = "Akademik", Menu = "Daftar Gambar/Tabel", Label = "Daftar Gambar", Tag = Tool("insert_table_of_figures", new { label = "Gambar" }), Image = "TableOfFiguresInsert" });
            list.Add(new UiCommand { Group = "Akademik", Menu = "Daftar Gambar/Tabel", Label = "Daftar Tabel", Tag = Tool("insert_table_of_figures", new { label = "Tabel" }), Image = "TableOfFiguresInsert" });
            list.Add(new UiCommand { Group = "Akademik", Menu = "Daftar Gambar/Tabel", Label = "Caption Gambar", Tag = "prefill:Beri caption gambar: ", Image = "CaptionInsert" });
            list.Add(new UiCommand { Group = "Akademik", Menu = "Daftar Gambar/Tabel", Label = "Caption Tabel", Tag = "prefill:Beri caption tabel: ", Image = "CaptionInsert" });
            list.Add(new UiCommand { Group = "Akademik", Label = "Perbarui Daftar Isi", Tag = Tool("update_fields"), Image = "UpdateTableOfContents" });
            list.Add(new UiCommand { Group = "Akademik", Label = "Sisipkan Sitasi", Tag = "prefill:Sisipkan sitasi APA: penulis ..., tahun ..., judul ..., penerbit/jurnal ...", Image = "CitationInsert" });
            list.Add(new UiCommand { Group = "Akademik", Label = "Daftar Pustaka", Tag = Tool("insert_bibliography", new { }), Image = "BibliographyGallery" });
            list.Add(new UiCommand { Group = "Akademik", Label = "Rapikan Daftar Pustaka", Tag = Tool("format_bibliography"), Image = "SortDialog", Hint = "Pilih paragraf daftar pustaka dulu" });
            list.Add(new UiCommand { Group = "Akademik", Label = "Cek Ejaan EYD", Tag = Tool("check_spelling", new { scope = "document" }), Image = "SpellingAndGrammar", Large = true, Context = true });
            list.Add(new UiCommand { Group = "Akademik", Label = "Perbaiki Kata Tidak Baku", Tag = Tool("check_spelling", new { scope = "document", auto_fix = true }), Image = "SpellingAndGrammar" });

            list.Add(new UiCommand { Group = "Rapikan", Label = "Rapikan Dokumen", Tag = Tool("clean_document", new { }), Image = "ClearFormatting", Large = true, Hint = "Spasi ganda, baris kosong, sisa copy-paste PDF" });
            foreach (var (k, l) in new[] { ("akademik", "Tabel Gaya Akademik (APA)"), ("grid", "Tabel Bergaris Penuh"), ("modern", "Tabel Modern") })
                list.Add(new UiCommand { Group = "Rapikan", Menu = "Format Semua Tabel", Label = l, Tag = Tool("format_table", new { scope = "all", style = k }), Image = "TableStyleModify" });
            foreach (var (pos, fmt, l) in new[] { ("bottom_center", "arabic", "Bawah tengah (1, 2, 3)"), ("bottom_right", "arabic", "Bawah kanan (1, 2, 3)"), ("top_right", "arabic", "Atas kanan (1, 2, 3)"), ("bottom_center", "roman_lower", "Bawah tengah (i, ii, iii)") })
                list.Add(new UiCommand { Group = "Rapikan", Menu = "Nomor Halaman", Label = l, Tag = Tool("insert_page_numbers", new { position = pos, format = fmt }), Image = "PageNumberInsertGallery" });
            list.Add(new UiCommand { Group = "Rapikan", Label = "Simpan PDF", Tag = Tool("export_pdf", new { }), Image = "FileSaveAsPdfOrXps" });
            list.Add(new UiCommand { Group = "Rapikan", Label = "Statistik Dokumen", Tag = Tool("document_stats"), Image = "WordCount" });
            list.Add(new UiCommand { Group = "Asisten", Label = "Tanya AI tentang teks ini…", Tag = "prefill:Tentang teks terpilih: ", Image = "HappyFace", Context = true });
            return list;
        }

        private List<UiCommand> ExcelCommands()
        {
            var list = Common("Excel");
            list.Add(new UiCommand { Group = "Rumus", Label = "Buat Rumus", Tag = "prefill:Buat rumus untuk ", Image = "FunctionWizard", Large = true, Hint = "Jelaskan kebutuhan dalam bahasa Indonesia" });
            list.Add(new UiCommand { Group = "Rumus", Label = "Jelaskan Rumus", Tag = Tool("explain_formula", new { }), Image = "FormulaEvaluate", Context = true });
            list.Add(new UiCommand { Group = "Rumus", Label = "Perbaiki Rumus", Tag = Tool("fix_formula", new { }), Image = "TraceError", Context = true });
            list.Add(new UiCommand { Group = "Rumus", Label = "Terbilang", Tag = "xl:terbilang", Image = "TextBoxInsert", Context = true, Hint = "Angka terpilih → terbilang di kolom kanannya" });

            foreach (var (ops, l) in new[] { ("trim,remove_nonprintable", "Hapus spasi & karakter aneh"), ("proper_name", "Rapikan nama orang"), ("upper", "HURUF BESAR"), ("lower", "huruf kecil"),
                                             ("text_to_number,fix_rupiah", "Teks/Rupiah → angka"), ("normalize_date", "Rapikan tanggal"), ("normalize_phone", "Rapikan nomor HP (+62)"),
                                             ("trim,remove_nonprintable,text_to_number,normalize_date", "Bersihkan semua (otomatis)") })
                list.Add(new UiCommand { Group = "Data", Menu = "Bersihkan Data", Label = l, Tag = "xl:clean:" + ops, Image = "ClearAll", Context = ops == "trim,remove_nonprintable" });
            list.Add(new UiCommand { Group = "Data", Label = "Hapus Duplikat", Tag = Tool("remove_duplicates", new { }), Image = "RemoveDuplicates" });
            list.Add(new UiCommand { Group = "Data", Label = "Hapus Baris Kosong", Tag = Tool("delete_cells", new { type = "blank_rows" }), Image = "TableDeleteRows" });
            list.Add(new UiCommand { Group = "Data", Label = "Jadikan Tabel", Tag = Tool("format_as_table", new { }), Image = "TableInsertExcel" });
            list.Add(new UiCommand { Group = "Data", Label = "Lebar Otomatis", Tag = Tool("autofit", new { }), Image = "ColumnWidth" });
            list.Add(new UiCommand { Group = "Data", Label = "Kunci Baris Judul", Tag = Tool("freeze_panes", new { cell = "A2" }), Image = "FreezePanes" });

            list.Add(new UiCommand { Group = "Analisis", Label = "Analisis & Insight", Tag = Tool("analyze_data", new { output = "chat" }), Image = "ChartInsert", Large = true });
            list.Add(new UiCommand { Group = "Analisis", Label = "Dashboard Otomatis", Tag = Tool("create_dashboard", new { }), Image = "PivotChartInsert", Large = true });
            list.Add(new UiCommand { Group = "Analisis", Label = "Rekap per Kategori", Tag = "prefill:Buat rekap total per ", Image = "Subtotals" });
            list.Add(new UiCommand { Group = "Analisis", Label = "PivotTable", Tag = "prefill:Buat pivot ", Image = "PivotTableInsert" });
            foreach (var (k, l) in new[] { ("column", "Kolom"), ("bar", "Batang"), ("line", "Garis"), ("pie", "Lingkaran"), ("doughnut", "Donat"), ("area", "Area"), ("scatter", "Sebar (XY)"), ("combo", "Kombinasi") })
                list.Add(new UiCommand { Group = "Analisis", Menu = "Grafik Cepat", Label = "Grafik " + l, Tag = "xl:chart:" + k, Image = "ChartInsert" });

            foreach (var (k, l) in new[] { ("descriptive_stats", "Statistik Deskriptif"), ("correlation", "Korelasi Pearson"), ("normality_test", "Uji Normalitas"), ("anova", "ANOVA Satu Jalur"),
                                           ("validity_reliability", "Uji Validitas & Reliabilitas"), ("chi_square", "Uji Chi-Square"), ("detect_outliers", "Deteksi Outlier"), ("forecast", "Peramalan Tren") })
                list.Add(new UiCommand { Group = "Statistik", Menu = "Uji Statistik", Label = l, Tag = "xl:stat:" + k, Image = "AnalysisToolPak", Context = k == "descriptive_stats" });
            list.Add(new UiCommand { Group = "Statistik", Menu = "Uji Statistik", Label = "Regresi Linear…", Tag = "prefill:Lakukan regresi linear dengan Y = kolom ... dan X = kolom ...", Image = "AnalysisToolPak" });
            list.Add(new UiCommand { Group = "Statistik", Menu = "Uji Statistik", Label = "Uji t…", Tag = "prefill:Lakukan uji t berpasangan antara kolom ... (pretest) dan kolom ... (posttest)", Image = "AnalysisToolPak" });

            foreach (var (l, p) in new[] { ("Tabel Angsuran Pinjaman", "Buat tabel angsuran pinjaman Rp100.000.000 bunga 10% per tahun selama 24 bulan metode anuitas"),
                                           ("Penyusutan Aset", "Buat jadwal penyusutan aset harga Rp50.000.000 nilai sisa Rp5.000.000 umur 5 tahun metode garis lurus"),
                                           ("Kelayakan Investasi (NPV/IRR)", "Hitung NPV dan IRR untuk arus kas -100000000, 30000000, 40000000, 50000000 dengan diskonto 10%"),
                                           ("PPh 21", "Hitung PPh 21 gaji Rp8.000.000 per bulan status K/1"), ("PPN", "Hitung PPN untuk harga jual Rp10.000.000"),
                                           ("Titik Impas (BEP)", "Hitung BEP biaya tetap Rp20.000.000 harga Rp50.000 biaya variabel Rp30.000"),
                                           ("Rasio Keuangan", "Hitung rasio keuangan: aset lancar ..., utang lancar ..., total aset ..., total utang ..., ekuitas ..., penjualan ..., laba bersih ...") })
                list.Add(new UiCommand { Group = "Keuangan", Menu = "Kalkulator Keuangan", Label = l, Tag = "prefill:" + p, Image = "CurrencyFormat" });
            foreach (var t in ExcelTemplates.All)
                list.Add(new UiCommand { Group = "Keuangan", Menu = "Templat Lembar Kerja", Label = "[" + t.Category + "] " + t.Name, Tag = Tool("create_template", new { template = t.Id }), Image = "TableStylesGallery", Hint = t.Description });
            list.Add(new UiCommand { Group = "Asisten", Label = "Tanya AI tentang data ini…", Tag = "prefill:Tentang data terpilih: ", Image = "HappyFace", Context = true });
            return list;
        }

        public IEnumerable<(string, Action)> QuickActions()
        {
            if (App == OfficeApp.Word)
            {
                yield return ("Ringkas", () => OnCommand(Tool("transform_text", new { task = "ringkas" })));
                yield return ("Parafrase", () => OnCommand(Tool("transform_text", new { task = "parafrase" })));
                yield return ("Perbaiki", () => OnCommand(Tool("transform_text", new { task = "perbaiki" })));
                yield return ("Formal", () => OnCommand(Tool("transform_text", new { task = "formal" })));
                yield return ("Cek EYD", () => OnCommand(Tool("check_spelling", new { scope = "document" })));
                yield return ("Templat", () => OnCommand("ui:templates"));
            }
            else
            {
                yield return ("Jelaskan rumus", () => OnCommand(Tool("explain_formula", new { })));
                yield return ("Perbaiki rumus", () => OnCommand(Tool("fix_formula", new { })));
                yield return ("Insight", () => OnCommand(Tool("analyze_data", new { output = "chat" })));
                yield return ("Dashboard", () => OnCommand(Tool("create_dashboard", new { })));
                yield return ("Deskriptif", () => OnCommand("xl:stat:descriptive_stats"));
                yield return ("Bersihkan", () => OnCommand("xl:clean:trim,remove_nonprintable,text_to_number,normalize_date"));
            }
        }

        public IEnumerable<string> Examples()
        {
            if (App == OfficeApp.Word)
                return new[] { "buat surat undangan rapat komite hari Senin jam 09.00 di aula", "format skripsi standar lalu beri penomoran BAB I", "tulis latar belakang tentang literasi digital siswa SD", "buat tabel jadwal piket 5 hari" };
            return new[] { "tambah kolom total = qty × harga format rupiah", "rekap total penjualan per wilayah", "uji validitas dan reliabilitas data kuesioner ini", "buat dashboard dari data ini" };
        }

        // ================================================================== dispatch
        public void OnCommand(string tag)
        {
            try { _ = OnCommandAsync(tag); }
            catch (Exception ex) { Log.Error("Perintah " + tag, ex); MessageBox.Show(ex.Message, "MASTER BW"); }
        }

        private async Task OnCommandAsync(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return;
            if (tag == "ui:chat") { if (TogglePane != null) TogglePane(); else ShowPane?.Invoke(); return; }
            if (tag == "ui:palette") { ShowPalette(); return; }
            if (tag == "ui:settings") { using (var f = new SettingsForm(Ai)) f.ShowDialog(); return; }
            if (tag == "ui:undo") { MessageBox.Show(Executor.UndoLastPlan(), "MASTER BW"); return; }
            if (tag.StartsWith("ui:template:"))
            {
                var t = DocumentTemplates.Find(tag.Substring(12));
                if (t != null) await RunPlan(ActionPlan.Single("create_from_template", new { template = t.Id, ai_fill = true }), "Templat: " + t.Name);
                return;
            }
            if (tag.StartsWith("ui:templates"))
            {
                var cat = tag.Length > 13 ? tag.Substring(13) : null;
                using (var f = new TemplatePickerForm(Ai.Settings, cat))
                    if (f.ShowDialog() == DialogResult.OK && f.Result != null) await RunPlan(f.Result, f.Result.Message);
                return;
            }
            if (tag.StartsWith("prefill:")) { ShowPane?.Invoke()?.Prefill(tag.Substring(8)); return; }
            if (tag.StartsWith("tool:"))
            {
                var bar = tag.IndexOf('|');
                var name = bar > 0 ? tag.Substring(5, bar - 5) : tag.Substring(5);
                var args = bar > 0 ? JObject.Parse(tag.Substring(bar + 1)) : new JObject();
                var def = ToolCatalog.Find(App, name);
                var plan = new ActionPlan { Message = def?.Label ?? name, Steps = { new ActionStep { Action = name, Args = args } } };
                ToolCatalog.Validate(App, plan);
                await RunPlan(plan, def?.Label ?? name);
                return;
            }
            if (tag.StartsWith("xl:"))
            {
                var sel = ExcelSelection();
                if (tag == "xl:terbilang")
                {
                    dynamic app = _application;
                    dynamic s = app.Selection;
                    string dest = (string)s.Offset(0, s.Columns.Count).Cells(1, 1).Address(false, false);
                    await RunPlan(ActionPlan.Single("terbilang", new { source = sel, destination = dest, rupiah = true }), "Terbilang " + sel);
                }
                else if (tag.StartsWith("xl:clean:"))
                    await RunPlan(ActionPlan.Single("clean_data", new { range = sel, operations = tag.Substring(9).Split(',') }), "Bersihkan data " + sel);
                else if (tag.StartsWith("xl:chart:"))
                    await RunPlan(ActionPlan.Single("create_chart", new { source = SingleCellToRegion(sel), type = tag.Substring(9), title = "" }), "Grafik " + tag.Substring(9));
                else if (tag.StartsWith("xl:stat:"))
                {
                    var name = tag.Substring(8);
                    object args = name == "forecast" ? (object)new { y_range = sel, periods = 3 } : name == "detect_outliers" ? (object)new { range = sel } : new { range = SingleCellToRegion(sel), has_headers = true };
                    await RunPlan(ActionPlan.Single(name, args), ToolCatalog.Find(App, name)?.Label ?? name);
                }
            }
        }

        private string ExcelSelection()
        {
            try { dynamic app = _application; return (string)app.Selection.Address(false, false); }
            catch { return ""; }
        }

        /// <summary>Jika yang dipilih hanya satu sel, pakai area data di sekitarnya.</summary>
        private string SingleCellToRegion(string sel) => sel.Contains(":") || sel.Contains(",") ? sel : "";

        private Task RunPlan(ActionPlan plan, string title)
        {
            var pane = ShowPane?.Invoke();
            if (pane == null) { MessageBox.Show("Panel MASTER BW tidak dapat dibuka.", "MASTER BW"); return Task.CompletedTask; }
            return pane.RunPlanAsync(plan, title);
        }

        public void ShowPalette()
        {
            var cmds = Commands.Where(c => c.Tag != "ui:palette").Select(c => new CommandPaletteForm.Command
            {
                Label = (c.Menu != null ? c.Menu + " › " : c.Group == "Palet" ? "" : c.Group + " › ") + c.Label,
                Hint = c.Hint,
                Run = () => OnCommand(c.Tag),
            }).ToList();
            foreach (var t in ToolCatalog.For(App))
                if (!cmds.Any(c => c.Label.EndsWith(t.Label)))
                    cmds.Add(new CommandPaletteForm.Command { Label = "Aksi › " + t.Label, Hint = t.Description, Run = () => ShowPane?.Invoke()?.Prefill(t.Label + ": ") });
            var f = new CommandPaletteForm(cmds, text => ShowPane?.Invoke()?.SubmitAsync(text));
            f.Show();
            f.Activate();
        }

        public void ShowText(string title, string body)
        {
            Action<string> insert = null;
            if (App == OfficeApp.Word)
                insert = text =>
                {
                    var step = new ActionStep { Action = "insert_text", Args = new JObject { ["text"] = text, ["position"] = "cursor" } };
                    _ = Executor.ExecuteAsync(step, new ExecutionSession { Ai = Ai });
                };
            new TextViewerForm(title, body, insert).Show();
        }

        // ================================================================== ribbon XML
        public string RibbonXml(string tabLabel = "MASTER BW")
        {
            XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
            int id = 0;
            XElement Button(UiCommand c, bool inMenu)
            {
                var b = new XElement(ns + "button",
                    new XAttribute("id", "mbw" + (id++)),
                    new XAttribute("label", c.Label),
                    new XAttribute("tag", c.Tag),
                    new XAttribute("imageMso", c.Image ?? "HappyFace"),
                    new XAttribute("onAction", "OnAction"));
                if (!inMenu) b.Add(new XAttribute("size", c.Large ? "large" : "normal"));
                if (!string.IsNullOrEmpty(c.Hint)) b.Add(new XAttribute("screentip", c.Label), new XAttribute("supertip", c.Hint));
                return b;
            }
            var tab = new XElement(ns + "tab", new XAttribute("id", "MasterBWTab"), new XAttribute("label", tabLabel));
            foreach (var g in Commands.Where(c => c.Group != "Palet").GroupBy(c => c.Group))
            {
                var group = new XElement(ns + "group", new XAttribute("id", "mbwg" + (id++)), new XAttribute("label", g.Key));
                foreach (var item in g.Where(c => c.Menu == null)) group.Add(Button(item, false));
                foreach (var m in g.Where(c => c.Menu != null).GroupBy(c => c.Menu))
                {
                    var menu = new XElement(ns + "menu", new XAttribute("id", "mbwm" + (id++)), new XAttribute("label", m.Key), new XAttribute("imageMso", m.First().Image), new XAttribute("size", "large"), new XAttribute("itemSize", "normal"));
                    foreach (var item in m) menu.Add(Button(item, true));
                    group.Add(menu);
                }
                tab.Add(group);
            }
            var ctxMenuId = App == OfficeApp.Word ? "ContextMenuText" : "ContextMenuCell";
            var ctxMenu = new XElement(ns + "menu", new XAttribute("id", "mbwctx"), new XAttribute("label", "MASTER BW"), new XAttribute("imageMso", "HappyFace"), new XAttribute("insertBeforeMso", App == OfficeApp.Word ? "Cut" : "Cut"));
            foreach (var c in Commands.Where(c => c.Context)) ctxMenu.Add(Button(c, true));
            var root = new XElement(ns + "customUI", new XAttribute("onLoad", "OnLoad"),
                new XElement(ns + "ribbon", new XElement(ns + "tabs", tab)),
                new XElement(ns + "contextMenus", new XElement(ns + "contextMenu", new XAttribute("idMso", ctxMenuId), ctxMenu)));
            return root.ToString();
        }
    }
}
