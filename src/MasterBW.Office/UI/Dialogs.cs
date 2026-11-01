using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using MasterBW.Core.Academic;
using MasterBW.Core.Ai;
using MasterBW.Core.Tools;
using MasterBW.Core.Util;
using MasterBW.Office.Common;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MasterBW.Office.UI
{
    /// <summary>Pratinjau rencana aksi: pengguna mencentang langkah, bisa mengedit argumen, lalu Jalankan/Batal.</summary>
    public sealed class PlanPreviewForm : Form
    {
        private readonly ActionPlan _plan;
        private readonly OfficeApp _app;
        private readonly ListView _list = new ListView { Dock = DockStyle.Fill, View = View.Details, CheckBoxes = true, FullRowSelect = true, HideSelection = false, Font = Theme.Body };
        private readonly TextBox _detail = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = Theme.Mono, WordWrap = true };
        private readonly Label _warn = new Label { Dock = DockStyle.Bottom, Height = 40, ForeColor = Theme.Error, Font = Theme.Small };
        private int _current = -1;

        public PlanPreviewForm(ActionPlan plan, OfficeApp app)
        {
            _plan = plan; _app = app;
            Text = "MASTER BW – Pratinjau aksi";
            Width = 760; Height = 520; StartPosition = FormStartPosition.CenterScreen; Font = Theme.Body; BackColor = Theme.Back;
            MinimizeBox = false; ShowIcon = false;

            var header = new Label { Dock = DockStyle.Top, Height = 54, Padding = new Padding(10, 8, 10, 4), Font = Theme.Body,
                                     Text = (string.IsNullOrWhiteSpace(plan.Message) ? "" : plan.Message + "\n") + "Periksa langkah di bawah. Hilangkan centang untuk melewati langkah; argumen bisa diedit di panel kanan." };
            _list.Columns.Add("No", 40); _list.Columns.Add("Langkah", 190); _list.Columns.Add("Ringkasan", 300);
            for (int i = 0; i < plan.Steps.Count; i++)
            {
                var s = plan.Steps[i];
                var def = ToolCatalog.Find(app, s.Action);
                var item = new ListViewItem(new[] { (i + 1).ToString(), (def?.Destructive == true ? "⚠ " : "") + (def?.Label ?? s.Action), Summary(s) }) { Checked = s.Enabled, Tag = s };
                if (s.Warnings.Count > 0) item.ForeColor = Theme.Error;
                _list.Items.Add(item);
            }
            _list.SelectedIndexChanged += (o, e) => SelectStep();
            var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 420 };
            split.Panel1.Controls.Add(_list);
            split.Panel2.Controls.Add(_detail);
            split.Panel2.Controls.Add(_warn);

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 46, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
            var run = Theme.Button("▶ Jalankan", true);
            var cancel = Theme.Button("Batal");
            run.Click += (o, e) => { SaveEdit(); foreach (ListViewItem it in _list.Items) ((ActionStep)it.Tag).Enabled = it.Checked; DialogResult = DialogResult.OK; };
            cancel.Click += (o, e) => DialogResult = DialogResult.Cancel;
            buttons.Controls.Add(run); buttons.Controls.Add(cancel);
            AcceptButton = run; CancelButton = cancel;

            Controls.Add(split); Controls.Add(buttons); Controls.Add(header);
            if (_list.Items.Count > 0) _list.Items[0].Selected = true;
        }

        private static string Summary(ActionStep s)
        {
            var parts = s.Args.Properties().Take(3).Select(p => p.Name + "=" + Short(p.Value.Type == JTokenType.String ? (string)p.Value : p.Value.ToString(Formatting.None)));
            return string.Join(", ", parts);
        }

        private static string Short(string v) => v == null ? "" : v.Length > 40 ? v.Substring(0, 40) + "…" : v;

        private void SelectStep()
        {
            SaveEdit();
            if (_list.SelectedItems.Count == 0) return;
            _current = _list.SelectedIndices[0];
            var s = (ActionStep)_list.Items[_current].Tag;
            _detail.Text = s.Args.ToString(Formatting.Indented).Replace("\n", "\r\n");
            var def = ToolCatalog.Find(_app, s.Action);
            _warn.Text = string.Join(" ", s.Warnings) + (def == null ? "" : "\r\n" + def.Description);
        }

        private void SaveEdit()
        {
            if (_current < 0 || _current >= _list.Items.Count) return;
            var s = (ActionStep)_list.Items[_current].Tag;
            try
            {
                var o = JObject.Parse(_detail.Text);
                s.Args = o;
                _list.Items[_current].SubItems[2].Text = Summary(s);
            }
            catch { /* abaikan JSON tidak valid; argumen lama dipakai */ }
        }
    }

    /// <summary>Jendela teks (penjelasan rumus, insight, kode VBA) dengan tombol salin.</summary>
    public sealed class TextViewerForm : Form
    {
        public TextViewerForm(string title, string text, Action<string> insert = null)
        {
            Text = "MASTER BW – " + title; Width = 640; Height = 480; StartPosition = FormStartPosition.CenterScreen; Font = Theme.Body; ShowIcon = false;
            var box = new RichTextBox { Dock = DockStyle.Fill, Text = text, Font = text.Contains("Sub ") ? Theme.Mono : Theme.Body, ReadOnly = false, BorderStyle = BorderStyle.None, BackColor = Color.White };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(6) };
            var close = Theme.Button("Tutup");
            var copy = Theme.Button("Salin", true);
            close.Click += (o, e) => Close();
            copy.Click += (o, e) => { try { Clipboard.SetText(box.Text); copy.Text = "Tersalin ✓"; } catch { } };
            buttons.Controls.Add(close); buttons.Controls.Add(copy);
            if (insert != null)
            {
                var ins = Theme.Button("Sisipkan ke dokumen");
                ins.Click += (o, e) => { insert(box.SelectionLength > 0 ? box.SelectedText : box.Text); Close(); };
                buttons.Controls.Add(ins);
            }
            Controls.Add(box); Controls.Add(buttons);
        }
    }

    /// <summary>Galeri templat dokumen Word dengan isian kolom.</summary>
    public sealed class TemplatePickerForm : Form
    {
        private readonly ListBox _list = new ListBox { Dock = DockStyle.Fill, Font = Theme.Body, IntegralHeight = false };
        private readonly DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, BackgroundColor = Color.White };
        private readonly Label _desc = new Label { Dock = DockStyle.Top, Height = 40, Font = Theme.Small, ForeColor = Theme.Muted };
        private readonly CheckBox _ai = new CheckBox { Text = "AI mengisi bagian isi secara otomatis", Checked = true, AutoSize = true };
        public ActionPlan Result { get; private set; }

        public TemplatePickerForm(MasterBWSettings settings, string filterCategory = null)
        {
            Text = "MASTER BW – Galeri templat dokumen"; Width = 860; Height = 560; StartPosition = FormStartPosition.CenterScreen; Font = Theme.Body; ShowIcon = false;
            var items = DocumentTemplates.All.Where(t => filterCategory == null || t.Category == filterCategory).ToList();
            foreach (var t in items) _list.Items.Add(new Item(t));
            _grid.Columns.Add("label", "Kolom"); _grid.Columns.Add("value", "Isi");
            _grid.Columns[0].ReadOnly = true; _grid.Columns[0].FillWeight = 40;
            _list.SelectedIndexChanged += (o, e) =>
            {
                var t = (_list.SelectedItem as Item)?.Template;
                if (t == null) return;
                _desc.Text = t.Category + " • " + t.Description;
                _grid.Rows.Clear();
                var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["instansi"] = settings.Institution, ["institusi"] = settings.Institution, ["alamat_instansi"] = settings.InstitutionAddress,
                    ["kota"] = settings.City, ["jabatan"] = settings.UserTitle, ["nama_penandatangan"] = settings.UserName, ["nip"] = settings.UserNip, ["penulis"] = settings.UserName,
                };
                foreach (var f in t.Fields)
                {
                    int r = _grid.Rows.Add(f.Label, defaults.TryGetValue(f.Key, out var v) && !string.IsNullOrEmpty(v) ? v : f.Default ?? "");
                    _grid.Rows[r].Tag = f.Key;
                }
            };
            var left = new Panel { Dock = DockStyle.Left, Width = 280 };
            left.Controls.Add(_list);
            left.Controls.Add(new Label { Dock = DockStyle.Top, Text = "Pilih templat (" + items.Count + ")", Font = Theme.Bold, Height = 24 });
            var right = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8, 0, 0, 0) };
            right.Controls.Add(_grid); right.Controls.Add(_desc);
            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 46, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
            var ok = Theme.Button("Buat dokumen", true); var cancel = Theme.Button("Batal");
            ok.Click += (o, e) =>
            {
                var t = (_list.SelectedItem as Item)?.Template;
                if (t == null) return;
                _grid.EndEdit();
                var fields = new JObject();
                foreach (DataGridViewRow row in _grid.Rows) fields[(string)row.Tag] = Convert.ToString(row.Cells[1].Value) ?? "";
                Result = ActionPlan.Single("create_from_template", new { template = t.Id, fields, ai_fill = _ai.Checked }, "Membuat " + t.Name);
                Result.Steps[0].Args["fields"] = fields;
                DialogResult = DialogResult.OK;
            };
            cancel.Click += (o, e) => DialogResult = DialogResult.Cancel;
            bottom.Controls.Add(ok); bottom.Controls.Add(cancel); bottom.Controls.Add(_ai);
            Controls.Add(right); Controls.Add(left); Controls.Add(bottom);
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
        }

        private sealed class Item
        {
            public readonly DocumentTemplate Template;
            public Item(DocumentTemplate t) { Template = t; }
            public override string ToString() => "[" + Template.Category + "] " + Template.Name;
        }
    }

    /// <summary>Pengaturan MASTER BW + status mesin AI.</summary>
    public sealed class SettingsForm : Form
    {
        public SettingsForm(AiService ai)
        {
            Text = "MASTER BW – Pengaturan"; Width = 620; Height = 640; StartPosition = FormStartPosition.CenterScreen; Font = Theme.Body; ShowIcon = false;
            var settings = MasterBWSettings.Load();
            var grid = new PropertyGrid { Dock = DockStyle.Fill, SelectedObject = settings, ToolbarVisible = false, HelpVisible = false, PropertySort = PropertySort.NoSort };
            var status = new Label { Dock = DockStyle.Top, Height = 60, Padding = new Padding(8), Font = Theme.Small, Text = "Memeriksa mesin AI…" };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 84, Padding = new Padding(6) };
            var save = Theme.Button("Simpan", true);
            var models = Theme.Button("Buka folder model");
            var download = Theme.Button("Unduh model AI");
            var logs = Theme.Button("Buka log");
            var restart = Theme.Button("Mulai ulang mesin AI");
            save.Click += (o, e) => { settings.Save(); ai.ReloadSettings(); DialogResult = DialogResult.OK; };
            models.Click += (o, e) => { Directory.CreateDirectory(AppPaths.ModelsDir); Process.Start("explorer.exe", AppPaths.ModelsDir); };
            logs.Click += (o, e) => Process.Start("explorer.exe", AppPaths.LogDir);
            download.Click += (o, e) =>
            {
                var script = Path.Combine(AppPaths.InstallDir, "scripts", "download-model.ps1");
                if (!File.Exists(script)) { MessageBox.Show("Skrip tidak ditemukan: " + script); return; }
                Process.Start(new ProcessStartInfo("powershell.exe", "-ExecutionPolicy Bypass -NoExit -File \"" + script + "\"") { UseShellExecute = true });
            };
            restart.Click += async (o, e) =>
            {
                try
                {
                    using (var http = new System.Net.Http.HttpClient())
                    {
                        var ep = AppPaths.ReadJson<EngineEndpoint>(AppPaths.EngineEndpointFile);
                        if (ep != null)
                        {
                            var req = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, "http://127.0.0.1:" + ep.Port + "/v1/shutdown");
                            req.Headers.Add("X-MasterBW-Token", ep.Token);
                            await http.SendAsync(req);
                        }
                    }
                }
                catch { }
                status.Text = "Mesin AI dihentikan; akan menyala lagi saat perintah berikutnya.";
            };
            foreach (var b in new[] { save, download, models, logs, restart }) buttons.Controls.Add(b);
            Controls.Add(grid); Controls.Add(status); Controls.Add(buttons);
            Load += async (o, e) =>
            {
                var st = await ai.Client.GetStatusAsync(CancellationToken.None);
                status.Text = st.Ok
                    ? $"Mesin AI aktif • model: {st.ModelName} ({(st.ModelLoaded ? "dimuat" : "belum dimuat")}) • konteks {st.ContextSize} • suara: {(st.SpeechAvailable ? "siap" : "model suara belum ada")}" + (string.IsNullOrEmpty(st.Error) ? "" : "\n⚠ " + st.Error)
                    : "Mesin AI belum berjalan (akan dinyalakan otomatis saat dipakai). " + st.Error;
            };
        }
    }

    /// <summary>Palet perintah cepat (Ctrl+Alt+Space): ketik untuk mencari perintah, Enter untuk menjalankan atau kirim ke AI.</summary>
    public sealed class CommandPaletteForm : Form
    {
        public sealed class Command
        {
            public string Label;
            public string Hint;
            public Action Run;
            public override string ToString() => Label + (string.IsNullOrEmpty(Hint) ? "" : "   — " + Hint);
        }

        private readonly TextBox _input = new TextBox { Dock = DockStyle.Top, Font = new Font("Segoe UI", 13f), BorderStyle = BorderStyle.FixedSingle };
        private readonly ListBox _list = new ListBox { Dock = DockStyle.Fill, Font = Theme.Body, BorderStyle = BorderStyle.None, IntegralHeight = false };
        private readonly List<Command> _all;
        private readonly Action<string> _freeText;

        public CommandPaletteForm(List<Command> commands, Action<string> freeText)
        {
            _all = commands; _freeText = freeText;
            FormBorderStyle = FormBorderStyle.FixedToolWindow; Text = "MASTER BW – ketik perintah (Enter = jalankan, Ctrl+Enter = kirim ke AI)";
            Width = 640; Height = 420; StartPosition = FormStartPosition.CenterScreen; KeyPreview = true; TopMost = true; ShowInTaskbar = false;
            Controls.Add(_list); Controls.Add(_input);
            _input.TextChanged += (o, e) => Filter();
            KeyDown += (o, e) =>
            {
                if (e.KeyCode == Keys.Escape) Close();
                else if (e.KeyCode == Keys.Down && _list.Items.Count > 0) { _list.SelectedIndex = Math.Min(_list.Items.Count - 1, _list.SelectedIndex + 1); e.Handled = true; }
                else if (e.KeyCode == Keys.Up && _list.Items.Count > 0) { _list.SelectedIndex = Math.Max(0, _list.SelectedIndex - 1); e.Handled = true; }
                else if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Execute(e.Control); }
            };
            _list.DoubleClick += (o, e) => Execute(false);
            Filter();
        }

        private void Filter()
        {
            var q = _input.Text.Trim().ToLowerInvariant();
            _list.BeginUpdate();
            _list.Items.Clear();
            var words = q.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var c in _all.Where(c => words.All(w => c.ToString().ToLowerInvariant().Contains(w))).Take(60)) _list.Items.Add(c);
            if (q.Length > 0) _list.Items.Add(new Command { Label = "🤖 Kirim ke AI: \"" + _input.Text.Trim() + "\"", Run = null });
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            _list.EndUpdate();
        }

        private void Execute(bool forceAi)
        {
            var text = _input.Text.Trim();
            var cmd = _list.SelectedItem as Command;
            Close();
            if (forceAi || cmd == null || cmd.Run == null) { if (text.Length > 0) _freeText(text); return; }
            cmd.Run();
        }
    }
}
