using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MasterBW.Core.Ai;
using MasterBW.Core.Tools;
using MasterBW.Core.Util;
using MasterBW.Office.Common;

namespace MasterBW.Office.UI
{
    /// <summary>Panel samping MASTER BW: chat dengan AI, tombol cepat, input suara, dan riwayat aksi.</summary>
    public sealed class ChatPane : UserControl
    {
        private readonly AddInController _ctl;
        private readonly RichTextBox _log = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = Theme.Back, Font = Theme.Body, DetectUrls = true };
        private readonly TextBox _input = new TextBox { Dock = DockStyle.Fill, Multiline = true, Font = Theme.Body, ScrollBars = ScrollBars.Vertical };
        private readonly ComboBox _mode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140, Font = Theme.Small };
        private readonly Label _status = new Label { Dock = DockStyle.Bottom, Height = 34, Font = Theme.Small, ForeColor = Theme.Muted, Padding = new Padding(6, 2, 6, 2) };
        private readonly Button _send = Theme.Button("Kirim", true);
        private readonly Button _stop = Theme.Button("■ Stop");
        private readonly Button _mic = Theme.Button("Suara");
        private readonly Button _undo = Theme.Button("↶ Batalkan");
        private readonly List<ChatMessage> _history = new List<ChatMessage>();
        private CancellationTokenSource _cts;
        private readonly VoiceRecorder _recorder = new VoiceRecorder();
        public bool Busy { get; private set; }

        public ChatPane(AddInController controller)
        {
            _ctl = controller;
            Dock = DockStyle.Fill;
            BackColor = Theme.Back;
            Font = Theme.Body;

            var header = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = Theme.Primary, Padding = new Padding(10, 6, 10, 6) };
            var title = new Label { Text = "MASTER BW", ForeColor = Color.White, Font = Theme.Title, AutoSize = true, Location = new Point(10, 6) };
            var sub = new Label { Text = "Asisten AI offline • " + (controller.App == OfficeApp.Word ? "Word" : "Excel"), ForeColor = Color.FromArgb(210, 225, 240), Font = Theme.Small, AutoSize = true, Location = new Point(12, 34) };
            header.Controls.Add(title); header.Controls.Add(sub);

            _mode.Items.AddRange(new object[] { "⚡ Cepat", "⚖ Seimbang", "🧠 Mendalam" });
            _mode.SelectedIndex = (int)controller.Ai.Settings.DefaultReasoning;
            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34, Padding = new Padding(4, 4, 4, 0), BackColor = Theme.PrimaryLight };
            toolbar.Controls.Add(new Label { Text = "Nalar:", AutoSize = true, Padding = new Padding(0, 5, 0, 0), Font = Theme.Small });
            toolbar.Controls.Add(_mode);
            _undo.Enabled = false;
            _undo.Click += (o, e) => { AppendSystem(_ctl.Executor.UndoLastPlan()); _undo.Enabled = _ctl.Executor.CanUndoLastPlan; };
            toolbar.Controls.Add(_undo);

            var quick = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, MaximumSize = new Size(2000, 0), Padding = new Padding(4), WrapContents = true };
            foreach (var (label, action) in controller.QuickActions())
            {
                var b = Theme.Button(label);
                b.Font = Theme.Small;
                b.Margin = new Padding(2);
                b.Click += (o, e) => action();
                quick.Controls.Add(b);
            }

            var inputPanel = new Panel { Dock = DockStyle.Bottom, Height = 96, Padding = new Padding(6) };
            var btns = new FlowLayoutPanel { Dock = DockStyle.Right, Width = 82, FlowDirection = FlowDirection.TopDown };
            _send.Width = 76; _stop.Width = 76; _mic.Width = 76;
            _stop.Enabled = false;
            btns.Controls.Add(_send); btns.Controls.Add(_mic); btns.Controls.Add(_stop);
            inputPanel.Controls.Add(_input); inputPanel.Controls.Add(btns);
            _send.Click += (o, e) => SubmitInput();
            _stop.Click += (o, e) => _cts?.Cancel();
            _mic.Click += async (o, e) => await ToggleMic();
            _input.KeyDown += (o, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift) { e.SuppressKeyPress = true; SubmitInput(); }
            };

            Controls.Add(_log);
            Controls.Add(quick);
            Controls.Add(toolbar);
            Controls.Add(header);
            Controls.Add(inputPanel);
            Controls.Add(_status);

            AppendSystem("Halo! Ketik perintah dalam bahasa Indonesia, misalnya:\n" + string.Join("\n", controller.Examples().Select(x => "• " + x)) +
                         "\nSetiap perubahan ditampilkan dulu untuk Anda setujui. Pintasan: " + controller.Ai.Settings.PaletteHotkey + " untuk palet perintah.");
            _ctl.Ai.Status = s => SetStatus(s);
        }

        public ReasoningMode Mode => (ReasoningMode)Math.Max(0, _mode.SelectedIndex);

        public void Prefill(string text)
        {
            _input.Text = text;
            _input.Focus();
            _input.SelectionStart = _input.Text.Length;
        }

        private void SubmitInput()
        {
            var text = _input.Text.Trim();
            if (text.Length == 0 || Busy) return;
            _input.Clear();
            _ = SubmitAsync(text);
        }

        // ------------------------------------------------------------------ alur utama
        public async Task SubmitAsync(string text)
        {
            if (Busy) { AppendSystem("Masih memproses perintah sebelumnya. Tekan Stop untuk membatalkan."); return; }
            AppendUser(text);
            await RunGuarded(async ct =>
            {
                var ctx = _ctl.Executor.GetContext();
                SetStatus("Memahami perintah…");
                int thinkChars = 0;
                var plan = await _ctl.Ai.PlanAsync(_ctl.App, text, ctx, Mode, _history, ct, chunk =>
                {
                    if (chunk.Kind == "think") { thinkChars += chunk.Delta?.Length ?? 0; SetStatus("Berpikir… (" + thinkChars + " karakter)"); }
                    else SetStatus("Menyusun rencana aksi…");
                });
                _history.Add(ChatMessage.User(text));
                _history.Add(ChatMessage.Assistant(string.IsNullOrWhiteSpace(plan.Message) ? "(aksi)" : plan.Message));
                if (plan.IsEmpty)
                {
                    AppendAssistant(string.IsNullOrWhiteSpace(plan.Message) ? "Maaf, saya belum memahami perintah itu. Coba jelaskan lebih spesifik." : plan.Message);
                    return;
                }
                if (!string.IsNullOrWhiteSpace(plan.Message)) AppendAssistant(plan.Message);
                await ExecutePlanCore(plan, text, ct);
            });
        }

        /// <summary>Jalankan rencana yang sudah jadi (dari tombol ribbon/menu), tetap lewat pratinjau.</summary>
        public Task RunPlanAsync(ActionPlan plan, string title)
        {
            if (Busy) { AppendSystem("Masih memproses perintah sebelumnya."); return Task.CompletedTask; }
            AppendUser(title);
            return RunGuarded(ct => ExecutePlanCore(plan, title, ct));
        }

        private async Task ExecutePlanCore(ActionPlan plan, string title, CancellationToken ct)
        {
            bool allReadOnly = plan.Steps.All(s => ToolCatalog.Find(_ctl.App, s.Action)?.ReadOnly == true);
            bool anyDanger = plan.Steps.Any(s => ToolCatalog.Find(_ctl.App, s.Action)?.Destructive == true || s.Warnings.Count > 0);
            if ((_ctl.Ai.Settings.ConfirmBeforeExecute && !allReadOnly) || anyDanger)
            {
                using (var preview = new PlanPreviewForm(plan, _ctl.App))
                {
                    if (preview.ShowDialog(this) != DialogResult.OK) { AppendSystem("Dibatalkan. Dokumen tidak diubah."); return; }
                }
            }
            var session = new ExecutionSession
            {
                Ai = _ctl.Ai, Mode = Mode, Ct = ct,
                Report = s => SetStatus(s),
                ShowText = (t, b) => _ctl.ShowText(t, b),
            };
            _ctl.Executor.BeginPlan(title);
            try
            {
                int i = 0;
                foreach (var step in plan.Steps.Where(s => s.Enabled))
                {
                    ct.ThrowIfCancellationRequested();
                    i++;
                    var def = ToolCatalog.Find(_ctl.App, step.Action);
                    SetStatus("Menjalankan: " + (def?.Label ?? step.Action) + "…");
                    try
                    {
                        var result = await _ctl.Executor.ExecuteAsync(step, session);
                        AppendResult(true, (def?.Label ?? step.Action) + ": " + result);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        Log.Error("Aksi " + step.Action + " gagal", ex);
                        AppendResult(false, (def?.Label ?? step.Action) + ": " + Friendly(ex));
                    }
                }
            }
            finally
            {
                _ctl.Executor.EndPlan();
                _undo.Enabled = _ctl.Executor.CanUndoLastPlan;
            }
        }

        private async Task RunGuarded(Func<CancellationToken, Task> work)
        {
            Busy = true;
            _cts = new CancellationTokenSource();
            _send.Enabled = false; _stop.Enabled = true;
            try { await work(_cts.Token); }
            catch (OperationCanceledException) { AppendSystem("Dihentikan."); }
            catch (Exception ex) { Log.Error("Chat gagal", ex); AppendResult(false, Friendly(ex)); }
            finally
            {
                Busy = false;
                _send.Enabled = true; _stop.Enabled = false;
                SetStatus("Siap.");
            }
        }

        private static string Friendly(Exception ex)
        {
            var e = ex is AggregateException ae && ae.InnerException != null ? ae.InnerException : ex;
            if (e is System.Runtime.InteropServices.COMException) return "Office menolak perintah (" + e.Message + "). Pastikan tidak sedang mengedit sel/teks dan dokumen tidak terkunci.";
            if (e is System.Net.Http.HttpRequestException) return "Mesin AI tidak dapat dihubungi. Buka Pengaturan → Mulai ulang mesin AI. (" + e.Message + ")";
            return e.Message;
        }

        // ------------------------------------------------------------------ suara
        private async Task ToggleMic()
        {
            try
            {
                if (!_recorder.IsRecording)
                {
                    _recorder.Start();
                    _mic.Text = "Selesai";
                    _mic.BackColor = Color.MistyRose;
                    SetStatus("Merekam… klik lagi bila selesai bicara.");
                    return;
                }
                var wav = _recorder.Stop();
                _mic.Text = "Suara"; _mic.BackColor = Color.White;
                SetStatus("Mengubah suara menjadi teks (offline)…");
                await _ctl.Ai.EnsureReadyAsync(CancellationToken.None);
                var text = await _ctl.Ai.Client.TranscribeAsync(wav, CancellationToken.None);
                Prefill((_input.Text + " " + text).Trim());
                SetStatus("Periksa teks lalu tekan Enter.");
            }
            catch (Exception ex)
            {
                _mic.Text = "Suara"; _mic.BackColor = Color.White;
                AppendResult(false, "Input suara: " + ex.Message);
            }
        }

        // ------------------------------------------------------------------ tampilan
        private void Ui(Action a)
        {
            if (IsDisposed) return;
            if (InvokeRequired) BeginInvoke(a); else a();
        }

        public void SetStatus(string text) => Ui(() => _status.Text = text);

        private void Append(string who, string text, Color whoColor, Color textColor)
        {
            Ui(() =>
            {
                _log.SelectionStart = _log.TextLength;
                _log.SelectionFont = Theme.Bold;
                _log.SelectionColor = whoColor;
                _log.AppendText((_log.TextLength > 0 ? "\n" : "") + who + "\n");
                _log.SelectionFont = Theme.Body;
                _log.SelectionColor = textColor;
                _log.AppendText(text.Trim() + "\n");
                _log.ScrollToCaret();
            });
        }

        public void AppendUser(string text) => Append("Anda", text, Theme.Primary, Color.Black);
        public void AppendAssistant(string text) => Append("MASTER BW", text, Theme.Ok, Color.FromArgb(30, 30, 30));
        public void AppendSystem(string text) => Append("ℹ", text, Theme.Muted, Theme.Muted);
        public void AppendResult(bool ok, string text) => Append(ok ? "✔ Selesai" : "✖ Gagal", text, ok ? Theme.Ok : Theme.Error, ok ? Color.FromArgb(30, 30, 30) : Theme.Error);
    }
}
