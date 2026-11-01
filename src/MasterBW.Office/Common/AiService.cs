using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MasterBW.Core.Ai;
using MasterBW.Core.Tools;
using MasterBW.Core.Util;

namespace MasterBW.Office.Common
{
    /// <summary>Semua pemanggilan AI dari add-in lewat kelas ini.</summary>
    public sealed class AiService
    {
        public EngineClient Client { get; } = new EngineClient();
        public MasterBWSettings Settings { get; private set; } = MasterBWSettings.Load();
        public Action<string> Status { get; set; }

        public void ReloadSettings() => Settings = MasterBWSettings.Load();

        public async Task EnsureReadyAsync(CancellationToken ct)
        {
            var st = await EngineLauncher.EnsureRunningAsync(Client, Status, ct);
            if (!st.Ok) throw new InvalidOperationException(st.Error ?? "Mesin AI belum siap.");
            if (!st.ModelLoaded && !string.IsNullOrEmpty(st.Error)) throw new InvalidOperationException(st.Error);
            if (!st.ModelLoaded) Status?.Invoke("Memuat model AI ke memori (sekali saja, ±10–30 detik)…");
        }

        public async Task<GenerateResponse> GenerateAsync(List<ChatMessage> messages, ReasoningMode mode, int maxTokens, CancellationToken ct,
                                                         Action<StreamChunk> onChunk = null, string grammar = null, float? temperature = null)
        {
            await EnsureReadyAsync(ct);
            var req = new GenerateRequest
            {
                Messages = messages,
                MaxTokens = maxTokens,
                Reasoning = mode,
                ThinkBudget = Settings.ThinkBudgetFor(mode),
                Grammar = grammar,
                Temperature = temperature ?? Settings.Temperature,
            };
            return await Client.GenerateAsync(req, ct, onChunk);
        }

        public async Task<string> CompleteAsync(List<ChatMessage> messages, ReasoningMode mode, int maxTokens, CancellationToken ct, Action<StreamChunk> onChunk = null)
        {
            var r = await GenerateAsync(messages, mode, maxTokens, ct, onChunk);
            return CleanOutput(r.Text);
        }

        /// <summary>Ubah perintah bahasa Indonesia menjadi rencana aksi yang tervalidasi.</summary>
        public async Task<ActionPlan> PlanAsync(OfficeApp app, string request, OfficeContext ctx, ReasoningMode mode, IList<ChatMessage> history,
                                                CancellationToken ct, Action<StreamChunk> onChunk = null)
        {
            var tools = ToolCatalog.Select(app, request, 10);
            var messages = Prompts.BuildPlanMessages(app, request, ctx, tools, history);
            var grammar = PlanGrammar.Build(tools);
            var r = await GenerateAsync(messages, mode, 1400, ct, onChunk, grammar);
            var plan = JsonRepair.ParsePlan(r.Text);
            if (plan.IsEmpty && string.IsNullOrWhiteSpace(plan.Message)) plan.Message = r.Text;
            ToolCatalog.Validate(app, plan);
            return plan;
        }

        /// <summary>Bersihkan pembuka khas AI dan blok kode pembungkus.</summary>
        public static string CleanOutput(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            var t = ThinkFilter.StripThinking(text).Trim();
            if (t.StartsWith("```"))
            {
                var nl = t.IndexOf('\n');
                if (nl > 0) t = t.Substring(nl + 1);
                if (t.EndsWith("```")) t = t.Substring(0, t.Length - 3);
            }
            foreach (var opener in new[] { "Berikut adalah", "Berikut ini", "Tentu,", "Tentu!", "Baik,", "Here is", "Sure," })
            {
                if (t.StartsWith(opener, StringComparison.OrdinalIgnoreCase))
                {
                    var nl = t.IndexOf('\n');
                    if (nl > 0 && nl < 160 && t.Substring(0, nl).TrimEnd().EndsWith(":")) t = t.Substring(nl + 1);
                    break;
                }
            }
            return t.Trim().Trim('"').Trim();
        }

        /// <summary>Potong teks panjang per paragraf agar muat di konteks model.</summary>
        public static List<string> Chunk(string text, int maxChars = 5000)
        {
            var chunks = new List<string>();
            var sb = new StringBuilder();
            foreach (var para in (text ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None))
            {
                if (sb.Length + para.Length + 1 > maxChars && sb.Length > 0)
                {
                    chunks.Add(sb.ToString());
                    sb.Clear();
                }
                if (para.Length > maxChars)
                {
                    for (int i = 0; i < para.Length; i += maxChars) chunks.Add(para.Substring(i, Math.Min(maxChars, para.Length - i)));
                    continue;
                }
                sb.AppendLine(para);
            }
            if (sb.Length > 0) chunks.Add(sb.ToString());
            return chunks;
        }

        /// <summary>Olah teks (ringkas/parafrase/terjemah…) termasuk teks panjang secara bertahap.</summary>
        public async Task<string> TransformAsync(string task, string text, ReasoningMode mode, CancellationToken ct, Action<string> progress = null)
        {
            var chunks = Chunk(text);
            if (chunks.Count == 1)
                return await CompleteAsync(Prompts.BuildTextTask(task, text), mode, Math.Max(512, text.Length / 2), ct);
            var outputs = new List<string>();
            for (int i = 0; i < chunks.Count; i++)
            {
                progress?.Invoke($"Memproses bagian {i + 1}/{chunks.Count}…");
                outputs.Add(await CompleteAsync(Prompts.BuildTextTask(task, chunks[i]), mode, Math.Max(512, chunks[i].Length / 2), ct));
            }
            if (string.Equals(task, "ringkas", StringComparison.OrdinalIgnoreCase) && outputs.Count > 1)
            {
                progress?.Invoke("Menggabungkan ringkasan…");
                return await CompleteAsync(Prompts.BuildTextTask("ringkas", string.Join("\n\n", outputs)), mode, 1200, ct);
            }
            return string.Join("\n\n", outputs);
        }
    }

    /// <summary>Konteks satu eksekusi rencana aksi.</summary>
    public sealed class ExecutionSession
    {
        public AiService Ai { get; set; }
        public MasterBWSettings Settings => Ai.Settings;
        public ReasoningMode Mode { get; set; } = ReasoningMode.Seimbang;
        public CancellationToken Ct { get; set; }
        /// <summary>Laporan progres ke panel chat.</summary>
        public Action<string> Report { get; set; } = _ => { };
        /// <summary>Tampilkan teks panjang (penjelasan, insight) ke pengguna.</summary>
        public Action<string, string> ShowText { get; set; } = (t, b) => { };
    }

    public interface IActionExecutor
    {
        OfficeApp App { get; }
        OfficeContext GetContext();
        /// <summary>Jalankan satu langkah. Mengembalikan ringkasan hasil untuk ditampilkan.</summary>
        Task<string> ExecuteAsync(ActionStep step, ExecutionSession session);
        /// <summary>Dipanggil sebelum dan sesudah seluruh rencana (untuk grup undo/snapshot).</summary>
        void BeginPlan(string title);
        void EndPlan();
        bool CanUndoLastPlan { get; }
        string UndoLastPlan();
    }
}
