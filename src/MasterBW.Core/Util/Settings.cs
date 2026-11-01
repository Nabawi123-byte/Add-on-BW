using System;
using MasterBW.Core.Ai;
using Newtonsoft.Json;

namespace MasterBW.Core.Util
{
    /// <summary>Pengaturan pengguna. Disimpan di %LOCALAPPDATA%\MasterBW\settings.json.</summary>
    public sealed class MasterBWSettings
    {
        // ---- Mesin AI ----
        [JsonProperty] public string ModelFile { get; set; } = "Qwen3-4B-Q4_K_M.gguf";
        [JsonProperty] public string SpeechModelFile { get; set; } = "ggml-base.bin";
        [JsonProperty] public int ContextSize { get; set; } = 8192;
        /// <summary>0 = otomatis (jumlah core fisik).</summary>
        [JsonProperty] public int Threads { get; set; } = 0;
        [JsonProperty] public int EnginePort { get; set; } = 47821;
        /// <summary>Model dilepas dari RAM setelah sekian menit tidak dipakai (0 = tidak pernah).</summary>
        [JsonProperty] public int IdleUnloadMinutes { get; set; } = 20;
        [JsonProperty] public ReasoningMode DefaultReasoning { get; set; } = ReasoningMode.Seimbang;
        [JsonProperty] public int ThinkBudgetSeimbang { get; set; } = 384;
        [JsonProperty] public int ThinkBudgetMendalam { get; set; } = 1536;
        [JsonProperty] public float Temperature { get; set; } = 0.4f;

        // ---- Perilaku ----
        /// <summary>Selalu tampilkan pratinjau dan minta konfirmasi sebelum dokumen diubah.</summary>
        [JsonProperty] public bool ConfirmBeforeExecute { get; set; } = true;
        [JsonProperty] public string PaletteHotkey { get; set; } = "Ctrl+Alt+Space";
        [JsonProperty] public bool CacheCellFunctions { get; set; } = true;

        // ---- Identitas (untuk surat, kop, tanda tangan) ----
        [JsonProperty] public string UserName { get; set; } = "";
        [JsonProperty] public string UserTitle { get; set; } = "";
        [JsonProperty] public string UserNip { get; set; } = "";
        [JsonProperty] public string Institution { get; set; } = "";
        [JsonProperty] public string InstitutionAddress { get; set; } = "";
        [JsonProperty] public string City { get; set; } = "";
        [JsonProperty] public string LogoPath { get; set; } = "";

        // ---- Akademik ----
        [JsonProperty] public string CitationStyle { get; set; } = "APA7";
        [JsonProperty] public string AcademicPreset { get; set; } = "Skripsi";

        public static MasterBWSettings Load()
        {
            return AppPaths.ReadJson<MasterBWSettings>(AppPaths.SettingsFile) ?? new MasterBWSettings();
        }

        public void Save()
        {
            try { AppPaths.WriteJson(AppPaths.SettingsFile, this); }
            catch (Exception ex) { Log.Error("Gagal menyimpan pengaturan", ex); }
        }

        public int ThinkBudgetFor(ReasoningMode mode)
        {
            switch (mode)
            {
                case ReasoningMode.Mendalam: return ThinkBudgetMendalam;
                case ReasoningMode.Seimbang: return ThinkBudgetSeimbang;
                default: return 0;
            }
        }
    }
}
