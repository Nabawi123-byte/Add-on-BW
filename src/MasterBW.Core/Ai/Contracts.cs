using System.Collections.Generic;
using Newtonsoft.Json;

namespace MasterBW.Core.Ai
{
    /// <summary>Mode nalar AI. Makin dalam, makin pintar tetapi makin lambat di CPU.</summary>
    public enum ReasoningMode
    {
        /// <summary>Jawaban langsung tanpa proses berpikir (paling cepat).</summary>
        Cepat = 0,
        /// <summary>Berpikir singkat untuk perintah yang butuh perencanaan.</summary>
        Seimbang = 1,
        /// <summary>Berpikir panjang (high reasoning) untuk analisis dan dokumen kompleks.</summary>
        Mendalam = 2
    }

    public sealed class ChatMessage
    {
        [JsonProperty("role")] public string Role { get; set; }
        [JsonProperty("content")] public string Content { get; set; }

        public ChatMessage() { }
        public ChatMessage(string role, string content) { Role = role; Content = content; }

        public static ChatMessage System(string text) => new ChatMessage("system", text);
        public static ChatMessage User(string text) => new ChatMessage("user", text);
        public static ChatMessage Assistant(string text) => new ChatMessage("assistant", text);
    }

    public sealed class GenerateRequest
    {
        [JsonProperty("messages")] public List<ChatMessage> Messages { get; set; } = new List<ChatMessage>();
        [JsonProperty("maxTokens")] public int MaxTokens { get; set; } = 1024;
        [JsonProperty("temperature")] public float Temperature { get; set; } = 0.4f;
        [JsonProperty("topP")] public float TopP { get; set; } = 0.9f;
        /// <summary>Grammar GBNF opsional untuk memaksa format keluaran (misalnya JSON rencana aksi).</summary>
        [JsonProperty("grammar")] public string Grammar { get; set; }
        [JsonProperty("reasoning")] public ReasoningMode Reasoning { get; set; } = ReasoningMode.Cepat;
        /// <summary>Batas token khusus untuk fase berpikir (agar CPU tidak "melamun" terlalu lama).</summary>
        [JsonProperty("thinkBudget")] public int ThinkBudget { get; set; } = 0;
        [JsonProperty("stop")] public List<string> Stop { get; set; } = new List<string>();
    }

    public sealed class GenerateResponse
    {
        [JsonProperty("text")] public string Text { get; set; } = "";
        [JsonProperty("reasoning")] public string Reasoning { get; set; } = "";
        [JsonProperty("promptTokens")] public int PromptTokens { get; set; }
        [JsonProperty("cachedTokens")] public int CachedTokens { get; set; }
        [JsonProperty("completionTokens")] public int CompletionTokens { get; set; }
        [JsonProperty("seconds")] public double Seconds { get; set; }
        [JsonProperty("stopReason")] public string StopReason { get; set; } = "";
        [JsonProperty("error")] public string Error { get; set; }
    }

    /// <summary>Potongan streaming yang dikirim engine lewat Server-Sent Events.</summary>
    public sealed class StreamChunk
    {
        /// <summary>"think" = teks penalaran, "text" = teks jawaban, "done" = selesai.</summary>
        [JsonProperty("kind")] public string Kind { get; set; }
        [JsonProperty("delta")] public string Delta { get; set; }
        [JsonProperty("result")] public GenerateResponse Result { get; set; }
    }

    public sealed class TranscribeResponse
    {
        [JsonProperty("text")] public string Text { get; set; } = "";
        [JsonProperty("error")] public string Error { get; set; }
    }

    public sealed class EngineStatus
    {
        [JsonProperty("ok")] public bool Ok { get; set; }
        [JsonProperty("version")] public string Version { get; set; }
        [JsonProperty("modelLoaded")] public bool ModelLoaded { get; set; }
        [JsonProperty("modelName")] public string ModelName { get; set; }
        [JsonProperty("modelPath")] public string ModelPath { get; set; }
        [JsonProperty("contextSize")] public int ContextSize { get; set; }
        [JsonProperty("threads")] public int Threads { get; set; }
        [JsonProperty("speechAvailable")] public bool SpeechAvailable { get; set; }
        [JsonProperty("busy")] public bool Busy { get; set; }
        [JsonProperty("error")] public string Error { get; set; }
    }

    /// <summary>Isi file engine.json yang ditulis engine saat berjalan (port + token akses lokal).</summary>
    public sealed class EngineEndpoint
    {
        [JsonProperty("port")] public int Port { get; set; }
        [JsonProperty("token")] public string Token { get; set; }
        [JsonProperty("pid")] public int Pid { get; set; }
    }
}
