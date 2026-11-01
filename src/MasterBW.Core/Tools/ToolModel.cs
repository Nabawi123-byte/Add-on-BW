using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MasterBW.Core.Tools
{
    public enum OfficeApp { Word, Excel }

    public sealed class ToolParam
    {
        public string Name { get; set; }
        /// <summary>string | number | integer | boolean | array | object</summary>
        public string Type { get; set; }
        public string Description { get; set; }
        public bool Required { get; set; }
        public string[] Enum { get; set; }
        public string Default { get; set; }
    }

    /// <summary>Definisi satu aksi yang boleh dipanggil AI.</summary>
    public sealed class ToolDef
    {
        public string Name { get; set; }
        public OfficeApp App { get; set; }
        /// <summary>Label bahasa Indonesia untuk pratinjau/palet perintah.</summary>
        public string Label { get; set; }
        public string Category { get; set; }
        public string Description { get; set; }
        public List<ToolParam> Params { get; } = new List<ToolParam>();
        /// <summary>Kata kunci (Indonesia &amp; Inggris) untuk memilih aksi yang relevan.</summary>
        public List<string> Keywords { get; } = new List<string>();
        /// <summary>Aksi hanya membaca (tidak mengubah dokumen), tidak perlu konfirmasi.</summary>
        public bool ReadOnly { get; set; }
        /// <summary>Aksi berisiko menghapus data — diberi peringatan merah di pratinjau.</summary>
        public bool Destructive { get; set; }
        /// <summary>Aksi memanggil AI lagi saat dieksekusi (misalnya menulis paragraf).</summary>
        public bool UsesAi { get; set; }

        // ---- builder fluent ----
        public ToolDef P(string name, string type, string description, bool required = false, string[] enumValues = null, string def = null)
        {
            Params.Add(new ToolParam { Name = name, Type = type, Description = description, Required = required, Enum = enumValues, Default = def });
            return this;
        }

        public ToolDef K(params string[] keywords)
        {
            Keywords.AddRange(keywords);
            return this;
        }

        public ToolDef Ro() { ReadOnly = true; return this; }
        public ToolDef Danger() { Destructive = true; return this; }
        public ToolDef Ai() { UsesAi = true; return this; }

        /// <summary>Deskripsi ringkas untuk prompt model (hemat token).</summary>
        public string ToPromptLine()
        {
            var sb = new StringBuilder();
            sb.Append("- ").Append(Name).Append(": ").Append(Description);
            if (Params.Count > 0)
            {
                sb.Append(" | args: ");
                sb.Append(string.Join("; ", Params.Select(p =>
                {
                    var s = p.Name + (p.Required ? "*" : "") + " (" + p.Type;
                    if (p.Enum != null && p.Enum.Length > 0) s += ": " + string.Join("/", p.Enum);
                    s += ") " + p.Description;
                    return s;
                })));
            }
            return sb.ToString();
        }
    }

    /// <summary>Rencana aksi yang dihasilkan AI. Ditampilkan sebagai pratinjau sebelum dijalankan.</summary>
    public sealed class ActionPlan
    {
        [JsonProperty("message")] public string Message { get; set; } = "";
        [JsonProperty("steps")] public List<ActionStep> Steps { get; set; } = new List<ActionStep>();

        [JsonIgnore] public bool IsEmpty => Steps == null || Steps.Count == 0;
        [JsonIgnore] public string RawText { get; set; }
        [JsonIgnore] public string Source { get; set; } = "ai";

        public static ActionPlan Single(string action, object args, string message = "")
        {
            return new ActionPlan
            {
                Message = message,
                Source = "aturan",
                Steps = new List<ActionStep> { new ActionStep { Action = action, Args = args == null ? new JObject() : JObject.FromObject(args) } }
            };
        }
    }

    public sealed class ActionStep
    {
        [JsonProperty("action")] public string Action { get; set; }
        [JsonProperty("args")] public JObject Args { get; set; } = new JObject();

        /// <summary>Dicentang/tidak oleh pengguna di dialog pratinjau.</summary>
        [JsonIgnore] public bool Enabled { get; set; } = true;
        [JsonIgnore] public List<string> Warnings { get; } = new List<string>();

        public string Str(string name, string def = null)
        {
            var t = Args?[name];
            if (t == null || t.Type == JTokenType.Null) return def;
            if (t.Type == JTokenType.String) return (string)t;
            return t.ToString(Formatting.None);
        }

        public double? Num(string name)
        {
            var t = Args?[name];
            if (t == null || t.Type == JTokenType.Null) return null;
            if (t.Type == JTokenType.Integer || t.Type == JTokenType.Float) return (double)t;
            if (Indonesia.IndoNumber.TryParse((string)t, out var v)) return v;
            return null;
        }

        public double NumOr(string name, double def) => Num(name) ?? def;

        public int IntOr(string name, int def)
        {
            var v = Num(name);
            return v.HasValue ? (int)Math.Round(v.Value) : def;
        }

        public bool Bool(string name, bool def = false)
        {
            var t = Args?[name];
            if (t == null || t.Type == JTokenType.Null) return def;
            if (t.Type == JTokenType.Boolean) return (bool)t;
            var s = t.ToString().Trim().ToLowerInvariant();
            if (s == "true" || s == "ya" || s == "yes" || s == "1") return true;
            if (s == "false" || s == "tidak" || s == "no" || s == "0") return false;
            return def;
        }

        public List<string> StrList(string name)
        {
            var t = Args?[name];
            if (t == null || t.Type == JTokenType.Null) return new List<string>();
            if (t.Type == JTokenType.Array) return t.Select(x => x.Type == JTokenType.String ? (string)x : x.ToString(Formatting.None)).ToList();
            return ((string)t).Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
        }

        /// <summary>Array 2D (baris × kolom) dari argumen, mis. data tabel.</summary>
        public List<List<string>> Grid(string name)
        {
            var result = new List<List<string>>();
            var t = Args?[name] as JArray;
            if (t == null) return result;
            foreach (var row in t)
            {
                if (row is JArray arr) result.Add(arr.Select(c => c.Type == JTokenType.Null ? "" : c.Type == JTokenType.String ? (string)c : c.ToString(Formatting.None)).ToList());
                else result.Add(new List<string> { row.Type == JTokenType.String ? (string)row : row.ToString(Formatting.None) });
            }
            return result;
        }

        public override string ToString() => Action + " " + (Args?.ToString(Formatting.None) ?? "{}");
    }
}
