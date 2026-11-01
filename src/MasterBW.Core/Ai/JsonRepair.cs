using System;
using System.Text;
using MasterBW.Core.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MasterBW.Core.Ai
{
    /// <summary>Mengambil JSON dari keluaran model kecil yang kadang "berantakan".</summary>
    public static class JsonRepair
    {
        /// <summary>Ambil objek JSON pertama yang seimbang kurung kurawalnya, abaikan teks/markdown di sekitarnya.</summary>
        public static string ExtractObject(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
            text = ThinkFilter.StripThinking(text);
            int start = text.IndexOf('{');
            while (start >= 0)
            {
                int depth = 0;
                bool inString = false, escape = false;
                for (int i = start; i < text.Length; i++)
                {
                    char c = text[i];
                    if (inString)
                    {
                        if (escape) escape = false;
                        else if (c == '\\') escape = true;
                        else if (c == '"') inString = false;
                        continue;
                    }
                    if (c == '"') inString = true;
                    else if (c == '{') depth++;
                    else if (c == '}')
                    {
                        depth--;
                        if (depth == 0) return text.Substring(start, i - start + 1);
                    }
                }
                // tidak seimbang: tutup paksa
                if (depth > 0) return text.Substring(start) + new string('}', depth);
                start = text.IndexOf('{', start + 1);
            }
            return null;
        }

        /// <summary>Perbaikan ringan: koma berlebih, kutip tunggal, komentar.</summary>
        public static string Repair(string json)
        {
            if (json == null) return null;
            var sb = new StringBuilder(json.Length);
            bool inString = false, escape = false;
            for (int i = 0; i < json.Length; i++)
            {
                char c = json[i];
                if (inString)
                {
                    if (escape) { escape = false; sb.Append(c); continue; }
                    if (c == '\\') { escape = true; sb.Append(c); continue; }
                    if (c == '"') inString = false;
                    if (c == '\n') { sb.Append("\\n"); continue; }
                    if (c == '\t') { sb.Append("\\t"); continue; }
                    if (c == '\r') continue;
                    sb.Append(c);
                    continue;
                }
                if (c == '"') { inString = true; sb.Append(c); continue; }
                if (c == '/' && i + 1 < json.Length && json[i + 1] == '/')
                {
                    while (i < json.Length && json[i] != '\n') i++;
                    continue;
                }
                if (c == ',')
                {
                    int j = i + 1;
                    while (j < json.Length && char.IsWhiteSpace(json[j])) j++;
                    if (j < json.Length && (json[j] == '}' || json[j] == ']')) continue; // trailing comma
                }
                sb.Append(c);
            }
            return sb.ToString();
        }

        public static JObject ParseObject(string text)
        {
            var obj = ExtractObject(text);
            if (obj == null) return null;
            try { return JObject.Parse(obj); }
            catch (JsonException)
            {
                try { return JObject.Parse(Repair(obj)); }
                catch (JsonException) { return null; }
            }
        }

        /// <summary>Parse keluaran model menjadi rencana aksi. Toleran terhadap variasi nama kunci.</summary>
        public static ActionPlan ParsePlan(string text)
        {
            var o = ParseObject(text);
            var plan = new ActionPlan { RawText = text };
            if (o == null)
            {
                plan.Message = ThinkFilter.StripThinking(text ?? "").Trim();
                return plan;
            }
            plan.Message = (string)(o["message"] ?? o["pesan"] ?? o["reply"] ?? "") ?? "";
            var steps = (o["steps"] ?? o["langkah"] ?? o["actions"]) as JArray;
            if (steps == null && o["action"] != null) steps = new JArray(o);
            if (steps != null)
            {
                foreach (var s in steps)
                {
                    if (!(s is JObject so)) continue;
                    var name = (string)(so["action"] ?? so["aksi"] ?? so["tool"] ?? so["name"]);
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var args = (so["args"] ?? so["param"] ?? so["parameters"] ?? so["arguments"]) as JObject;
                    if (args == null)
                    {
                        // argumen ditaruh sejajar dengan "action"
                        args = new JObject();
                        foreach (var prop in so.Properties())
                            if (prop.Name != "action" && prop.Name != "aksi" && prop.Name != "tool" && prop.Name != "name")
                                args[prop.Name] = prop.Value;
                    }
                    plan.Steps.Add(new ActionStep { Action = name.Trim(), Args = args });
                }
            }
            return plan;
        }
    }

    /// <summary>Memisahkan blok &lt;think&gt; (penalaran model) dari jawaban.</summary>
    public static class ThinkFilter
    {
        public static string StripThinking(string text)
        {
            if (string.IsNullOrEmpty(text)) return text ?? "";
            int open = text.IndexOf("<think>", StringComparison.Ordinal);
            int close = text.LastIndexOf("</think>", StringComparison.Ordinal);
            if (close >= 0) return text.Substring(close + 8).TrimStart();
            if (open >= 0) return text.Substring(0, open).TrimEnd();
            return text;
        }
    }
}
