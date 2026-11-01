using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MasterBW.Core.Academic
{
    /// <summary>
    /// Blok penyusun dokumen Word. Dipakai templat bawaan maupun AI (aksi build_document).
    /// Tipe: title, heading, paragraph, center, bullets, numbered, table, keyvalue, signature, letterhead,
    /// toc, page_break, spacer, ai, cover, right.
    /// </summary>
    public sealed class DocBlock
    {
        [JsonProperty("type")] public string Type { get; set; } = "paragraph";
        [JsonProperty("text")] public string Text { get; set; }
        [JsonProperty("level")] public int Level { get; set; } = 1;
        [JsonProperty("items")] public List<string> Items { get; set; }
        [JsonProperty("rows")] public List<List<string>> Rows { get; set; }
        [JsonProperty("bold")] public bool Bold { get; set; }
        /// <summary>Data tambahan, mis. untuk signature: city, date, title, name, id_label, id_number.</summary>
        [JsonProperty("data")] public Dictionary<string, string> Data { get; set; }

        public static DocBlock Of(string type, string text = null, int level = 1) => new DocBlock { Type = type, Text = text, Level = level };

        /// <summary>Parse array blok dari JSON AI dengan toleransi bentuk.</summary>
        public static List<DocBlock> ParseMany(JToken token)
        {
            var list = new List<DocBlock>();
            if (!(token is JArray arr)) return list;
            foreach (var t in arr)
            {
                if (t.Type == JTokenType.String) { list.Add(Of("paragraph", (string)t)); continue; }
                if (!(t is JObject o)) continue;
                var b = new DocBlock
                {
                    Type = ((string)(o["type"] ?? "paragraph")).ToLowerInvariant(),
                    Text = (string)(o["text"] ?? o["content"] ?? o["title"]),
                    Level = (int?)o["level"] ?? 1,
                    Bold = (bool?)o["bold"] ?? false,
                };
                if (o["items"] is JArray items) b.Items = items.Select(x => (string)x).ToList();
                if (o["rows"] is JArray rows)
                    b.Rows = rows.Select(r => r is JArray ra ? ra.Select(c => c.Type == JTokenType.Null ? "" : c.ToString()).ToList() : new List<string> { r.ToString() }).ToList();
                if (o["data"] is JObject data) b.Data = data.Properties().ToDictionary(p => p.Name, p => p.Value.ToString());
                if (b.Type == "h1" || b.Type == "h2" || b.Type == "h3") { b.Level = b.Type[1] - '0'; b.Type = "heading"; }
                if (b.Type == "list" || b.Type == "bullet") b.Type = "bullets";
                list.Add(b);
            }
            return list;
        }
    }

    /// <summary>Preset format akademik/kantor (font, spasi, margin).</summary>
    public sealed class FormatPreset
    {
        public string Name;
        public string Font = "Times New Roman";
        public double Size = 12;
        public double LineSpacing = 2;
        public double MarginTop = 4, MarginLeft = 4, MarginBottom = 3, MarginRight = 3;
        public string Paper = "A4";
        public double FirstLineIndentCm = 1.27;
        public double SpaceAfter = 0;
        public bool Justify = true;
        public string HeadingNumbering = "bab_romawi";

        public static readonly Dictionary<string, FormatPreset> All = new Dictionary<string, FormatPreset>(System.StringComparer.OrdinalIgnoreCase)
        {
            ["Skripsi"] = new FormatPreset { Name = "Skripsi", LineSpacing = 2, MarginTop = 4, MarginLeft = 4, MarginBottom = 3, MarginRight = 3, HeadingNumbering = "bab_romawi" },
            ["Proposal"] = new FormatPreset { Name = "Proposal", LineSpacing = 2, MarginTop = 4, MarginLeft = 4, MarginBottom = 3, MarginRight = 3, HeadingNumbering = "bab_romawi" },
            ["Makalah"] = new FormatPreset { Name = "Makalah", LineSpacing = 1.5, MarginTop = 3, MarginLeft = 4, MarginBottom = 3, MarginRight = 3, HeadingNumbering = "bab_romawi" },
            ["Laporan"] = new FormatPreset { Name = "Laporan", LineSpacing = 1.5, MarginTop = 3, MarginLeft = 3, MarginBottom = 3, MarginRight = 3, HeadingNumbering = "desimal" },
            ["Jurnal"] = new FormatPreset { Name = "Jurnal", Size = 11, LineSpacing = 1, MarginTop = 2.5, MarginLeft = 2.5, MarginBottom = 2.5, MarginRight = 2, SpaceAfter = 6, HeadingNumbering = "none" },
            ["SuratDinas"] = new FormatPreset { Name = "SuratDinas", Font = "Arial", Size = 12, LineSpacing = 1.15, MarginTop = 2, MarginLeft = 2.5, MarginBottom = 2, MarginRight = 2, Paper = "F4", FirstLineIndentCm = 0, SpaceAfter = 6, HeadingNumbering = "none" },
        };
    }
}
