using System.Collections.Generic;
using System.Text;

namespace MasterBW.Core.Stats
{
    /// <summary>Tabel hasil analisis yang siap ditulis ke Excel atau Word.</summary>
    public sealed class StatTable
    {
        public string Title { get; set; }
        public List<string> Headers { get; } = new List<string>();
        public List<object[]> Rows { get; } = new List<object[]>();
        public List<string> Notes { get; } = new List<string>();
        /// <summary>Format angka per kolom (indeks kolom → format Excel), opsional.</summary>
        public Dictionary<int, string> ColumnFormats { get; } = new Dictionary<int, string>();

        public StatTable(string title, params string[] headers)
        {
            Title = title;
            Headers.AddRange(headers);
        }

        public StatTable Row(params object[] values)
        {
            Rows.Add(values);
            return this;
        }

        public StatTable Note(string note)
        {
            Notes.Add(note);
            return this;
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.AppendLine(Title);
            sb.AppendLine(string.Join(" | ", Headers));
            foreach (var r in Rows) sb.AppendLine(string.Join(" | ", r));
            foreach (var n in Notes) sb.AppendLine("* " + n);
            return sb.ToString();
        }
    }

    /// <summary>Kumpulan tabel + interpretasi bahasa Indonesia.</summary>
    public sealed class AnalysisResult
    {
        public string Title { get; set; }
        public List<StatTable> Tables { get; } = new List<StatTable>();
        public List<string> Interpretation { get; } = new List<string>();

        public AnalysisResult(string title) { Title = title; }

        public string ToText()
        {
            var sb = new StringBuilder();
            sb.AppendLine(Title);
            foreach (var t in Tables) sb.AppendLine(t.ToString());
            if (Interpretation.Count > 0)
            {
                sb.AppendLine("Interpretasi:");
                foreach (var i in Interpretation) sb.AppendLine("- " + i);
            }
            return sb.ToString();
        }
    }
}
