using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace MasterBW.Core.Academic
{
    /// <summary>Data satu sumber pustaka.</summary>
    public sealed class Reference
    {
        [JsonProperty("id")] public string Id { get; set; }
        /// <summary>journal | book | chapter | web | thesis | report | proceeding | regulation</summary>
        [JsonProperty("type")] public string Type { get; set; } = "book";
        [JsonProperty("authors")] public List<string> Authors { get; set; } = new List<string>();
        [JsonProperty("year")] public string Year { get; set; }
        [JsonProperty("title")] public string Title { get; set; }
        [JsonProperty("journal")] public string Journal { get; set; }
        [JsonProperty("volume")] public string Volume { get; set; }
        [JsonProperty("issue")] public string Issue { get; set; }
        [JsonProperty("pages")] public string Pages { get; set; }
        [JsonProperty("publisher")] public string Publisher { get; set; }
        [JsonProperty("city")] public string City { get; set; }
        [JsonProperty("edition")] public string Edition { get; set; }
        [JsonProperty("editors")] public List<string> Editors { get; set; } = new List<string>();
        [JsonProperty("book_title")] public string BookTitle { get; set; }
        [JsonProperty("institution")] public string Institution { get; set; }
        [JsonProperty("degree")] public string Degree { get; set; }
        [JsonProperty("url")] public string Url { get; set; }
        [JsonProperty("doi")] public string Doi { get; set; }
        [JsonProperty("accessed")] public string Accessed { get; set; }

        public string Key => (Id ?? (Authors.FirstOrDefault() ?? "") + Year).Trim();
    }

    /// <summary>Nama penulis yang sudah diurai.</summary>
    public sealed class PersonName
    {
        public string Last;
        public List<string> Given = new List<string>();
        public bool Single => Given.Count == 0;

        /// <summary>Terima "Arikunto, Suharsimi", "Suharsimi Arikunto", "Sugiyono", "Creswell, J. W.", juga lembaga dalam {kurung kurawal}.</summary>
        public static PersonName Parse(string raw)
        {
            var p = new PersonName();
            var s = (raw ?? "").Trim();
            if (s.StartsWith("{") && s.EndsWith("}")) { p.Last = s.Trim('{', '}'); return p; }
            if (s.Contains(","))
            {
                var parts = s.Split(new[] { ',' }, 2);
                p.Last = parts[0].Trim();
                p.Given = parts[1].Split(new[] { ' ', '.' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }
            else
            {
                var parts = s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (parts.Count <= 1) { p.Last = s; return p; }
                p.Last = parts.Last();
                p.Given = parts.Take(parts.Count - 1).Select(x => x.TrimEnd('.')).ToList();
            }
            return p;
        }

        public string Initials(bool spaced = true, bool dots = true)
            => string.Join(spaced ? " " : "", Given.Select(g => g.Substring(0, 1).ToUpperInvariant() + (dots ? "." : "")));

        public string GivenFull => string.Join(" ", Given.Select(g => g.Length == 1 ? g + "." : g));
    }

    /// <summary>
    /// Pemformat sitasi &amp; daftar pustaka: APA 7, IEEE, Harvard, MLA 9, Chicago (author-date), Vancouver.
    /// Teks miring ditandai *...* agar dirender sebagai italic oleh Word.
    /// </summary>
    public static class CitationFormatter
    {
        public static readonly string[] Styles = { "APA7", "IEEE", "Harvard", "MLA", "Chicago", "Vancouver" };

        public static string Normalize(string style)
        {
            var s = (style ?? "APA7").ToUpperInvariant().Replace(" ", "");
            if (s.StartsWith("APA")) return "APA7";
            if (s.StartsWith("IEEE")) return "IEEE";
            if (s.StartsWith("HARV")) return "Harvard";
            if (s.StartsWith("MLA")) return "MLA";
            if (s.StartsWith("CHIC")) return "Chicago";
            if (s.StartsWith("VANC")) return "Vancouver";
            return "APA7";
        }

        // ------------------------------------------------------------ sitasi dalam teks
        public static string InText(Reference r, string style, string page = null, int number = 0, bool indonesian = true)
        {
            style = Normalize(style);
            var names = r.Authors.Select(PersonName.Parse).ToList();
            string etal = indonesian ? "dkk." : "et al.";
            string year = string.IsNullOrWhiteSpace(r.Year) ? "t.t." : r.Year;
            string authorPart;
            if (names.Count == 0) authorPart = Shorten(r.Title);
            else if (names.Count == 1) authorPart = names[0].Last;
            else if (names.Count == 2) authorPart = names[0].Last + (style == "APA7" ? " & " : " dan ") + names[1].Last;
            else authorPart = names[0].Last + " " + etal;

            switch (style)
            {
                case "IEEE": return "[" + Math.Max(1, number) + "]";
                case "Vancouver": return "(" + Math.Max(1, number) + ")";
                case "MLA": return "(" + authorPart + (string.IsNullOrEmpty(page) ? "" : " " + page) + ")";
                case "Harvard": return "(" + authorPart + ", " + year + (string.IsNullOrEmpty(page) ? "" : ", hlm. " + page) + ")";
                case "Chicago": return "(" + authorPart + " " + year + (string.IsNullOrEmpty(page) ? "" : ", " + page) + ")";
                default: return "(" + authorPart + ", " + year + (string.IsNullOrEmpty(page) ? "" : ", hlm. " + page) + ")";
            }
        }

        // ------------------------------------------------------------ entri daftar pustaka
        public static string Entry(Reference r, string style, int number = 0)
        {
            style = Normalize(style);
            switch (style)
            {
                case "IEEE": return "[" + Math.Max(1, number) + "] " + Ieee(r);
                case "Harvard": return Harvard(r);
                case "MLA": return Mla(r);
                case "Chicago": return Chicago(r);
                case "Vancouver": return Math.Max(1, number) + ". " + Vancouver(r);
                default: return Apa(r);
            }
        }

        /// <summary>Daftar pustaka lengkap. APA/Harvard/MLA/Chicago diurutkan abjad; IEEE/Vancouver sesuai urutan kutipan.</summary>
        public static List<string> Bibliography(IEnumerable<Reference> refs, string style)
        {
            style = Normalize(style);
            var list = refs.ToList();
            if (style != "IEEE" && style != "Vancouver")
                list = list.OrderBy(r => PersonName.Parse(r.Authors.FirstOrDefault() ?? r.Title ?? "").Last, StringComparer.OrdinalIgnoreCase).ThenBy(r => r.Year).ToList();
            return list.Select((r, i) => Entry(r, style, i + 1)).ToList();
        }

        private static string Apa(Reference r)
        {
            var names = r.Authors.Select(PersonName.Parse).ToList();
            string authors;
            var formatted = names.Select(n => n.Single ? n.Last : n.Last + ", " + n.Initials()).ToList();
            if (formatted.Count == 0) authors = "";
            else if (formatted.Count == 1) authors = formatted[0];
            else if (formatted.Count <= 20) authors = string.Join(", ", formatted.Take(formatted.Count - 1)) + ", & " + formatted.Last();
            else authors = string.Join(", ", formatted.Take(19)) + ", . . . " + formatted.Last();
            var year = "(" + (string.IsNullOrWhiteSpace(r.Year) ? "t.t." : r.Year) + ").";
            var sb = new StringBuilder();
            if (authors.Length > 0) sb.Append(EndDot(authors)).Append(' ').Append(year);
            else sb.Append("*" + EndDot(r.Title).TrimEnd('.') + "*. ").Append(year);
            switch ((r.Type ?? "").ToLowerInvariant())
            {
                case "journal":
                case "proceeding":
                    sb.Append(' ').Append(EndDot(r.Title));
                    if (!string.IsNullOrWhiteSpace(r.Journal))
                    {
                        sb.Append(" *").Append(r.Journal);
                        if (!string.IsNullOrWhiteSpace(r.Volume)) sb.Append(", ").Append(r.Volume).Append('*');
                        else sb.Append('*');
                        if (!string.IsNullOrWhiteSpace(r.Issue)) sb.Append('(').Append(r.Issue).Append(')');
                        if (!string.IsNullOrWhiteSpace(r.Pages)) sb.Append(", ").Append(Dash(r.Pages));
                        sb.Append('.');
                    }
                    break;
                case "chapter":
                    sb.Append(' ').Append(EndDot(r.Title)).Append(" Dalam ");
                    if (r.Editors.Count > 0) sb.Append(string.Join(", ", r.Editors.Select(PersonName.Parse).Select(n => n.Initials() + " " + n.Last))).Append(r.Editors.Count > 1 ? " (Ed.), " : " (Ed.), ");
                    sb.Append('*').Append(r.BookTitle).Append('*');
                    if (!string.IsNullOrWhiteSpace(r.Pages)) sb.Append(" (hlm. ").Append(Dash(r.Pages)).Append(')');
                    sb.Append(". ").Append(EndDot(r.Publisher));
                    break;
                case "thesis":
                    sb.Append(" *").Append(r.Title.TrimEnd('.')).Append("* [").Append(string.IsNullOrWhiteSpace(r.Degree) ? "Skripsi" : r.Degree)
                      .Append(", ").Append(r.Institution).Append("].");
                    break;
                case "web":
                    sb.Append(" *").Append(r.Title.TrimEnd('.')).Append("*.");
                    if (!string.IsNullOrWhiteSpace(r.Publisher)) sb.Append(' ').Append(EndDot(r.Publisher));
                    break;
                case "regulation":
                    sb.Append(" *").Append(r.Title.TrimEnd('.')).Append("*.");
                    if (!string.IsNullOrWhiteSpace(r.Publisher)) sb.Append(' ').Append(EndDot(r.Publisher));
                    break;
                default:
                    sb.Append(" *").Append(r.Title.TrimEnd('.')).Append('*');
                    if (!string.IsNullOrWhiteSpace(r.Edition)) sb.Append(" (Ed. ke-").Append(r.Edition).Append(')');
                    sb.Append('.');
                    if (!string.IsNullOrWhiteSpace(r.Publisher)) sb.Append(' ').Append(EndDot(r.Publisher));
                    break;
            }
            if (!string.IsNullOrWhiteSpace(r.Doi)) sb.Append(" https://doi.org/").Append(r.Doi.Replace("https://doi.org/", ""));
            else if (!string.IsNullOrWhiteSpace(r.Url)) sb.Append(' ').Append(r.Url);
            return sb.ToString().Replace("..", ".").Trim();
        }

        private static string Ieee(Reference r)
        {
            var names = r.Authors.Select(PersonName.Parse).Select(n => n.Single ? n.Last : n.Initials() + " " + n.Last).ToList();
            string authors = names.Count == 0 ? "" : names.Count <= 2 ? string.Join(" and ", names) : names.Count > 6 ? names[0] + " et al." : string.Join(", ", names.Take(names.Count - 1)) + ", and " + names.Last();
            var sb = new StringBuilder(authors.Length > 0 ? authors + ", " : "");
            switch ((r.Type ?? "").ToLowerInvariant())
            {
                case "journal":
                case "proceeding":
                    sb.Append('"').Append(r.Title.TrimEnd('.')).Append(",\" *").Append(r.Journal).Append('*');
                    if (!string.IsNullOrWhiteSpace(r.Volume)) sb.Append(", vol. ").Append(r.Volume);
                    if (!string.IsNullOrWhiteSpace(r.Issue)) sb.Append(", no. ").Append(r.Issue);
                    if (!string.IsNullOrWhiteSpace(r.Pages)) sb.Append(", pp. ").Append(Dash(r.Pages));
                    sb.Append(", ").Append(r.Year);
                    if (!string.IsNullOrWhiteSpace(r.Doi)) sb.Append(", doi: ").Append(r.Doi);
                    sb.Append('.');
                    break;
                case "web":
                    sb.Append('"').Append(r.Title.TrimEnd('.')).Append(".\" ").Append(r.Publisher).Append(". [Online]. Available: ").Append(r.Url);
                    if (!string.IsNullOrWhiteSpace(r.Accessed)) sb.Append(" (accessed ").Append(r.Accessed).Append(')');
                    sb.Append('.');
                    break;
                case "thesis":
                    sb.Append('"').Append(r.Title.TrimEnd('.')).Append(",\" ").Append(string.IsNullOrWhiteSpace(r.Degree) ? "Skripsi" : r.Degree).Append(", ").Append(r.Institution).Append(", ").Append(r.City).Append(", ").Append(r.Year).Append('.');
                    break;
                default:
                    sb.Append('*').Append(r.Title.TrimEnd('.')).Append('*');
                    if (!string.IsNullOrWhiteSpace(r.Edition)) sb.Append(", ").Append(r.Edition).Append(" ed");
                    sb.Append(". ");
                    if (!string.IsNullOrWhiteSpace(r.City)) sb.Append(r.City).Append(": ");
                    sb.Append(r.Publisher).Append(", ").Append(r.Year).Append('.');
                    break;
            }
            return sb.ToString().Replace(", ,", ",").Trim();
        }

        private static string Harvard(Reference r)
        {
            var names = r.Authors.Select(PersonName.Parse).Select(n => n.Single ? n.Last : n.Last + ", " + n.Initials(false)).ToList();
            string authors = names.Count == 0 ? "" : names.Count == 1 ? names[0] : names.Count <= 3 ? string.Join(", ", names.Take(names.Count - 1)) + " dan " + names.Last() : names[0] + " dkk.";
            var sb = new StringBuilder(authors).Append(" (").Append(r.Year).Append(") ");
            if ((r.Type ?? "") == "journal")
            {
                sb.Append('\'').Append(r.Title.TrimEnd('.')).Append("', *").Append(r.Journal).Append("*, ").Append(r.Volume);
                if (!string.IsNullOrWhiteSpace(r.Issue)) sb.Append('(').Append(r.Issue).Append(')');
                if (!string.IsNullOrWhiteSpace(r.Pages)) sb.Append(", hlm. ").Append(Dash(r.Pages));
                sb.Append('.');
            }
            else
            {
                sb.Append('*').Append(r.Title.TrimEnd('.')).Append("*. ");
                if (!string.IsNullOrWhiteSpace(r.City)) sb.Append(r.City).Append(": ");
                sb.Append(EndDot(r.Publisher));
            }
            if (!string.IsNullOrWhiteSpace(r.Url)) sb.Append(" Tersedia di: ").Append(r.Url).Append(string.IsNullOrWhiteSpace(r.Accessed) ? "" : " (Diakses: " + r.Accessed + ")").Append('.');
            return sb.ToString().Trim();
        }

        private static string Mla(Reference r)
        {
            var names = r.Authors.Select(PersonName.Parse).ToList();
            string authors = names.Count == 0 ? "" : names.Count == 1 ? Mla1(names[0]) : names.Count == 2 ? Mla1(names[0]) + ", and " + names[1].GivenFull + " " + names[1].Last : Mla1(names[0]) + ", et al";
            var sb = new StringBuilder(authors.Length > 0 ? EndDot(authors) + " " : "");
            if ((r.Type ?? "") == "journal")
            {
                sb.Append('"').Append(r.Title.TrimEnd('.')).Append(".\" *").Append(r.Journal).Append("*");
                if (!string.IsNullOrWhiteSpace(r.Volume)) sb.Append(", vol. ").Append(r.Volume);
                if (!string.IsNullOrWhiteSpace(r.Issue)) sb.Append(", no. ").Append(r.Issue);
                sb.Append(", ").Append(r.Year);
                if (!string.IsNullOrWhiteSpace(r.Pages)) sb.Append(", pp. ").Append(Dash(r.Pages));
                sb.Append('.');
            }
            else sb.Append('*').Append(r.Title.TrimEnd('.')).Append("*. ").Append(r.Publisher).Append(", ").Append(r.Year).Append('.');
            return sb.ToString().Trim();
        }

        private static string Mla1(PersonName n) => n.Single ? n.Last : n.Last + ", " + n.GivenFull;

        private static string Chicago(Reference r)
        {
            var names = r.Authors.Select(PersonName.Parse).ToList();
            string authors = names.Count == 0 ? "" : names.Count == 1 ? Mla1(names[0]) : string.Join(", ", new[] { Mla1(names[0]) }.Concat(names.Skip(1).Take(names.Count - 2).Select(n => n.GivenFull + " " + n.Last))) + ", and " + names.Last().GivenFull + " " + names.Last().Last;
            var sb = new StringBuilder(authors.Length > 0 ? EndDot(authors) + " " : "").Append(r.Year).Append(". ");
            if ((r.Type ?? "") == "journal")
            {
                sb.Append('"').Append(r.Title.TrimEnd('.')).Append(".\" *").Append(r.Journal).Append("* ").Append(r.Volume);
                if (!string.IsNullOrWhiteSpace(r.Issue)) sb.Append(" (").Append(r.Issue).Append(')');
                if (!string.IsNullOrWhiteSpace(r.Pages)) sb.Append(": ").Append(Dash(r.Pages));
                sb.Append('.');
            }
            else
            {
                sb.Append('*').Append(r.Title.TrimEnd('.')).Append("*. ");
                if (!string.IsNullOrWhiteSpace(r.City)) sb.Append(r.City).Append(": ");
                sb.Append(EndDot(r.Publisher));
            }
            if (!string.IsNullOrWhiteSpace(r.Doi)) sb.Append(" https://doi.org/").Append(r.Doi);
            return sb.ToString().Trim();
        }

        private static string Vancouver(Reference r)
        {
            var names = r.Authors.Select(PersonName.Parse).Select(n => n.Single ? n.Last : n.Last + " " + n.Initials(false, false)).ToList();
            string authors = names.Count > 6 ? string.Join(", ", names.Take(6)) + ", et al" : string.Join(", ", names);
            var sb = new StringBuilder(authors.Length > 0 ? authors + ". " : "").Append(r.Title.TrimEnd('.')).Append(". ");
            if ((r.Type ?? "") == "journal")
            {
                sb.Append(r.Journal).Append(". ").Append(r.Year).Append(';').Append(r.Volume);
                if (!string.IsNullOrWhiteSpace(r.Issue)) sb.Append('(').Append(r.Issue).Append(')');
                if (!string.IsNullOrWhiteSpace(r.Pages)) sb.Append(':').Append(r.Pages);
                sb.Append('.');
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(r.City)) sb.Append(r.City).Append(": ");
                sb.Append(r.Publisher).Append("; ").Append(r.Year).Append('.');
            }
            return sb.ToString().Trim();
        }

        private static string EndDot(string s)
        {
            s = (s ?? "").Trim();
            return s.Length == 0 || s.EndsWith(".") || s.EndsWith("?") || s.EndsWith("!") ? s : s + ".";
        }

        private static string Dash(string pages) => (pages ?? "").Replace("--", "–").Replace("-", "–");

        private static string Shorten(string title)
        {
            var words = (title ?? "Anonim").Split(' ');
            return "*" + string.Join(" ", words.Take(Math.Min(4, words.Length))) + "*";
        }
    }
}
