using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MasterBW.Core.Indonesia
{
    public sealed class EydIssue
    {
        public int Start { get; set; }
        public int Length { get; set; }
        public string Original { get; set; }
        public string Suggestion { get; set; }
        public string Message { get; set; }
        /// <summary>baku | spasi | kapital | tanda_baca | penulisan_kata | pengulangan</summary>
        public string Category { get; set; }
        public override string ToString() => $"[{Category}] '{Original}' → '{Suggestion}': {Message}";
    }

    /// <summary>
    /// Pemeriksa ejaan bahasa Indonesia offline berbasis aturan EYD (Edisi V) dan KBBI:
    /// kata tidak baku, penulisan "di"/"ke" sebagai kata depan, partikel "pun", tanda baca, spasi,
    /// kapital awal kalimat, kata berulang, penulisan Rp dan persen. Cepat, tanpa AI.
    /// </summary>
    public static class EydChecker
    {
        /// <summary>Kata tidak baku → baku (KBBI). Bisa ditambah lewat file kamus pengguna.</summary>
        public static readonly Dictionary<string, string> KataBaku = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["aktifitas"] = "aktivitas", ["analisa"] = "analisis", ["apotik"] = "apotek", ["atlit"] = "atlet", ["praktek"] = "praktik",
            ["sistim"] = "sistem", ["resiko"] = "risiko", ["ijin"] = "izin", ["jaman"] = "zaman", ["nasehat"] = "nasihat",
            ["kwalitas"] = "kualitas", ["kwantitas"] = "kuantitas", ["kwitansi"] = "kuitansi", ["obyek"] = "objek", ["subyek"] = "subjek",
            ["obyektif"] = "objektif", ["subyektif"] = "subjektif", ["propinsi"] = "provinsi", ["standarisasi"] = "standardisasi",
            ["tehnik"] = "teknik", ["tehnologi"] = "teknologi", ["telpon"] = "telepon", ["merubah"] = "mengubah", ["ijasah"] = "ijazah",
            ["hutang"] = "utang", ["sekedar"] = "sekadar", ["silahkan"] = "silakan", ["nampak"] = "tampak", ["antri"] = "antre",
            ["cidera"] = "cedera", ["detil"] = "detail", ["efektifitas"] = "efektivitas", ["produktifitas"] = "produktivitas",
            ["kreatifitas"] = "kreativitas", ["sensitifitas"] = "sensitivitas", ["frekwensi"] = "frekuensi", ["hakekat"] = "hakikat",
            ["jadual"] = "jadwal", ["karir"] = "karier", ["komplek"] = "kompleks", ["lembab"] = "lembap", ["menterjemahkan"] = "menerjemahkan",
            ["mempengaruhi"] = "memengaruhi", ["mempesona"] = "memesona", ["memperhatikan"] = "memperhatikan", ["nafas"] = "napas",
            ["orisinil"] = "orisinal", ["pebruari"] = "Februari", ["nopember"] = "November", ["prosentase"] = "persentase",
            ["rapot"] = "rapor", ["relijius"] = "religius", ["sholat"] = "salat", ["shalat"] = "salat", ["tauladan"] = "teladan",
            ["trampil"] = "terampil", ["ketrampilan"] = "keterampilan", ["walikota"] = "wali kota", ["apapun"] = "apa pun",
            ["siapapun"] = "siapa pun", ["dimanapun"] = "di mana pun", ["kapanpun"] = "kapan pun", ["manapun"] = "mana pun",
            ["satupun"] = "satu pun", ["sedikitpun"] = "sedikit pun", ["kemana"] = "ke mana", ["dimana"] = "di mana",
            ["kesana"] = "ke sana", ["kesini"] = "ke sini", ["disini"] = "di sini", ["disana"] = "di sana", ["disitu"] = "di situ",
            ["diatas"] = "di atas", ["dibawah"] = "di bawah", ["didalam"] = "di dalam", ["diluar"] = "di luar", ["didepan"] = "di depan",
            ["dibelakang"] = "di belakang", ["disamping"] = "di samping", ["diantara"] = "di antara", ["ditengah"] = "di tengah",
            ["disekitar"] = "di sekitar", ["kedalam"] = "ke dalam", ["keluar negeri"] = "ke luar negeri", ["keatas"] = "ke atas",
            ["kebawah"] = "ke bawah", ["sekolah dasar"] = "sekolah dasar", ["fikir"] = "pikir", ["difikirkan"] = "dipikirkan",
            ["faham"] = "paham", ["fihak"] = "pihak", ["fasal"] = "pasal", ["ruwet"] = "rumit", ["gimana"] = "bagaimana",
            ["nggak"] = "tidak", ["gak"] = "tidak", ["udah"] = "sudah", ["aja"] = "saja", ["kalo"] = "kalau", ["bikin"] = "buat",
            ["gini"] = "begini", ["gitu"] = "begitu", ["kayak"] = "seperti", ["tapi"] = "tetapi", ["trus"] = "terus", ["dgn"] = "dengan",
            ["yg"] = "yang", ["utk"] = "untuk", ["tdk"] = "tidak", ["krn"] = "karena", ["sdh"] = "sudah", ["blm"] = "belum",
            ["dlm"] = "dalam", ["thd"] = "terhadap", ["pd"] = "pada", ["spt"] = "seperti", ["bhw"] = "bahwa", ["jg"] = "juga",
            ["mengkoordinir"] = "mengoordinasi", ["koordinir"] = "koordinasi", ["mensukseskan"] = "menyukseskan", ["mentaati"] = "menaati",
            ["menyolok"] = "mencolok", ["merubahnya"] = "mengubahnya", ["mengkonsumsi"] = "mengonsumsi", ["mengkritik"] = "mengkritik",
            ["menkonfirmasi"] = "mengonfirmasi", ["mengkonfirmasi"] = "mengonfirmasi", ["memproduksi"] = "memproduksi",
            ["advis"] = "advis", ["asas"] = "asas", ["azas"] = "asas", ["azan"] = "azan", ["bis"] = "bus", ["cabe"] = "cabai",
            ["cendikiawan"] = "cendekiawan", ["diagnosa"] = "diagnosis", ["do'a"] = "doa", ["ekstrim"] = "ekstrem", ["familiar"] = "familier",
            ["formil"] = "formal", ["foto copy"] = "fotokopi", ["fotocopy"] = "fotokopi", ["hipotesa"] = "hipotesis", ["himbau"] = "imbau",
            ["himbauan"] = "imbauan", ["insaf"] = "insaf", ["istirahat"] = "istirahat", ["jum'at"] = "Jumat", ["kaos"] = "kaus",
            ["katagori"] = "kategori", ["kerjasama"] = "kerja sama", ["konkrit"] = "konkret", ["korp"] = "korps", ["legalisir"] = "legalisasi",
            ["lisensi"] = "lisensi", ["managemen"] = "manajemen", ["manajer"] = "manajer", ["metoda"] = "metode", ["methode"] = "metode",
            ["mie"] = "mi", ["moderen"] = "modern", ["nasionalisasi"] = "nasionalisasi", ["negri"] = "negeri", ["olah raga"] = "olahraga",
            ["paska"] = "pasca", ["pasca sarjana"] = "pascasarjana", ["permisi"] = "permisi", ["photo"] = "foto", ["proyek"] = "proyek",
            ["projek"] = "proyek", ["ramadhan"] = "Ramadan", ["respon"] = "respons", ["sintesa"] = "sintesis", ["sorga"] = "surga",
            ["supir"] = "sopir", ["syah"] = "sah", ["tanggungjawab"] = "tanggung jawab", ["terlanjur"] = "telanjur", ["thema"] = "tema",
            ["tolok ukur"] = "tolok ukur", ["tolak ukur"] = "tolok ukur", ["ustad"] = "ustaz", ["ustadz"] = "ustaz", ["varitas"] = "varietas",
            ["zona"] = "zona", ["sekretaris"] = "sekretaris", ["sekertaris"] = "sekretaris", ["sistimatis"] = "sistematis",
            ["pondasi"] = "fondasi", ["silabus"] = "silabus", ["kurikulum"] = "kurikulum", ["mushola"] = "musala", ["musholla"] = "musala",
            ["Al-Qur'an"] = "Al-Qur'an", ["alquran"] = "Al-Qur'an", ["ektra"] = "ekstra", ["extra"] = "ekstra", ["design"] = "desain",
            ["konsekwen"] = "konsekuen", ["konsekwensi"] = "konsekuensi", ["kongkrit"] = "konkret", ["mengcopy"] = "menyalin",
            ["subsidi"] = "subsidi", ["investasi"] = "investasi", ["absen"] = "absen", ["indera"] = "indra", ["inderawi"] = "indrawi",
        };

        /// <summary>Kata dasar kerja umum: "di" + kata ini seharusnya DISAMBUNG (pasif), mis. "di tulis" → "ditulis".</summary>
        private static readonly HashSet<string> VerbBases = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "tulis", "baca", "buat", "lakukan", "laksanakan", "kirim", "terima", "ambil", "beri", "berikan", "jual", "beli", "makan", "minum",
            "pakai", "gunakan", "bahas", "jelaskan", "uraikan", "analisis", "teliti", "ukur", "hitung", "susun", "kerjakan", "bangun",
            "simpan", "hapus", "ubah", "tetapkan", "sahkan", "tandatangani", "periksa", "uji", "nilai", "harapkan", "mohon", "minta",
            "undang", "hadiri", "ikuti", "pilih", "temukan", "lihat", "dengar", "rasakan", "anggap", "sebut", "namakan", "jadikan",
            "lampirkan", "sampaikan", "tujukan", "tunjuk", "angkat", "berhentikan", "bayar", "cetak", "isi", "olah", "proses", "kelola",
            "atur", "tentukan", "peroleh", "dapat", "dapatkan", "pengaruhi", "sajikan", "tampilkan", "kumpulkan", "catat", "rekam",
        };

        /// <summary>Kata yang setelah "di" harus DIPISAH (kata depan tempat).</summary>
        private static readonly HashSet<string> PlaceWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "atas", "bawah", "dalam", "luar", "sini", "sana", "situ", "mana", "depan", "belakang", "samping", "antara", "tengah",
            "sekitar", "rumah", "sekolah", "kantor", "kampus", "kelas", "pasar", "kota", "desa", "jalan", "sebelah", "seberang", "hadapan",
            "balik", "muka", "pinggir", "ujung", "pojok", "lapangan", "masjid", "gereja", "perpustakaan", "laboratorium", "indonesia", "jakarta",
        };

        private static readonly string[] Days = { "senin", "selasa", "rabu", "kamis", "jumat", "sabtu", "minggu" };

        public static List<EydIssue> Check(string text)
        {
            var issues = new List<EydIssue>();
            if (string.IsNullOrEmpty(text)) return issues;

            // 1. Kata tidak baku (termasuk frasa 2 kata)
            foreach (Match m in Regex.Matches(text, @"\b(\p{L}+) (?=(\p{L}+)\b)"))
            {
                // frasa 2 kata (pencocokan tumpang-tindih), mis. "tolak ukur", "olah raga"
                var phrase = m.Groups[1].Value + " " + m.Groups[2].Value;
                if (KataBaku.TryGetValue(phrase, out var bakuPhrase) && !string.Equals(bakuPhrase, phrase, StringComparison.OrdinalIgnoreCase))
                    issues.Add(Issue(m.Index, phrase.Length, phrase, MatchCase(phrase, bakuPhrase), "Bentuk tidak baku; KBBI: " + bakuPhrase + ".", "baku"));
            }
            foreach (Match m in Regex.Matches(text, @"\b[\p{L}][\p{L}'’]*\b"))
            {
                var w = m.Value;
                if (KataBaku.TryGetValue(w, out var baku) && !string.Equals(baku, w, StringComparison.OrdinalIgnoreCase)
                    && !issues.Any(i => i.Start <= m.Index && i.Start + i.Length >= m.Index + m.Length))
                {
                    var cat = baku.Contains(" ") ? "penulisan_kata" : "baku";
                    issues.Add(Issue(m.Index, w.Length, w, MatchCase(w, baku), cat == "baku" ? "Kata tidak baku; KBBI: " + baku + "." : "Ditulis terpisah: " + baku + ".", cat));
                }
            }

            // 2. "di"/"ke" sebagai awalan vs kata depan
            foreach (Match m in Regex.Matches(text, @"\b([Dd]i) ([\p{L}]+)\b"))
            {
                var next = m.Groups[2].Value;
                if (VerbBases.Contains(next) && !PlaceWords.Contains(next))
                    issues.Add(Issue(m.Index, m.Length, m.Value, m.Groups[1].Value + next.ToLowerInvariant(), "Awalan di- pada kata kerja pasif ditulis serangkai.", "penulisan_kata"));
            }
            foreach (Match m in Regex.Matches(text, @"\b([Dd]i)(atas|bawah|dalam|luar|depan|belakang|samping|antara|tengah|sekitar|sebelah|rumah|sekolah|kantor|kampus|kelas)\b", RegexOptions.IgnoreCase))
            {
                if (issues.Any(i => i.Start == m.Index)) continue;
                issues.Add(Issue(m.Index, m.Length, m.Value, m.Groups[1].Value + " " + m.Groups[2].Value, "Kata depan di ditulis terpisah dari kata yang menunjukkan tempat.", "penulisan_kata"));
            }
            foreach (Match m in Regex.Matches(text, @"\b([Kk]e) ([0-9]+)\b"))
                issues.Add(Issue(m.Index, m.Length, m.Value, m.Groups[1].Value + "-" + m.Groups[2].Value, "Awalan ke- dengan angka ditulis dengan tanda hubung (ke-2).", "penulisan_kata"));

            // 3. Spasi & tanda baca
            foreach (Match m in Regex.Matches(text, @"[^\S\r\n]+([,.;:!?])(?!\d)"))
                issues.Add(Issue(m.Index, m.Length, m.Value, m.Groups[1].Value, "Tidak ada spasi sebelum tanda baca.", "spasi"));
            foreach (Match m in Regex.Matches(text, @"([,;:])(?=[\p{L}])"))
                issues.Add(Issue(m.Index, 1, m.Value, m.Value + " ", "Beri spasi setelah tanda baca.", "spasi"));
            foreach (Match m in Regex.Matches(text, @"(?<=[\p{Ll}]{2})\.(?=[\p{Lu}][\p{Ll}])"))
                issues.Add(Issue(m.Index, 1, ".", ". ", "Beri spasi setelah titik.", "spasi"));
            foreach (Match m in Regex.Matches(text, @"(?<=\S)[^\S\r\n]{2,}(?=\S)"))
                issues.Add(Issue(m.Index, m.Length, m.Value, " ", "Spasi ganda.", "spasi"));
            foreach (Match m in Regex.Matches(text, @"([!?,;])\1+"))
                issues.Add(Issue(m.Index, m.Length, m.Value, m.Groups[1].Value, "Tanda baca berulang.", "tanda_baca"));
            foreach (Match m in Regex.Matches(text, @"\bRp\.?\s+(?=\d)|\bRp\.(?=\d)"))
                issues.Add(Issue(m.Index, m.Length, m.Value, "Rp", "Lambang Rp ditulis tanpa titik dan tanpa spasi (Rp10.000).", "tanda_baca"));
            foreach (Match m in Regex.Matches(text, @"(?<=\d) %"))
                issues.Add(Issue(m.Index, m.Length, m.Value, "%", "Lambang persen ditulis rapat dengan angka (50%).", "spasi"));

            // 4. Kata berulang tidak sengaja ("yang yang")
            foreach (Match m in Regex.Matches(text, @"\b([\p{L}]+)\s+\1\b", RegexOptions.IgnoreCase))
            {
                var w = m.Groups[1].Value.ToLowerInvariant();
                if (w.Length <= 1) continue;
                issues.Add(Issue(m.Index, m.Length, m.Value, m.Groups[1].Value, "Kata berulang. Jika bentuk ulang, gunakan tanda hubung (" + w + "-" + w + ").", "pengulangan"));
            }

            // 5. Huruf kapital awal kalimat & nama hari
            foreach (Match m in Regex.Matches(text, @"(?:^|[.!?]\s+)([\p{Ll}])", RegexOptions.Multiline))
            {
                var g = m.Groups[1];
                // jangan tandai singkatan umum: "dll. dan", "dsb. yang", "a.n. kepala"
                var before = text.Substring(Math.Max(0, g.Index - 6), Math.Min(6, g.Index));
                if (Regex.IsMatch(before, @"(dll|dsb|dst|a\.n|u\.p|no|hlm|tgl|yth|dkk|Bpk|Sdr|Kec|Kab|Jl|vol)\.\s*$", RegexOptions.IgnoreCase)) continue;
                if (Regex.IsMatch(before, @"\d\.\s*$")) continue;
                issues.Add(Issue(g.Index, 1, g.Value, g.Value.ToUpperInvariant(), "Awal kalimat diawali huruf kapital.", "kapital"));
            }
            foreach (Match m in Regex.Matches(text, @"\b(" + string.Join("|", Days) + @")\b"))
            {
                if (issues.Any(i => i.Start == m.Index)) continue;
                issues.Add(Issue(m.Index, m.Length, m.Value, char.ToUpperInvariant(m.Value[0]) + m.Value.Substring(1), "Nama hari diawali huruf kapital.", "kapital"));
            }

            return issues.GroupBy(i => i.Start).Select(g => g.First()).OrderBy(i => i.Start).ToList();
        }

        /// <summary>Terapkan semua saran (dari belakang agar indeks tidak bergeser).</summary>
        public static string ApplyAll(string text, IEnumerable<EydIssue> issues)
        {
            var result = text;
            int lastStart = int.MaxValue;
            foreach (var i in issues.OrderByDescending(x => x.Start))
            {
                if (i.Start + i.Length > lastStart) continue; // tumpang tindih
                result = result.Substring(0, i.Start) + i.Suggestion + result.Substring(i.Start + i.Length);
                lastStart = i.Start;
            }
            return result;
        }

        private static EydIssue Issue(int start, int len, string orig, string sug, string msg, string cat)
            => new EydIssue { Start = start, Length = len, Original = orig, Suggestion = sug, Message = msg, Category = cat };

        private static string MatchCase(string original, string replacement)
        {
            if (original.Length == 0 || replacement.Length == 0) return replacement;
            if (original.All(c => !char.IsLetter(c) || char.IsUpper(c)) && original.Length > 1) return replacement.ToUpperInvariant();
            if (char.IsUpper(original[0])) return char.ToUpperInvariant(replacement[0]) + replacement.Substring(1);
            return replacement;
        }
    }
}
