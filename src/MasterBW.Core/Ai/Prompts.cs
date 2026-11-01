using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MasterBW.Core.ExcelKit;
using MasterBW.Core.Tools;

namespace MasterBW.Core.Ai
{
    /// <summary>Seluruh prompt MASTER BW (bahasa Indonesia), dioptimalkan untuk model kecil offline.</summary>
    public static class Prompts
    {
        public const string Persona =
            "Kamu adalah MASTER BW, asisten AI offline untuk Microsoft Office 2016 yang ahli di bidang akademik, perkantoran, administrasi, " +
            "dan keuangan di Indonesia. Selalu menjawab dalam bahasa Indonesia yang baku, jelas, dan sesuai EYD, kecuali diminta bahasa lain.";

        // ------------------------------------------------------------------ Perencana aksi
        public static List<ChatMessage> BuildPlanMessages(OfficeApp app, string request, OfficeContext ctx, IList<ToolDef> tools,
                                                          IList<ChatMessage> history = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine(Persona);
            sb.AppendLine("Tugasmu: ubah perintah pengguna untuk Microsoft " + app + " menjadi rencana aksi JSON.");
            sb.AppendLine("Format WAJIB: {\"message\":\"penjelasan singkat untuk pengguna\",\"steps\":[{\"action\":\"nama_aksi\",\"args\":{...}}]}");
            sb.AppendLine("Aturan:");
            sb.AppendLine("1. Hanya pakai aksi dari daftar di bawah. Argumen bertanda * wajib diisi.");
            sb.AppendLine("2. Jika pengguna hanya bertanya/ngobrol, jawab di \"message\" dan biarkan \"steps\" kosong [].");
            sb.AppendLine("3. Pecah perintah besar menjadi beberapa langkah berurutan. Jangan mengarang data yang tidak diberikan; tulis [isi] bila perlu.");
            if (app == OfficeApp.Excel)
            {
                sb.AppendLine("4. Rumus memakai nama fungsi bahasa Inggris dan pemisah koma, misal =SUMIFS(F:F,B:B,\"Jawa\").");
                sb.AppendLine("5. Ini Excel 2016: DILARANG memakai " + string.Join(", ", ExcelFunctionInfo.NotIn2016.Take(14)) + ". Pakai SUMPRODUCT/INDEX-MATCH/IF bertingkat sebagai gantinya.");
                sb.AppendLine("6. Gunakan alamat sel/kolom sesuai konteks. Rumus untuk sel pertama saja, akan disalin otomatis ke bawah.");
            }
            else
            {
                sb.AppendLine("4. Untuk menulis isi panjang (paragraf, bab) pakai generate_content; untuk mengolah teks terpilih pakai transform_text.");
                sb.AppendLine("5. Surat/SK/notulen/proposal/laporan memakai create_from_template bila ada templatnya.");
            }
            sb.AppendLine();
            sb.AppendLine("Daftar aksi:");
            foreach (var t in tools) sb.AppendLine(t.ToPromptLine());
            if (app == OfficeApp.Word)
            {
                sb.AppendLine();
                sb.AppendLine("Templat Word: " + string.Join(", ", Academic.DocumentTemplates.All.Select(t => t.Id)));
            }
            else
            {
                sb.AppendLine();
                sb.AppendLine("Templat Excel: " + string.Join(", ", ExcelTemplates.All.Select(t => t.Id)));
            }
            sb.AppendLine();
            sb.AppendLine("Contoh:");
            sb.AppendLine(app == OfficeApp.Excel ? ExcelExamples : WordExamples);

            var messages = new List<ChatMessage> { ChatMessage.System(sb.ToString().TrimEnd()) };
            if (history != null) messages.AddRange(history.Skip(Math.Max(0, history.Count - 4)));
            messages.Add(ChatMessage.User("Konteks saat ini:\n" + (ctx?.ToPromptText() ?? "-") + "\n\nPerintah: " + request));
            return messages;
        }

        private const string WordExamples =
            "Perintah: buat judul BAB I PENDAHULUAN lalu tulis latar belakang tentang literasi digital\n" +
            "{\"message\":\"Saya tambahkan judul BAB I dan latar belakang.\",\"steps\":[{\"action\":\"insert_heading\",\"args\":{\"text\":\"BAB I PENDAHULUAN\",\"level\":1}},{\"action\":\"insert_heading\",\"args\":{\"text\":\"1.1 Latar Belakang\",\"level\":2}},{\"action\":\"generate_content\",\"args\":{\"instruction\":\"latar belakang penelitian tentang literasi digital siswa\",\"length\":\"panjang\",\"style\":\"akademik\"}}]}\n" +
            "Perintah: semua teks jadi Times New Roman 12 spasi 1,5 rata kiri kanan\n" +
            "{\"message\":\"Format seluruh dokumen diubah.\",\"steps\":[{\"action\":\"format_font\",\"args\":{\"scope\":\"document\",\"name\":\"Times New Roman\",\"size\":12}},{\"action\":\"format_paragraph\",\"args\":{\"scope\":\"document\",\"line_spacing\":1.5,\"alignment\":\"justify\"}}]}";

        private const string ExcelExamples =
            "Konteks: Kolom: A=Tanggal, B=Wilayah, C=Produk, D=Qty, E=Harga. Area data A1:E200.\n" +
            "Perintah: tambah kolom total lalu format rupiah\n" +
            "{\"message\":\"Kolom F berisi Total = Qty × Harga dengan format Rupiah.\",\"steps\":[{\"action\":\"write_formula\",\"args\":{\"range\":\"F2:F200\",\"formula\":\"=D2*E2\",\"header\":\"Total\"}},{\"action\":\"format_number\",\"args\":{\"range\":\"F2:F200\",\"format\":\"rupiah\"}}]}\n" +
            "Perintah: total penjualan wilayah Jawa di H2\n" +
            "{\"message\":\"H2 berisi total untuk wilayah Jawa.\",\"steps\":[{\"action\":\"write_formula\",\"args\":{\"range\":\"H2\",\"formula\":\"=SUMIFS(F:F,B:B,\\\"Jawa\\\")\"}}]}";

        // ------------------------------------------------------------------ Olah teks
        public static readonly Dictionary<string, string> TextTasks = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["ringkas"] = "Ringkas teks berikut menjadi sekitar sepertiga panjangnya. Pertahankan gagasan utama, angka, dan istilah penting.",
            ["parafrase"] = "Parafrasekan teks berikut dengan struktur kalimat dan pilihan kata yang berbeda tanpa mengubah makna. Hindari kalimat yang sama persis dengan aslinya.",
            ["perbaiki"] = "Perbaiki ejaan, tanda baca, kata tidak baku, dan tata bahasa teks berikut sesuai EYD. Jangan mengubah isi.",
            ["formal"] = "Ubah teks berikut menjadi bahasa Indonesia formal dan sopan untuk dokumen resmi kantor.",
            ["akademik"] = "Ubah teks berikut menjadi gaya bahasa ilmiah/akademik: objektif, kalimat efektif, tanpa kata ganti orang pertama, istilah tepat.",
            ["santai"] = "Ubah teks berikut menjadi bahasa yang santai, ramah, dan mudah dipahami.",
            ["sederhanakan"] = "Sederhanakan teks berikut agar mudah dipahami pembaca awam. Gunakan kalimat pendek.",
            ["perpanjang"] = "Kembangkan teks berikut menjadi lebih lengkap dan mendalam (sekitar dua kali lipat), tambahkan penjelasan dan contoh yang relevan.",
            ["lanjutkan"] = "Lanjutkan tulisan berikut dengan gaya, nada, dan topik yang sama sebanyak 2-3 paragraf. Tulis lanjutannya saja.",
            ["poin"] = "Ubah teks berikut menjadi daftar poin yang ringkas dan terstruktur (gunakan '- ' di awal setiap poin).",
            ["terjemah_en"] = "Terjemahkan teks berikut ke bahasa Inggris yang alami dan tepat. Tulis terjemahannya saja.",
            ["terjemah_id"] = "Terjemahkan teks berikut ke bahasa Indonesia baku yang alami. Tulis terjemahannya saja.",
            ["judul"] = "Buat 5 alternatif judul yang menarik dan tepat untuk teks berikut, masing-masing satu baris diawali '- '.",
            ["abstrak"] = "Susun abstrak akademik 150-250 kata (latar belakang, tujuan, metode, hasil, simpulan) dari teks berikut, diakhiri 'Kata kunci:' 3-5 kata.",
            ["kesimpulan"] = "Tulis kesimpulan yang padat dari teks berikut.",
            ["pertanyaan"] = "Buat 5 pertanyaan pemahaman beserta jawabannya dari teks berikut.",
            ["email"] = "Ubah teks berikut menjadi email resmi yang sopan dengan salam pembuka dan penutup.",
        };

        public static List<ChatMessage> BuildTextTask(string task, string text, string extra = null)
        {
            if (!TextTasks.TryGetValue(task ?? "", out var instruction)) instruction = task;
            var sys = Persona + " Keluarkan HANYA hasil akhirnya tanpa pembuka, tanpa komentar, tanpa tanda kutip pembungkus.";
            var user = instruction + (string.IsNullOrWhiteSpace(extra) ? "" : "\nCatatan tambahan: " + extra) + "\n\nTeks:\n\"\"\"\n" + text + "\n\"\"\"";
            return new List<ChatMessage> { ChatMessage.System(sys), ChatMessage.User(user) };
        }

        public static List<ChatMessage> BuildContent(string instruction, string length, string style, string documentContext)
        {
            string words;
            switch ((length ?? "").ToLowerInvariant())
            {
                case "pendek": words = "sekitar 1 paragraf (80-120 kata)"; break;
                case "panjang": words = "4-6 paragraf (400-600 kata)"; break;
                default: words = "2-3 paragraf (200-300 kata)"; break;
            }
            var sys = Persona + " Kamu penulis profesional. Tulis langsung isinya tanpa pembuka seperti 'Berikut adalah'. " +
                      "Gunakan paragraf yang dipisah baris kosong. Boleh memakai '## ' untuk subjudul dan '- ' untuk poin bila perlu. " +
                      "Jangan mengarang kutipan, angka statistik, atau nama sumber yang tidak pasti; tulis [sumber] sebagai penanda bila perlu rujukan.";
            var user = "Tulis: " + instruction + "\nPanjang: " + words + "\nGaya: " + (style ?? "formal") +
                       (string.IsNullOrWhiteSpace(documentContext) ? "" : "\nKonteks dokumen:\n" + documentContext);
            return new List<ChatMessage> { ChatMessage.System(sys), ChatMessage.User(user) };
        }

        public static List<ChatMessage> BuildTemplateFill(string templateName, string sectionPurpose, string fieldsSummary)
        {
            var sys = Persona + " Kamu menulis isi dokumen resmi Indonesia. Tulis hanya isi bagian yang diminta, ringkas dan baku, tanpa judul, tanpa salam pembuka jika tidak diminta.";
            var user = "Dokumen: " + templateName + "\nBagian: " + sectionPurpose + "\nData yang diketahui:\n" + fieldsSummary;
            return new List<ChatMessage> { ChatMessage.System(sys), ChatMessage.User(user) };
        }

        // ------------------------------------------------------------------ Excel
        public static List<ChatMessage> BuildFormulaFix(string cell, string formula, string error, string diagnosis, string context)
        {
            var sys = Persona + " Kamu ahli rumus Excel 2016. Jawab dengan JSON: {\"explanation\":\"penyebab singkat\",\"formula\":\"=rumus perbaikan\"}. " +
                      "Rumus memakai nama fungsi Inggris dan pemisah koma. Dilarang memakai fungsi yang tidak ada di Excel 2016.";
            var user = "Sel " + cell + " berisi rumus " + formula + " menghasilkan " + (error ?? "hasil salah") + ".\nDiagnosis awal: " + diagnosis + "\nKonteks:\n" + context;
            return new List<ChatMessage> { ChatMessage.System(sys), ChatMessage.User(user) };
        }

        public static List<ChatMessage> BuildInsight(string statsSummary, string question)
        {
            var sys = Persona + " Kamu analis data senior. Dari ringkasan statistik berikut, tulis temuan penting (insight), tren, anomali, dan rekomendasi tindakan. " +
                      "Gunakan poin '- '. Sebut angka secara spesifik. Jangan mengarang angka yang tidak ada di ringkasan.";
            var user = (string.IsNullOrWhiteSpace(question) ? "" : "Fokus pertanyaan: " + question + "\n") + "Ringkasan data:\n" + statsSummary;
            return new List<ChatMessage> { ChatMessage.System(sys), ChatMessage.User(user) };
        }

        public static List<ChatMessage> BuildRowTask(string instruction, IList<string> categories, string value)
        {
            var sys = "Kamu memproses data satu baris. Jawab HANYA dengan hasil akhirnya dalam satu baris pendek, tanpa penjelasan.";
            var cat = categories != null && categories.Count > 0 ? "\nPilih tepat satu dari: " + string.Join(", ", categories) : "";
            return new List<ChatMessage> { ChatMessage.System(sys), ChatMessage.User(instruction + cat + "\nData: " + value) };
        }

        public static List<ChatMessage> BuildVba(string description, OfficeApp app, string context)
        {
            var sys = Persona + " Kamu programmer VBA ahli untuk " + app + " 2016. Tulis satu modul VBA lengkap (Option Explicit, Sub utama bernama MasterBW_Makro) " +
                      "dengan komentar bahasa Indonesia. Jangan menghapus file, mengakses internet, atau registry. Keluarkan kode saja di dalam ```vba```.";
            return new List<ChatMessage> { ChatMessage.System(sys), ChatMessage.User("Buat makro: " + description + "\nKonteks:\n" + context) };
        }

        public static List<ChatMessage> BuildChat(string question, OfficeContext ctx, IList<ChatMessage> history)
        {
            var messages = new List<ChatMessage> { ChatMessage.System(Persona + " Jawab dengan ringkas, terstruktur, dan praktis.") };
            if (history != null) messages.AddRange(history.Skip(Math.Max(0, history.Count - 6)));
            messages.Add(ChatMessage.User((ctx == null ? "" : "Konteks:\n" + ctx.ToPromptText() + "\n\n") + question));
            return messages;
        }
    }
}
