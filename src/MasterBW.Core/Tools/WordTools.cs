using System.Collections.Generic;

namespace MasterBW.Core.Tools
{
    /// <summary>Katalog aksi Microsoft Word. Nama aksi dalam bahasa Inggris (lebih stabil untuk model kecil), label &amp; deskripsi bahasa Indonesia.</summary>
    internal static class WordTools
    {
        private static readonly string[] Scope = { "selection", "paragraph", "document" };
        private static readonly string[] Pos = { "cursor", "start", "end", "replace_selection" };

        private static ToolDef T(string name, string label, string category, string desc)
            => new ToolDef { Name = name, App = OfficeApp.Word, Label = label, Category = category, Description = desc };

        public static List<ToolDef> Build()
        {
            return new List<ToolDef>
            {
                // ===================== Menulis =====================
                T("insert_text", "Sisipkan teks", "Menulis", "Sisipkan teks. Mendukung markdown ringan: # judul, ## subjudul, - poin, 1. nomor, **tebal**, *miring*.")
                    .P("text", "string", "teks yang disisipkan", true)
                    .P("position", "string", "lokasi", false, Pos, "cursor")
                    .K("tulis", "ketik", "sisip", "tambah teks", "insert", "write"),

                T("generate_content", "Tulis dengan AI", "Menulis", "AI menulis konten baru (paragraf/bagian) sesuai instruksi lalu menyisipkannya.")
                    .P("instruction", "string", "apa yang harus ditulis, mis. 'latar belakang tentang stunting'", true)
                    .P("length", "string", "panjang", false, new[] { "pendek", "sedang", "panjang" }, "sedang")
                    .P("style", "string", "gaya bahasa", false, new[] { "formal", "akademik", "santai", "persuasif", "jurnalistik" }, "formal")
                    .P("position", "string", "lokasi", false, Pos, "cursor")
                    .K("buatkan", "tuliskan", "karang", "susun", "paragraf", "latar belakang", "pendahuluan", "kesimpulan", "abstrak", "kata pengantar", "pembahasan", "esai", "artikel", "pidato", "sambutan")
                    .Ai(),

                T("transform_text", "Olah teks dengan AI", "Menulis", "AI mengolah teks terpilih (atau seluruh dokumen) lalu menggantinya.")
                    .P("task", "string", "jenis olahan", true, new[] { "ringkas", "parafrase", "perbaiki", "formal", "sederhanakan", "perpanjang", "lanjutkan", "poin", "terjemah_en", "terjemah_id", "judul", "akademik", "santai" })
                    .P("scope", "string", "sumber teks", false, new[] { "selection", "document" }, "selection")
                    .P("output", "string", "hasil diletakkan", false, new[] { "replace", "below", "new_document", "comment" }, "replace")
                    .K("ringkas", "rangkum", "parafrase", "ubah kalimat", "perbaiki", "terjemah", "translate", "sederhanakan", "perpanjang", "lanjutkan", "formal", "summary", "paraphrase")
                    .Ai(),

                T("build_document", "Susun dokumen lengkap", "Dokumen", "Bangun dokumen terstruktur dari daftar blok. Tipe blok: title, heading(level 1-3), paragraph, bullets(items), numbered(items), table(rows), signature, page_break, toc, letterhead, spacer.")
                    .P("blocks", "array", "daftar blok {type, text, level, items, rows}", true)
                    .P("position", "string", "lokasi", false, new[] { "cursor", "end", "new_document" }, "cursor")
                    .K("buat dokumen", "susun dokumen", "kerangka", "outline", "struktur"),

                T("create_from_template", "Buat dari templat", "Dokumen", "Buat dokumen dari templat bawaan (surat, SK, notulen, proposal, laporan, skripsi, dll). Kolom yang tidak diketahui dibiarkan [isi].")
                    .P("template", "string", "id templat, lihat daftar templat", true)
                    .P("fields", "object", "isian {nama_field: nilai}")
                    .P("ai_fill", "boolean", "AI mengisi bagian isi/paragraf otomatis", false, null, "true")
                    .K("surat", "undangan", "sk", "surat keputusan", "notulen", "proposal", "laporan", "berita acara", "memo", "surat kuasa", "pernyataan", "skripsi", "makalah", "cv", "lamaran", "templat", "template", "rpp", "modul ajar")
                    .Ai(),

                T("insert_cover", "Halaman sampul", "Akademik", "Sisipkan halaman sampul/cover akademik di awal dokumen.")
                    .P("title", "string", "judul karya", true)
                    .P("kind", "string", "jenis karya", false, new[] { "SKRIPSI", "TESIS", "MAKALAH", "LAPORAN", "PROPOSAL", "TUGAS AKHIR", "LAPORAN PKL", "KARYA TULIS ILMIAH" }, "MAKALAH")
                    .P("author", "string", "nama penulis")
                    .P("id_number", "string", "NIM/NIS/NIP")
                    .P("program", "string", "program studi/jurusan")
                    .P("institution", "string", "nama kampus/sekolah/instansi")
                    .P("city", "string", "kota")
                    .P("year", "string", "tahun")
                    .P("logo_path", "string", "path file logo")
                    .K("sampul", "cover", "halaman judul", "halaman depan"),

                // ===================== Format =====================
                T("insert_heading", "Sisipkan judul", "Format", "Sisipkan judul/heading bergaya Heading 1-3.")
                    .P("text", "string", "teks judul", true)
                    .P("level", "integer", "1-3", false, null, "1")
                    .K("judul", "heading", "subjudul", "bab"),

                T("apply_style", "Terapkan style", "Format", "Terapkan style Word (Title, Heading 1, Heading 2, Normal, Quote, List Paragraph, Caption, dll).")
                    .P("style", "string", "nama style", true)
                    .P("scope", "string", "cakupan", false, Scope, "paragraph")
                    .K("style", "gaya", "heading 1", "normal"),

                T("format_font", "Format huruf", "Format", "Ubah huruf: jenis, ukuran, tebal, miring, garis bawah, warna, sorot.")
                    .P("scope", "string", "cakupan", false, Scope, "selection")
                    .P("name", "string", "nama font, mis. Times New Roman")
                    .P("size", "number", "ukuran pt")
                    .P("bold", "boolean", "tebal").P("italic", "boolean", "miring").P("underline", "boolean", "garis bawah")
                    .P("color", "string", "warna: nama (merah/biru/hitam) atau #RRGGBB")
                    .P("highlight", "string", "warna stabilo: kuning/hijau/biru/merah/none")
                    .K("font", "huruf", "tebal", "bold", "miring", "italic", "garis bawah", "ukuran", "warna", "times new roman", "arial", "stabilo"),

                T("format_paragraph", "Format paragraf", "Format", "Ubah perataan, spasi baris, jarak, dan indentasi paragraf.")
                    .P("scope", "string", "cakupan", false, Scope, "selection")
                    .P("alignment", "string", "perataan", false, new[] { "left", "center", "right", "justify" })
                    .P("line_spacing", "number", "spasi baris: 1, 1.15, 1.5, 2")
                    .P("space_before", "number", "jarak sebelum (pt)").P("space_after", "number", "jarak sesudah (pt)")
                    .P("first_line_indent_cm", "number", "indentasi baris pertama (cm)")
                    .P("left_indent_cm", "number", "indentasi kiri (cm)")
                    .P("hanging_indent_cm", "number", "indentasi gantung (cm), untuk daftar pustaka")
                    .K("spasi", "rata", "justify", "rata kiri kanan", "tengah", "indentasi", "menjorok", "paragraf"),

                T("page_setup", "Atur halaman", "Format", "Atur ukuran kertas, orientasi, dan margin (cm).")
                    .P("paper", "string", "kertas", false, new[] { "A4", "F4", "Letter", "Legal", "A5" })
                    .P("orientation", "string", "orientasi", false, new[] { "portrait", "landscape" })
                    .P("margin_top_cm", "number", "atas").P("margin_bottom_cm", "number", "bawah")
                    .P("margin_left_cm", "number", "kiri").P("margin_right_cm", "number", "kanan")
                    .K("margin", "kertas", "a4", "f4", "folio", "landscape", "portrait", "orientasi", "ukuran halaman"),

                T("apply_academic_format", "Format akademik", "Akademik", "Terapkan format baku sekaligus (font, spasi, margin, perataan, heading).")
                    .P("preset", "string", "preset", true, new[] { "Skripsi", "Makalah", "Jurnal", "SuratDinas", "Laporan", "Proposal" })
                    .P("font", "string", "font", false, null, "Times New Roman")
                    .P("size", "number", "ukuran", false, null, "12")
                    .P("line_spacing", "number", "spasi", false, null, "2")
                    .K("format skripsi", "format makalah", "format baku", "aturan penulisan", "pedoman", "4 3 3 3", "4433"),

                T("change_case", "Ubah huruf besar/kecil", "Format", "Ubah kapitalisasi teks.")
                    .P("case", "string", "jenis", true, new[] { "upper", "lower", "title", "sentence" })
                    .P("scope", "string", "cakupan", false, Scope, "selection")
                    .K("kapital", "huruf besar", "huruf kecil", "uppercase", "lowercase", "kapitalisasi"),

                T("set_columns", "Kolom koran", "Format", "Bagi teks menjadi beberapa kolom (gaya jurnal/koran).")
                    .P("count", "integer", "jumlah kolom", true).P("spacing_cm", "number", "jarak antarkolom")
                    .P("scope", "string", "cakupan", false, new[] { "selection", "document" }, "document")
                    .K("kolom", "dua kolom", "columns", "gaya jurnal"),

                T("clean_document", "Rapikan dokumen", "Format", "Bersihkan spasi ganda, paragraf kosong, enter manual, tab, dan tanda kutip.")
                    .P("double_spaces", "boolean", "", false, null, "true")
                    .P("empty_paragraphs", "boolean", "", false, null, "true")
                    .P("manual_line_breaks", "boolean", "", false, null, "false")
                    .P("tabs", "boolean", "", false, null, "false")
                    .P("smart_quotes", "boolean", "", false, null, "false")
                    .K("rapikan", "bersihkan", "spasi ganda", "baris kosong", "clean", "hasil copy", "pdf"),

                // ===================== Struktur =====================
                T("heading_numbering", "Penomoran bab", "Struktur", "Pasang penomoran bertingkat otomatis pada Heading 1-3.")
                    .P("scheme", "string", "skema", true, new[] { "bab_romawi", "bab_angka", "desimal", "none" })
                    .K("penomoran bab", "bab i", "nomor judul", "multilevel", "1.1"),

                T("insert_toc", "Daftar isi", "Struktur", "Sisipkan daftar isi otomatis dari heading.")
                    .P("levels", "integer", "jumlah level", false, null, "3")
                    .P("title", "string", "judul", false, null, "DAFTAR ISI")
                    .K("daftar isi", "toc", "table of contents"),

                T("insert_table_of_figures", "Daftar gambar/tabel", "Struktur", "Sisipkan daftar gambar atau daftar tabel dari caption.")
                    .P("label", "string", "label caption", true, new[] { "Gambar", "Tabel", "Lampiran" })
                    .P("title", "string", "judul")
                    .K("daftar gambar", "daftar tabel", "daftar lampiran"),

                T("update_fields", "Perbarui field", "Struktur", "Perbarui daftar isi, nomor caption, dan referensi silang.")
                    .K("perbarui daftar isi", "update", "refresh", "perbarui nomor"),

                T("insert_page_numbers", "Nomor halaman", "Struktur", "Sisipkan nomor halaman.")
                    .P("position", "string", "posisi", false, new[] { "bottom_center", "bottom_right", "top_right", "top_center" }, "bottom_center")
                    .P("format", "string", "format", false, new[] { "arabic", "roman_lower", "roman_upper" }, "arabic")
                    .P("start_at", "integer", "mulai dari")
                    .P("scope", "string", "cakupan", false, new[] { "document", "current_section" }, "document")
                    .K("nomor halaman", "page number", "halaman romawi"),

                T("set_header_footer", "Header & footer", "Struktur", "Isi teks header/footer.")
                    .P("header", "string", "teks header").P("footer", "string", "teks footer")
                    .P("different_first_page", "boolean", "halaman pertama berbeda")
                    .K("header", "footer", "kepala halaman", "kaki halaman"),

                T("insert_break", "Pemisah", "Struktur", "Sisipkan pemisah halaman/bagian/kolom.")
                    .P("type", "string", "jenis", true, new[] { "page", "section_next", "section_continuous", "column" })
                    .K("halaman baru", "page break", "section break", "pemisah"),

                T("insert_list", "Daftar poin/nomor", "Struktur", "Sisipkan daftar berpoin atau bernomor.")
                    .P("items", "array", "isi daftar", true)
                    .P("style", "string", "gaya", false, new[] { "bullet", "number", "letter", "roman" }, "bullet")
                    .K("poin", "bullet", "daftar", "penomoran", "list"),

                // ===================== Tabel & objek =====================
                T("insert_table", "Sisipkan tabel", "Tabel", "Sisipkan tabel. rows = array baris, baris pertama = header.")
                    .P("rows", "array", "data tabel [[..],[..]]", true)
                    .P("header", "boolean", "baris pertama header", false, null, "true")
                    .P("style", "string", "gaya tabel", false, new[] { "grid", "akademik", "modern", "polos" }, "grid")
                    .P("caption", "string", "caption tabel (opsional)")
                    .P("autofit", "string", "lebar", false, new[] { "content", "window" }, "window")
                    .K("tabel", "table"),

                T("format_table", "Format tabel", "Tabel", "Format tabel yang dipilih/semua tabel.")
                    .P("scope", "string", "cakupan", false, new[] { "selected", "all" }, "selected")
                    .P("style", "string", "gaya", false, new[] { "grid", "akademik", "modern", "polos" }, "akademik")
                    .P("header_fill", "string", "warna isi header")
                    .P("font_size", "number", "ukuran huruf")
                    .P("autofit", "string", "lebar", false, new[] { "content", "window" })
                    .P("repeat_header", "boolean", "ulangi header di tiap halaman")
                    .K("format tabel", "rapikan tabel", "garis tabel", "border"),

                T("convert_text_to_table", "Teks jadi tabel", "Tabel", "Ubah teks terpilih menjadi tabel.")
                    .P("separator", "string", "pemisah", false, new[] { "tab", "comma", "semicolon", "pipe" }, "tab")
                    .K("ubah jadi tabel", "teks ke tabel"),

                T("insert_caption", "Caption", "Tabel", "Sisipkan caption bernomor otomatis (Gambar 1.1, Tabel 2).")
                    .P("label", "string", "label", true, new[] { "Gambar", "Tabel", "Grafik", "Persamaan", "Lampiran" })
                    .P("text", "string", "keterangan", true)
                    .P("position", "string", "posisi", false, new[] { "above", "below" }, "below")
                    .K("caption", "keterangan gambar", "judul tabel"),

                T("insert_image", "Sisipkan gambar", "Tabel", "Sisipkan gambar dari file.")
                    .P("path", "string", "path file", true).P("width_cm", "number", "lebar cm")
                    .P("caption", "string", "caption")
                    .K("gambar", "foto", "image", "logo"),

                T("insert_equation", "Persamaan", "Tabel", "Sisipkan persamaan matematika (format linear Word, mis. x=(-b±√(b^2-4ac))/2a).")
                    .P("linear", "string", "rumus format linear", true)
                    .K("persamaan", "rumus matematika", "equation"),

                T("insert_footnote", "Catatan kaki", "Akademik", "Sisipkan catatan kaki di posisi kursor.")
                    .P("text", "string", "isi catatan kaki", true)
                    .K("catatan kaki", "footnote"),

                // ===================== Akademik =====================
                T("insert_citation", "Sitasi", "Akademik", "Sisipkan sitasi dalam teks dari pustaka tersimpan atau data yang diberikan.")
                    .P("reference", "object", "{authors:[..], year, title, ...} atau {id}", true)
                    .P("page", "string", "halaman")
                    .P("style", "string", "gaya", false, new[] { "APA7", "IEEE", "Harvard", "MLA", "Chicago", "Vancouver" })
                    .K("sitasi", "kutipan", "citation", "cite"),

                T("insert_bibliography", "Daftar pustaka", "Akademik", "Sisipkan daftar pustaka terurut dengan indentasi gantung.")
                    .P("style", "string", "gaya", false, new[] { "APA7", "IEEE", "Harvard", "MLA", "Chicago", "Vancouver" })
                    .P("references", "array", "daftar pustaka [{type, authors, year, title, journal, volume, issue, pages, publisher, city, url, doi}]")
                    .P("title", "string", "judul", false, null, "DAFTAR PUSTAKA")
                    .K("daftar pustaka", "referensi", "bibliografi", "bibliography", "apa"),

                T("format_bibliography", "Rapikan daftar pustaka", "Akademik", "Urutkan abjad paragraf terpilih dan beri indentasi gantung 1,27 cm.")
                    .K("urutkan pustaka", "rapikan daftar pustaka", "indentasi gantung"),

                T("check_spelling", "Cek ejaan EYD", "Akademik", "Periksa ejaan & kata tidak baku bahasa Indonesia (EYD), beri komentar/sorot.")
                    .P("scope", "string", "cakupan", false, new[] { "selection", "document" }, "document")
                    .P("auto_fix", "boolean", "langsung perbaiki kata tidak baku", false, null, "false")
                    .K("ejaan", "eyd", "puebi", "typo", "kata baku", "tata bahasa", "salah ketik"),

                // ===================== Kantor =====================
                T("insert_letterhead", "Kop surat", "Kantor", "Sisipkan kop surat di header halaman pertama.")
                    .P("institution", "string", "nama instansi", true)
                    .P("subtitle", "string", "sub-instansi/induk")
                    .P("address", "string", "alamat").P("contact", "string", "telepon/email/web")
                    .P("logo_path", "string", "path logo")
                    .K("kop", "kop surat", "letterhead"),

                T("insert_signature", "Blok tanda tangan", "Kantor", "Sisipkan blok tempat, tanggal, jabatan, nama, NIP.")
                    .P("city", "string", "kota").P("date", "string", "tanggal (kosong = hari ini)")
                    .P("position_title", "string", "jabatan").P("name", "string", "nama")
                    .P("id_label", "string", "label nomor", false, new[] { "NIP", "NIK", "NIM", "NRP", "" }, "NIP")
                    .P("id_number", "string", "nomor")
                    .P("align", "string", "letak", false, new[] { "right", "left", "center", "two_columns" }, "right")
                    .P("second_title", "string", "jabatan kedua (two_columns)").P("second_name", "string", "nama kedua")
                    .K("tanda tangan", "ttd", "mengetahui", "menyetujui"),

                T("insert_date", "Sisipkan tanggal", "Kantor", "Sisipkan tanggal format Indonesia.")
                    .P("format", "string", "format", false, new[] { "panjang", "pendek", "hari_tanggal" }, "panjang")
                    .K("tanggal", "hari ini"),

                T("insert_watermark", "Watermark", "Kantor", "Tambahkan watermark teks (RAHASIA, DRAFT, CONTOH).")
                    .P("text", "string", "teks", true)
                    .K("watermark", "draft", "rahasia", "tanda air"),

                T("mail_merge", "Surat massal", "Kantor", "Gabungkan dokumen ini (berisi {{Nama}}, {{Alamat}}, …) dengan data Excel menjadi dokumen baru per baris.")
                    .P("data_file", "string", "path file .xlsx/.csv", true)
                    .P("sheet", "string", "nama sheet")
                    .K("mail merge", "surat massal", "sertifikat massal", "undangan banyak"),

                T("find_replace", "Cari & ganti", "Edit", "Cari dan ganti teks di seluruh dokumen.")
                    .P("find", "string", "teks dicari", true).P("replace", "string", "pengganti", true)
                    .P("match_case", "boolean", "").P("whole_word", "boolean", "").P("wildcards", "boolean", "")
                    .K("ganti", "cari", "replace", "find"),

                T("highlight_text", "Sorot kata", "Edit", "Sorot semua kemunculan kata/frasa.")
                    .P("find", "string", "kata", true).P("color", "string", "warna", false, new[] { "kuning", "hijau", "biru", "merah", "pink" }, "kuning")
                    .K("sorot", "highlight", "tandai"),

                T("add_comment", "Komentar", "Edit", "Tambahkan komentar pada teks terpilih.")
                    .P("text", "string", "isi komentar", true)
                    .K("komentar", "comment", "catatan"),

                T("track_changes", "Lacak perubahan", "Edit", "Hidupkan/matikan Track Changes.")
                    .P("on", "boolean", "", true)
                    .K("track changes", "lacak perubahan", "revisi"),

                T("insert_hyperlink", "Tautan", "Edit", "Sisipkan tautan.")
                    .P("text", "string", "teks", true).P("url", "string", "alamat", true)
                    .K("link", "tautan", "hyperlink"),

                T("export_pdf", "Simpan PDF", "Berkas", "Ekspor dokumen ke PDF.")
                    .P("path", "string", "path file (kosong = sama dengan dokumen)")
                    .K("pdf", "ekspor", "export"),

                T("document_stats", "Statistik dokumen", "Berkas", "Hitung kata, karakter, paragraf, halaman, dan perkiraan waktu baca.")
                    .K("hitung kata", "jumlah kata", "statistik", "word count")
                    .Ro(),
            };
        }
    }
}
