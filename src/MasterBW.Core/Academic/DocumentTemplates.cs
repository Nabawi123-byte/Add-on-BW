using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using MasterBW.Core.Indonesia;

namespace MasterBW.Core.Academic
{
    public sealed class TemplateField
    {
        public string Key;
        public string Label;
        public string Default;
    }

    /// <summary>Templat dokumen Word. Teks memakai {{kunci}}; blok "ai" diisi AI bila diizinkan.</summary>
    public sealed class DocumentTemplate
    {
        public string Id;
        public string Name;
        public string Category;
        public string Description;
        public string Preset = "SuratDinas";
        public List<TemplateField> Fields = new List<TemplateField>();
        public Func<List<DocBlock>> Blocks;

        /// <summary>Bangun blok final dengan isian. Kunci yang tidak diisi menjadi [Label].</summary>
        public List<DocBlock> Render(IDictionary<string, string> values)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var now = DateTime.Now;
            map["tanggal"] = IndoDate.Long(now);
            map["hari_tanggal"] = IndoDate.WithDay(now);
            map["tahun"] = now.Year.ToString();
            map["bulan_romawi"] = IndoDate.RomanMonth(now);
            foreach (var f in Fields) if (!string.IsNullOrEmpty(f.Default)) map[f.Key] = f.Default;
            if (values != null) foreach (var kv in values) if (!string.IsNullOrWhiteSpace(kv.Value)) map[kv.Key] = kv.Value;

            string Fill(string s)
            {
                if (s == null) return null;
                return Regex.Replace(s, @"\{\{(\w+)\}\}", m =>
                {
                    var key = m.Groups[1].Value;
                    if (map.TryGetValue(key, out var v)) return v;
                    var label = Fields.FirstOrDefault(f => string.Equals(f.Key, key, StringComparison.OrdinalIgnoreCase))?.Label ?? key;
                    return "[" + label + "]";
                });
            }

            var result = new List<DocBlock>();
            foreach (var b in Blocks())
            {
                b.Text = Fill(b.Text);
                if (b.Items != null) b.Items = b.Items.Select(Fill).ToList();
                if (b.Rows != null) b.Rows = b.Rows.Select(r => r.Select(Fill).ToList()).ToList();
                if (b.Data != null) b.Data = b.Data.ToDictionary(k => k.Key, k => Fill(k.Value));
                result.Add(b);
            }
            return result;
        }
    }

    /// <summary>Pustaka templat dokumen bawaan: perkantoran, administrasi pemerintahan/sekolah, akademik, pendidikan.</summary>
    public static class DocumentTemplates
    {
        // ---------------------------------------------------------------- DSL
        private static DocBlock Title(string t) => DocBlock.Of("title", t);
        private static DocBlock H(string t, int level = 1) => DocBlock.Of("heading", t, level);
        private static DocBlock P(string t) => DocBlock.Of("paragraph", t);
        private static DocBlock Center(string t, bool bold = false) => new DocBlock { Type = "center", Text = t, Bold = bold };
        private static DocBlock Right(string t) => DocBlock.Of("right", t);
        private static DocBlock Ai(string instruction) => DocBlock.Of("ai", instruction);
        private static DocBlock Bul(params string[] items) => new DocBlock { Type = "bullets", Items = items.ToList() };
        private static DocBlock Num(params string[] items) => new DocBlock { Type = "numbered", Items = items.ToList() };
        private static DocBlock KV(params string[] pairs) => new DocBlock { Type = "keyvalue", Rows = pairs.Select(p => p.Split(new[] { '|' }, 2).ToList()).ToList() };
        private static DocBlock Table(params string[] rows) => new DocBlock { Type = "table", Rows = rows.Select(r => r.Split('|').ToList()).ToList() };
        private static DocBlock Space() => DocBlock.Of("spacer");
        private static DocBlock Break() => DocBlock.Of("page_break");
        private static DocBlock Toc() => DocBlock.Of("toc");
        private static DocBlock Kop() => DocBlock.Of("letterhead");

        private static DocBlock Sign(string title = "{{jabatan}}", string name = "{{nama_penandatangan}}", string idLabel = "NIP", string idNum = "{{nip}}", string align = "right")
            => new DocBlock { Type = "signature", Data = new Dictionary<string, string> { ["city"] = "{{kota}}", ["date"] = "{{tanggal}}", ["title"] = title, ["name"] = name, ["id_label"] = idLabel, ["id_number"] = idNum, ["align"] = align } };

        private static DocBlock Sign2(string leftTitle, string leftName, string rightTitle, string rightName)
            => new DocBlock { Type = "signature", Data = new Dictionary<string, string> { ["city"] = "{{kota}}", ["date"] = "{{tanggal}}", ["title"] = rightTitle, ["name"] = rightName, ["second_title"] = leftTitle, ["second_name"] = leftName, ["align"] = "two_columns", ["id_label"] = "NIP", ["id_number"] = "" } };

        private static TemplateField F(string key, string label, string def = null) => new TemplateField { Key = key, Label = label, Default = def };

        private static List<TemplateField> Office(params TemplateField[] extra)
        {
            var list = new List<TemplateField>
            {
                F("instansi", "Nama instansi"), F("alamat_instansi", "Alamat instansi"), F("kota", "Kota"),
                F("jabatan", "Jabatan penandatangan"), F("nama_penandatangan", "Nama penandatangan"), F("nip", "NIP penandatangan"),
            };
            list.AddRange(extra);
            return list;
        }

        // ---------------------------------------------------------------- daftar templat
        public static readonly List<DocumentTemplate> All = new List<DocumentTemplate>
        {
            // ======================= PERKANTORAN / ADMINISTRASI =======================
            new DocumentTemplate
            {
                Id = "surat_dinas", Name = "Surat Dinas Resmi", Category = "Surat", Description = "Surat resmi instansi dengan kop, nomor, lampiran, perihal.",
                Fields = Office(F("nomor", "Nomor surat", "{{nomor_urut}}/[kode]/{{bulan_romawi}}/{{tahun}}"), F("nomor_urut", "Nomor urut", "001"), F("lampiran", "Lampiran", "-"), F("perihal", "Perihal"), F("tujuan", "Kepada (Yth.)"), F("tempat_tujuan", "Tempat tujuan", "Tempat"), F("isi", "Pokok isi surat")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Right("{{kota}}, {{tanggal}}"),
                    KV("Nomor|: {{nomor}}", "Lampiran|: {{lampiran}}", "Perihal|: {{perihal}}"), Space(),
                    P("Yth. {{tujuan}}"), P("di {{tempat_tujuan}}"), Space(),
                    P("Dengan hormat,"),
                    Ai("Tulis isi surat dinas resmi tentang: {{perihal}}. Pokok isi: {{isi}}. 1-3 paragraf, bahasa baku, tanpa salam pembuka/penutup."),
                    P("Demikian surat ini kami sampaikan. Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih."),
                    Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "surat_undangan", Name = "Surat Undangan", Category = "Surat", Description = "Undangan rapat/acara dengan rincian hari, waktu, tempat, acara.",
                Fields = Office(F("nomor", "Nomor surat"), F("perihal", "Perihal", "Undangan"), F("tujuan", "Kepada (Yth.)", "Bapak/Ibu [nama]"), F("hari_acara", "Hari, tanggal acara"), F("waktu", "Waktu", "09.00 WIB s.d. selesai"), F("tempat", "Tempat"), F("acara", "Acara")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Right("{{kota}}, {{tanggal}}"),
                    KV("Nomor|: {{nomor}}", "Lampiran|: -", "Perihal|: {{perihal}}"), Space(),
                    P("Yth. {{tujuan}}"), P("di Tempat"), Space(), P("Dengan hormat,"),
                    P("Sehubungan dengan akan dilaksanakannya {{acara}}, kami mengharapkan kehadiran Bapak/Ibu pada:"),
                    KV("Hari, tanggal|: {{hari_acara}}", "Waktu|: {{waktu}}", "Tempat|: {{tempat}}", "Acara|: {{acara}}"),
                    P("Mengingat pentingnya acara tersebut, kami mohon Bapak/Ibu hadir tepat waktu. Atas perhatian dan kehadirannya, kami ucapkan terima kasih."),
                    Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "surat_keputusan", Name = "Surat Keputusan (SK)", Category = "Surat", Description = "SK lengkap: menimbang, mengingat, memutuskan, diktum.",
                Fields = Office(F("nomor", "Nomor SK"), F("tentang", "Tentang"), F("pejabat", "Pejabat penetap (mis. KEPALA SEKOLAH ...)")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Center("KEPUTUSAN {{pejabat}}", true), Center("NOMOR {{nomor}}", true), Space(), Center("TENTANG", true), Center("{{tentang}}", true), Space(),
                    Center("{{pejabat}},", true),
                    KV("Menimbang|: a. bahwa [pertimbangan pertama];", "|  b. bahwa [pertimbangan kedua];", "|  c. bahwa berdasarkan pertimbangan sebagaimana dimaksud dalam huruf a dan huruf b, perlu menetapkan Keputusan tentang {{tentang}};"),
                    KV("Mengingat|: 1. Undang-Undang Nomor [..] Tahun [..] tentang [..];", "|  2. Peraturan [..] Nomor [..] Tahun [..] tentang [..];"),
                    Center("MEMUTUSKAN:", true),
                    KV("Menetapkan|: KEPUTUSAN {{pejabat}} TENTANG {{tentang}}.", "KESATU|: [isi diktum kesatu];", "KEDUA|: [isi diktum kedua];", "KETIGA|: Keputusan ini mulai berlaku pada tanggal ditetapkan, dengan ketentuan apabila di kemudian hari terdapat kekeliruan akan diadakan perbaikan sebagaimana mestinya."),
                    new DocBlock { Type = "signature", Data = new Dictionary<string, string> { ["city"] = "Ditetapkan di {{kota}}", ["date"] = "pada tanggal {{tanggal}}", ["title"] = "{{jabatan}}", ["name"] = "{{nama_penandatangan}}", ["id_label"] = "NIP", ["id_number"] = "{{nip}}", ["align"] = "right", ["raw_place"] = "1" } },
                },
            },
            new DocumentTemplate
            {
                Id = "surat_tugas", Name = "Surat Tugas", Category = "Surat", Description = "Penugasan pegawai/guru untuk kegiatan.",
                Fields = Office(F("nomor", "Nomor surat"), F("nama_petugas", "Nama yang ditugaskan"), F("nip_petugas", "NIP yang ditugaskan"), F("jabatan_petugas", "Jabatan yang ditugaskan"), F("tugas", "Uraian tugas"), F("waktu_tugas", "Waktu pelaksanaan"), F("tempat", "Tempat")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Center("SURAT TUGAS", true), Center("Nomor: {{nomor}}"), Space(),
                    P("Yang bertanda tangan di bawah ini {{jabatan}} {{instansi}}, menugaskan kepada:"),
                    KV("Nama|: {{nama_petugas}}", "NIP|: {{nip_petugas}}", "Jabatan|: {{jabatan_petugas}}"),
                    P("Untuk {{tugas}} yang dilaksanakan pada:"),
                    KV("Waktu|: {{waktu_tugas}}", "Tempat|: {{tempat}}"),
                    P("Demikian surat tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab."),
                    Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "surat_keterangan", Name = "Surat Keterangan", Category = "Surat", Description = "Keterangan aktif kerja/sekolah/domisili, dll.",
                Fields = Office(F("nomor", "Nomor surat"), F("nama", "Nama yang diterangkan"), F("identitas", "NIK/NIP/NIS"), F("ttl", "Tempat, tanggal lahir"), F("alamat", "Alamat"), F("keterangan", "Isi keterangan")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Center("SURAT KETERANGAN", true), Center("Nomor: {{nomor}}"), Space(),
                    P("Yang bertanda tangan di bawah ini {{jabatan}} {{instansi}} menerangkan bahwa:"),
                    KV("Nama|: {{nama}}", "NIK/NIP/NIS|: {{identitas}}", "Tempat, tanggal lahir|: {{ttl}}", "Alamat|: {{alamat}}"),
                    P("{{keterangan}}"),
                    P("Demikian surat keterangan ini dibuat dengan sebenarnya untuk dipergunakan sebagaimana mestinya."),
                    Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "surat_kuasa", Name = "Surat Kuasa", Category = "Surat", Preset = "SuratDinas", Description = "Pemberian kuasa dari pemberi kepada penerima kuasa.",
                Fields = new List<TemplateField> { F("kota", "Kota"), F("pemberi", "Nama pemberi kuasa"), F("nik_pemberi", "NIK pemberi"), F("alamat_pemberi", "Alamat pemberi"), F("penerima", "Nama penerima kuasa"), F("nik_penerima", "NIK penerima"), F("alamat_penerima", "Alamat penerima"), F("keperluan", "Keperluan kuasa") },
                Blocks = () => new List<DocBlock>
                {
                    Center("SURAT KUASA", true), Space(), P("Yang bertanda tangan di bawah ini:"),
                    KV("Nama|: {{pemberi}}", "NIK|: {{nik_pemberi}}", "Alamat|: {{alamat_pemberi}}"),
                    P("Selanjutnya disebut PEMBERI KUASA, dengan ini memberikan kuasa kepada:"),
                    KV("Nama|: {{penerima}}", "NIK|: {{nik_penerima}}", "Alamat|: {{alamat_penerima}}"),
                    P("Selanjutnya disebut PENERIMA KUASA, untuk {{keperluan}}."),
                    P("Demikian surat kuasa ini dibuat dengan sebenarnya untuk dipergunakan sebagaimana mestinya."),
                    Sign2("Penerima Kuasa,", "{{penerima}}", "Pemberi Kuasa,", "{{pemberi}}"),
                },
            },
            new DocumentTemplate
            {
                Id = "surat_pernyataan", Name = "Surat Pernyataan", Category = "Surat", Description = "Pernyataan pribadi (bermeterai bila perlu).",
                Fields = new List<TemplateField> { F("kota", "Kota"), F("nama", "Nama"), F("identitas", "NIK/NIM/NIP"), F("ttl", "Tempat, tanggal lahir"), F("alamat", "Alamat"), F("pernyataan", "Isi pernyataan") },
                Blocks = () => new List<DocBlock>
                {
                    Center("SURAT PERNYATAAN", true), Space(), P("Yang bertanda tangan di bawah ini:"),
                    KV("Nama|: {{nama}}", "NIK/NIM/NIP|: {{identitas}}", "Tempat, tanggal lahir|: {{ttl}}", "Alamat|: {{alamat}}"),
                    P("Dengan ini menyatakan dengan sesungguhnya bahwa {{pernyataan}}."),
                    P("Demikian surat pernyataan ini saya buat dengan sebenar-benarnya dalam keadaan sadar dan tanpa paksaan dari pihak mana pun. Apabila di kemudian hari pernyataan ini tidak benar, saya bersedia menerima sanksi sesuai ketentuan yang berlaku."),
                    Sign("Yang membuat pernyataan,", "{{nama}}", "", "", "right"),
                },
            },
            new DocumentTemplate
            {
                Id = "surat_permohonan", Name = "Surat Permohonan", Category = "Surat", Description = "Permohonan izin/bantuan/data/kerja sama.",
                Fields = Office(F("nomor", "Nomor"), F("perihal", "Perihal"), F("tujuan", "Kepada (Yth.)"), F("permohonan", "Hal yang dimohon")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Right("{{kota}}, {{tanggal}}"), KV("Nomor|: {{nomor}}", "Lampiran|: -", "Perihal|: {{perihal}}"), Space(),
                    P("Yth. {{tujuan}}"), P("di Tempat"), Space(), P("Dengan hormat,"),
                    Ai("Tulis isi surat permohonan resmi: {{permohonan}}. Sertakan alasan dan harapan, 2 paragraf."),
                    P("Demikian permohonan ini kami sampaikan. Atas perhatian dan perkenan Bapak/Ibu, kami ucapkan terima kasih."), Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "surat_lamaran", Name = "Surat Lamaran Kerja", Category = "Surat", Description = "Lamaran kerja dengan daftar lampiran.",
                Fields = new List<TemplateField> { F("kota", "Kota"), F("nama", "Nama pelamar"), F("ttl", "Tempat, tanggal lahir"), F("pendidikan", "Pendidikan terakhir"), F("alamat", "Alamat"), F("telepon", "No. HP/email"), F("posisi", "Posisi yang dilamar"), F("perusahaan", "Nama perusahaan/HRD"), F("keunggulan", "Pengalaman/keunggulan") },
                Blocks = () => new List<DocBlock>
                {
                    Right("{{kota}}, {{tanggal}}"), P("Perihal: Lamaran Pekerjaan sebagai {{posisi}}"), Space(),
                    P("Yth. Bagian HRD {{perusahaan}}"), P("di Tempat"), Space(), P("Dengan hormat,"),
                    P("Saya yang bertanda tangan di bawah ini:"),
                    KV("Nama|: {{nama}}", "Tempat, tanggal lahir|: {{ttl}}", "Pendidikan|: {{pendidikan}}", "Alamat|: {{alamat}}", "Telepon/Email|: {{telepon}}"),
                    Ai("Tulis 2 paragraf lamaran kerja yang meyakinkan untuk posisi {{posisi}} di {{perusahaan}}, berdasarkan keunggulan: {{keunggulan}}."),
                    P("Sebagai bahan pertimbangan, saya lampirkan:"),
                    Num("Daftar riwayat hidup", "Fotokopi ijazah dan transkrip nilai", "Fotokopi KTP", "Pasfoto terbaru", "Sertifikat pendukung"),
                    P("Besar harapan saya untuk dapat diberi kesempatan wawancara. Atas perhatian Bapak/Ibu, saya ucapkan terima kasih."),
                    Sign("Hormat saya,", "{{nama}}", "", "", "right"),
                },
            },
            new DocumentTemplate
            {
                Id = "nota_dinas", Name = "Nota Dinas / Memo", Category = "Surat", Description = "Komunikasi internal antarunit.",
                Fields = Office(F("nomor", "Nomor"), F("kepada", "Kepada"), F("dari", "Dari"), F("hal", "Hal"), F("isi", "Pokok isi")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Center("NOTA DINAS", true), Center("Nomor: {{nomor}}"), Space(),
                    KV("Yth.|: {{kepada}}", "Dari|: {{dari}}", "Tanggal|: {{tanggal}}", "Hal|: {{hal}}"), Space(),
                    Ai("Tulis isi nota dinas singkat dan lugas tentang {{hal}}: {{isi}}."),
                    P("Demikian disampaikan, atas perhatiannya diucapkan terima kasih."), Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "notulen_rapat", Name = "Notulen Rapat", Category = "Rapat", Description = "Notulen lengkap: identitas rapat, agenda, pembahasan, keputusan, tindak lanjut.",
                Fields = Office(F("acara", "Nama rapat"), F("hari_acara", "Hari, tanggal"), F("waktu", "Waktu"), F("tempat", "Tempat"), F("pimpinan", "Pimpinan rapat"), F("notulis", "Notulis"), F("peserta", "Peserta"), F("agenda", "Agenda"), F("catatan", "Catatan/poin pembahasan mentah")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Center("NOTULEN RAPAT", true), Center("{{acara}}", true), Space(),
                    KV("Hari, tanggal|: {{hari_acara}}", "Waktu|: {{waktu}}", "Tempat|: {{tempat}}", "Pimpinan rapat|: {{pimpinan}}", "Notulis|: {{notulis}}", "Peserta|: {{peserta}}", "Agenda|: {{agenda}}"),
                    H("A. Pembukaan", 2), Ai("Tulis paragraf singkat pembukaan rapat {{acara}} oleh {{pimpinan}}."),
                    H("B. Pembahasan", 2), Ai("Susun poin-poin pembahasan rapat secara rapi dari catatan berikut: {{catatan}}. Gunakan daftar bernomor."),
                    H("C. Keputusan/Kesimpulan", 2), Ai("Tulis keputusan rapat dalam daftar bernomor berdasarkan catatan: {{catatan}}."),
                    H("D. Tindak Lanjut", 2),
                    Table("No|Kegiatan|Penanggung Jawab|Tenggat", "1|[kegiatan]|[PIC]|[tanggal]", "2|[kegiatan]|[PIC]|[tanggal]"),
                    H("E. Penutup", 2), P("Rapat ditutup oleh pimpinan rapat pada pukul [jam] WIB."),
                    Sign2("Notulis,", "{{notulis}}", "Pimpinan Rapat,", "{{pimpinan}}"),
                },
            },
            new DocumentTemplate
            {
                Id = "berita_acara", Name = "Berita Acara", Category = "Rapat", Description = "Berita acara kegiatan/serah terima/pemeriksaan.",
                Fields = Office(F("nomor", "Nomor"), F("judul", "Judul berita acara (mis. SERAH TERIMA BARANG)"), F("hari_acara", "Hari, tanggal"), F("pihak1", "Pihak pertama (nama, jabatan)"), F("pihak2", "Pihak kedua (nama, jabatan)"), F("uraian", "Uraian kejadian")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Center("BERITA ACARA {{judul}}", true), Center("Nomor: {{nomor}}"), Space(),
                    P("Pada hari ini {{hari_acara}}, yang bertanda tangan di bawah ini:"),
                    KV("1. |{{pihak1}}, selanjutnya disebut PIHAK PERTAMA.", "2. |{{pihak2}}, selanjutnya disebut PIHAK KEDUA."),
                    Ai("Tulis isi berita acara {{judul}} berdasarkan uraian: {{uraian}}. Gunakan bahasa resmi dan poin bernomor bila ada rincian."),
                    P("Demikian berita acara ini dibuat dengan sebenarnya dalam rangkap dua untuk dipergunakan sebagaimana mestinya."),
                    Sign2("PIHAK KEDUA,", "[nama]", "PIHAK PERTAMA,", "[nama]"),
                },
            },
            new DocumentTemplate
            {
                Id = "laporan_kegiatan", Name = "Laporan Kegiatan", Category = "Laporan", Preset = "Laporan", Description = "Laporan pertanggungjawaban kegiatan.",
                Fields = Office(F("kegiatan", "Nama kegiatan"), F("waktu_tugas", "Waktu pelaksanaan"), F("tempat", "Tempat"), F("peserta", "Peserta"), F("ringkasan", "Ringkasan pelaksanaan & hasil")),
                Blocks = () => new List<DocBlock>
                {
                    Center("LAPORAN KEGIATAN", true), Center("{{kegiatan}}", true), Center("{{instansi}}"), Center("{{tahun}}"), Break(),
                    H("BAB I PENDAHULUAN"), H("A. Latar Belakang", 2), Ai("Latar belakang kegiatan {{kegiatan}}. Ringkasan: {{ringkasan}}."),
                    H("B. Tujuan", 2), Ai("Tujuan kegiatan {{kegiatan}} dalam daftar bernomor."),
                    H("C. Dasar Pelaksanaan", 2), Num("[Dasar hukum/surat tugas]"),
                    H("BAB II PELAKSANAAN"), KV("Waktu|: {{waktu_tugas}}", "Tempat|: {{tempat}}", "Peserta|: {{peserta}}"),
                    H("A. Rangkaian Kegiatan", 2), Ai("Uraikan rangkaian pelaksanaan kegiatan {{kegiatan}} berdasarkan: {{ringkasan}}."),
                    H("BAB III HASIL DAN EVALUASI"), H("A. Hasil yang Dicapai", 2), Ai("Hasil kegiatan berdasarkan: {{ringkasan}}."),
                    H("B. Kendala dan Solusi", 2), Ai("Kendala yang mungkin dihadapi dan solusinya dalam poin."),
                    H("BAB IV PENUTUP"), H("A. Kesimpulan", 2), Ai("Kesimpulan kegiatan."), H("B. Saran", 2), Ai("Saran untuk kegiatan berikutnya."),
                    Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "proposal_kegiatan", Name = "Proposal Kegiatan", Category = "Laporan", Preset = "Laporan", Description = "Proposal acara/kegiatan lengkap dengan anggaran dan susunan panitia.",
                Fields = Office(F("kegiatan", "Nama kegiatan"), F("tema", "Tema"), F("waktu_tugas", "Waktu"), F("tempat", "Tempat"), F("sasaran", "Sasaran/peserta"), F("ketua", "Ketua panitia")),
                Blocks = () => new List<DocBlock>
                {
                    Center("PROPOSAL KEGIATAN", true), Center("{{kegiatan}}", true), Center("\"{{tema}}\""), Center("{{instansi}}"), Center("{{tahun}}"), Break(),
                    H("A. Latar Belakang", 2), Ai("Latar belakang proposal kegiatan {{kegiatan}} bertema {{tema}}."),
                    H("B. Nama dan Tema Kegiatan", 2), KV("Nama kegiatan|: {{kegiatan}}", "Tema|: {{tema}}"),
                    H("C. Tujuan", 2), Ai("Tujuan kegiatan dalam daftar bernomor."),
                    H("D. Sasaran", 2), P("{{sasaran}}"),
                    H("E. Waktu dan Tempat", 2), KV("Waktu|: {{waktu_tugas}}", "Tempat|: {{tempat}}"),
                    H("F. Susunan Acara", 2), Table("Waktu|Acara|Penanggung Jawab", "[jam]|Pembukaan|[PIC]", "[jam]|[acara inti]|[PIC]", "[jam]|Penutup|[PIC]"),
                    H("G. Susunan Panitia", 2), Table("Jabatan|Nama", "Penanggung jawab|{{nama_penandatangan}}", "Ketua|{{ketua}}", "Sekretaris|[nama]", "Bendahara|[nama]"),
                    H("H. Rencana Anggaran", 2), Table("No|Uraian|Volume|Harga Satuan|Jumlah", "1|[uraian]|[vol]|[harga]|[jumlah]", "|TOTAL|||[total]"),
                    H("I. Penutup", 2), Ai("Paragraf penutup proposal yang mengharapkan dukungan."),
                    Sign2("Ketua Panitia,", "{{ketua}}", "Mengetahui, {{jabatan}}", "{{nama_penandatangan}}"),
                },
            },
            new DocumentTemplate
            {
                Id = "kak_tor", Name = "Kerangka Acuan Kerja (KAK/TOR)", Category = "Laporan", Preset = "Laporan", Description = "KAK/TOR untuk kegiatan/pengadaan instansi pemerintah.",
                Fields = Office(F("kegiatan", "Nama kegiatan"), F("anggaran", "Total anggaran"), F("waktu_tugas", "Jangka waktu")),
                Blocks = () => new List<DocBlock>
                {
                    Center("KERANGKA ACUAN KERJA (KAK)", true), Center("{{kegiatan}}", true), Space(),
                    H("A. Latar Belakang", 2), Ai("Latar belakang KAK {{kegiatan}} (dasar hukum, gambaran umum, alasan kegiatan)."),
                    H("B. Maksud dan Tujuan", 2), Ai("Maksud dan tujuan kegiatan {{kegiatan}}."),
                    H("C. Sasaran dan Indikator Keluaran", 2), Ai("Sasaran, indikator keluaran (output) dan hasil (outcome) dalam poin."),
                    H("D. Metode Pelaksanaan", 2), Ai("Metode pelaksanaan (swakelola/kontraktual) dan tahapan kegiatan."),
                    H("E. Jadwal Pelaksanaan", 2), P("Kegiatan dilaksanakan selama {{waktu_tugas}}."),
                    H("F. Biaya", 2), P("Total biaya yang diperlukan sebesar {{anggaran}} dibebankan pada DIPA/DPA {{instansi}} Tahun Anggaran {{tahun}}."),
                    Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "sop", Name = "Standar Operasional Prosedur (SOP)", Category = "Laporan", Description = "SOP dengan tujuan, ruang lingkup, definisi, prosedur.",
                Fields = Office(F("judul", "Judul SOP"), F("nomor", "Nomor SOP"), F("proses", "Gambaran proses")),
                Blocks = () => new List<DocBlock>
                {
                    Kop(), Table("Nomor SOP|{{nomor}}", "Tanggal pembuatan|{{tanggal}}", "Disahkan oleh|{{jabatan}}", "Nama SOP|{{judul}}"),
                    H("1. Tujuan", 2), Ai("Tujuan SOP {{judul}}."), H("2. Ruang Lingkup", 2), Ai("Ruang lingkup SOP {{judul}}."),
                    H("3. Definisi", 2), Ai("Definisi istilah penting pada SOP {{judul}} dalam poin."),
                    H("4. Prosedur", 2), Ai("Langkah-langkah prosedur bernomor untuk {{judul}}: {{proses}}."),
                    H("5. Dokumen Terkait", 2), Bul("[dokumen/formulir terkait]"), Sign(),
                },
            },
            new DocumentTemplate
            {
                Id = "cv", Name = "Daftar Riwayat Hidup (CV)", Category = "Pribadi", Description = "CV formal format Indonesia.",
                Fields = new List<TemplateField> { F("kota", "Kota"), F("nama", "Nama lengkap"), F("ttl", "Tempat, tanggal lahir"), F("jk", "Jenis kelamin"), F("agama", "Agama"), F("alamat", "Alamat"), F("telepon", "No. HP"), F("email", "Email"), F("ringkasan", "Ringkasan profil/pengalaman") },
                Blocks = () => new List<DocBlock>
                {
                    Center("DAFTAR RIWAYAT HIDUP", true), Space(), H("A. Data Pribadi", 2),
                    KV("Nama lengkap|: {{nama}}", "Tempat, tanggal lahir|: {{ttl}}", "Jenis kelamin|: {{jk}}", "Agama|: {{agama}}", "Alamat|: {{alamat}}", "No. HP|: {{telepon}}", "Email|: {{email}}"),
                    H("B. Profil Singkat", 2), Ai("Ringkasan profil profesional 3-4 kalimat dari: {{ringkasan}}."),
                    H("C. Riwayat Pendidikan", 2), Table("Tahun|Jenjang|Institusi|Jurusan", "[tahun]|[jenjang]|[institusi]|[jurusan]"),
                    H("D. Pengalaman Kerja/Organisasi", 2), Table("Tahun|Posisi|Instansi|Keterangan", "[tahun]|[posisi]|[instansi]|[keterangan]"),
                    H("E. Keahlian", 2), Bul("[keahlian 1]", "[keahlian 2]"),
                    P("Demikian daftar riwayat hidup ini saya buat dengan sebenarnya."), Sign("Hormat saya,", "{{nama}}", "", "", "right"),
                },
            },

            // ======================= AKADEMIK =======================
            new DocumentTemplate
            {
                Id = "makalah", Name = "Makalah", Category = "Akademik", Preset = "Makalah", Description = "Makalah lengkap: sampul, kata pengantar, daftar isi, BAB I–III, daftar pustaka.",
                Fields = new List<TemplateField> { F("judul", "Judul makalah"), F("mata_kuliah", "Mata kuliah/pelajaran"), F("dosen", "Dosen/guru pengampu"), F("penulis", "Nama penulis (pisahkan koma)"), F("nim", "NIM/NIS"), F("prodi", "Program studi/kelas"), F("institusi", "Kampus/sekolah"), F("kota", "Kota") },
                Blocks = () => new List<DocBlock>
                {
                    new DocBlock { Type = "cover", Data = new Dictionary<string, string> { ["kind"] = "MAKALAH", ["title"] = "{{judul}}", ["subtitle"] = "Disusun untuk memenuhi tugas mata kuliah {{mata_kuliah}}\nDosen pengampu: {{dosen}}", ["author"] = "{{penulis}}", ["id_number"] = "{{nim}}", ["program"] = "{{prodi}}", ["institution"] = "{{institusi}}", ["city"] = "{{kota}}", ["year"] = "{{tahun}}" } },
                    Center("KATA PENGANTAR", true), Ai("Kata pengantar makalah berjudul {{judul}} untuk mata kuliah {{mata_kuliah}}, ucapan syukur, terima kasih kepada {{dosen}}, harapan dan permohonan saran. 3 paragraf."),
                    Right("{{kota}}, {{tanggal}}"), Right("Penulis"), Break(),
                    Toc(), Break(),
                    H("BAB I PENDAHULUAN"), H("1.1 Latar Belakang", 2), Ai("Latar belakang makalah tentang {{judul}}, 3-4 paragraf, akademik."),
                    H("1.2 Rumusan Masalah", 2), Ai("3 rumusan masalah berbentuk pertanyaan bernomor untuk makalah {{judul}}."),
                    H("1.3 Tujuan", 2), Ai("Tujuan penulisan yang sejalan dengan rumusan masalah, bernomor."),
                    H("1.4 Manfaat", 2), Ai("Manfaat teoretis dan praktis makalah {{judul}}."), Break(),
                    H("BAB II PEMBAHASAN"), H("2.1 [Subbab pertama]", 2), Ai("Pembahasan mendalam bagian pertama makalah {{judul}} (pengertian/konsep), 4 paragraf."),
                    H("2.2 [Subbab kedua]", 2), Ai("Pembahasan bagian kedua makalah {{judul}} (pembahasan utama), 4 paragraf."),
                    H("2.3 [Subbab ketiga]", 2), Ai("Pembahasan bagian ketiga makalah {{judul}} (contoh/penerapan/analisis), 3 paragraf."), Break(),
                    H("BAB III PENUTUP"), H("3.1 Kesimpulan", 2), Ai("Kesimpulan makalah {{judul}} yang menjawab rumusan masalah."),
                    H("3.2 Saran", 2), Ai("Saran yang relevan dari makalah {{judul}}."), Break(),
                    Center("DAFTAR PUSTAKA", true), P("[Tambahkan referensi lewat menu Daftar Pustaka MASTER BW]"),
                },
            },
            new DocumentTemplate
            {
                Id = "skripsi", Name = "Kerangka Skripsi", Category = "Akademik", Preset = "Skripsi", Description = "Kerangka skripsi/tugas akhir: halaman awal, BAB I–V, daftar pustaka, lampiran.",
                Fields = new List<TemplateField> { F("judul", "Judul skripsi"), F("penulis", "Nama mahasiswa"), F("nim", "NIM"), F("prodi", "Program studi"), F("fakultas", "Fakultas"), F("institusi", "Universitas"), F("kota", "Kota"), F("pembimbing", "Dosen pembimbing"), F("metode", "Metode penelitian (kuantitatif/kualitatif/R&D/PTK)") },
                Blocks = () => new List<DocBlock>
                {
                    new DocBlock { Type = "cover", Data = new Dictionary<string, string> { ["kind"] = "SKRIPSI", ["title"] = "{{judul}}", ["subtitle"] = "Diajukan untuk memenuhi sebagian persyaratan memperoleh gelar Sarjana", ["author"] = "{{penulis}}", ["id_number"] = "{{nim}}", ["program"] = "PROGRAM STUDI {{prodi}}\nFAKULTAS {{fakultas}}", ["institution"] = "{{institusi}}", ["city"] = "{{kota}}", ["year"] = "{{tahun}}" } },
                    Center("LEMBAR PERSETUJUAN PEMBIMBING", true), P("Skripsi berjudul \"{{judul}}\" yang disusun oleh {{penulis}} (NIM {{nim}}) telah disetujui untuk diujikan."),
                    Sign("Dosen Pembimbing,", "{{pembimbing}}", "NIDN", ""), Break(),
                    Center("PERNYATAAN KEASLIAN", true), P("Saya yang bertanda tangan di bawah ini menyatakan bahwa skripsi ini merupakan karya saya sendiri dan bukan plagiasi. Semua sumber yang dikutip telah saya nyatakan dengan benar."),
                    Sign("Yang menyatakan,", "{{penulis}}", "NIM", "{{nim}}"), Break(),
                    Center("ABSTRAK", true), Ai("Abstrak skripsi 200-250 kata berjudul {{judul}} dengan metode {{metode}} (latar belakang, tujuan, metode, hasil [sementara], simpulan) dan Kata kunci."), Break(),
                    Center("KATA PENGANTAR", true), Ai("Kata pengantar skripsi {{judul}}, berterima kasih kepada {{pembimbing}}, orang tua, dan pihak terkait."), Break(),
                    Toc(), Break(),
                    H("BAB I PENDAHULUAN"), H("1.1 Latar Belakang Masalah", 2), Ai("Latar belakang masalah skripsi {{judul}}: kondisi ideal, fakta lapangan, kesenjangan, urgensi. 5 paragraf akademik."),
                    H("1.2 Identifikasi Masalah", 2), Ai("Identifikasi masalah bernomor untuk {{judul}}."),
                    H("1.3 Batasan Masalah", 2), Ai("Batasan masalah penelitian {{judul}}."),
                    H("1.4 Rumusan Masalah", 2), Ai("Rumusan masalah berbentuk pertanyaan untuk {{judul}}."),
                    H("1.5 Tujuan Penelitian", 2), Ai("Tujuan penelitian sesuai rumusan masalah."),
                    H("1.6 Manfaat Penelitian", 2), Ai("Manfaat teoretis dan praktis penelitian {{judul}}."), Break(),
                    H("BAB II KAJIAN PUSTAKA"), H("2.1 Landasan Teori", 2), Ai("Landasan teori untuk variabel/konsep utama penelitian {{judul}}. Tandai rujukan dengan [sumber]."),
                    H("2.2 Penelitian yang Relevan", 2), P("[Uraikan 3–5 penelitian terdahulu: peneliti, tahun, judul, hasil, persamaan & perbedaan]"),
                    H("2.3 Kerangka Berpikir", 2), Ai("Kerangka berpikir penelitian {{judul}}."),
                    H("2.4 Hipotesis Penelitian", 2), Ai("Hipotesis penelitian (H0 dan Ha) untuk {{judul}} jika kuantitatif; jika kualitatif tulis 'Tidak menggunakan hipotesis'. Metode: {{metode}}."), Break(),
                    H("BAB III METODE PENELITIAN"), H("3.1 Jenis dan Desain Penelitian", 2), Ai("Jenis dan desain penelitian {{metode}} untuk {{judul}}."),
                    H("3.2 Tempat dan Waktu Penelitian", 2), P("[Lokasi dan waktu penelitian]"),
                    H("3.3 Populasi dan Sampel / Subjek Penelitian", 2), Ai("Populasi, sampel, dan teknik sampling yang sesuai untuk {{judul}} dengan metode {{metode}}."),
                    H("3.4 Teknik Pengumpulan Data", 2), Ai("Teknik pengumpulan data yang sesuai metode {{metode}}."),
                    H("3.5 Instrumen Penelitian", 2), Ai("Instrumen penelitian beserta uji validitas dan reliabilitas."),
                    H("3.6 Teknik Analisis Data", 2), Ai("Teknik analisis data untuk metode {{metode}} (mis. uji prasyarat, uji hipotesis atau reduksi-display-verifikasi)."), Break(),
                    H("BAB IV HASIL PENELITIAN DAN PEMBAHASAN"), H("4.1 Deskripsi Data", 2), P("[Isi setelah data terkumpul — gunakan Statistik Deskriptif MASTER BW di Excel]"),
                    H("4.2 Hasil Uji Prasyarat/Analisis", 2), P("[Hasil uji normalitas, homogenitas, dsb.]"), H("4.3 Pembahasan", 2), P("[Bandingkan hasil dengan teori dan penelitian relevan]"), Break(),
                    H("BAB V KESIMPULAN DAN SARAN"), H("5.1 Kesimpulan", 2), P("[Jawaban atas rumusan masalah]"), H("5.2 Saran", 2), P("[Saran praktis dan untuk peneliti selanjutnya]"), Break(),
                    Center("DAFTAR PUSTAKA", true), P("[Gunakan menu Daftar Pustaka MASTER BW]"), Break(), Center("LAMPIRAN", true),
                },
            },
            new DocumentTemplate
            {
                Id = "proposal_penelitian", Name = "Proposal Penelitian", Category = "Akademik", Preset = "Proposal", Description = "Proposal penelitian BAB I–III + jadwal.",
                Fields = new List<TemplateField> { F("judul", "Judul penelitian"), F("penulis", "Peneliti"), F("nim", "NIM/NIP"), F("institusi", "Institusi"), F("kota", "Kota"), F("metode", "Metode penelitian") },
                Blocks = () => new List<DocBlock>
                {
                    new DocBlock { Type = "cover", Data = new Dictionary<string, string> { ["kind"] = "PROPOSAL PENELITIAN", ["title"] = "{{judul}}", ["author"] = "{{penulis}}", ["id_number"] = "{{nim}}", ["institution"] = "{{institusi}}", ["city"] = "{{kota}}", ["year"] = "{{tahun}}" } },
                    Toc(), Break(),
                    H("BAB I PENDAHULUAN"), H("1.1 Latar Belakang", 2), Ai("Latar belakang penelitian {{judul}}, 4-5 paragraf akademik."),
                    H("1.2 Rumusan Masalah", 2), Ai("Rumusan masalah penelitian {{judul}}."), H("1.3 Tujuan Penelitian", 2), Ai("Tujuan penelitian."),
                    H("1.4 Manfaat Penelitian", 2), Ai("Manfaat penelitian."), Break(),
                    H("BAB II TINJAUAN PUSTAKA"), H("2.1 Kajian Teori", 2), Ai("Kajian teori singkat untuk {{judul}} dengan penanda [sumber]."),
                    H("2.2 Penelitian Relevan", 2), P("[Penelitian terdahulu]"), H("2.3 Kerangka Pikir dan Hipotesis", 2), Ai("Kerangka pikir dan hipotesis {{judul}} ({{metode}})."), Break(),
                    H("BAB III METODE PENELITIAN"), Ai("Uraikan metode {{metode}}: desain, lokasi & waktu, populasi & sampel, teknik pengumpulan data, instrumen, teknik analisis untuk {{judul}}. Gunakan subjudul '## '."),
                    H("3.x Jadwal Penelitian", 2), Table("Kegiatan|Bulan 1|Bulan 2|Bulan 3|Bulan 4", "Penyusunan proposal|✓|||", "Pengumpulan data||✓|✓|", "Analisis data|||✓|", "Penyusunan laporan||||✓"),
                    Break(), Center("DAFTAR PUSTAKA", true),
                },
            },
            new DocumentTemplate
            {
                Id = "artikel_jurnal", Name = "Artikel Jurnal Ilmiah", Category = "Akademik", Preset = "Jurnal", Description = "Struktur IMRaD: abstrak, pendahuluan, metode, hasil & pembahasan, simpulan.",
                Fields = new List<TemplateField> { F("judul", "Judul artikel"), F("penulis", "Penulis"), F("afiliasi", "Afiliasi"), F("email", "Email korespondensi"), F("ringkasan", "Ringkasan penelitian") },
                Blocks = () => new List<DocBlock>
                {
                    Title("{{judul}}"), Center("{{penulis}}"), Center("{{afiliasi}}"), Center("{{email}}"), Space(),
                    H("Abstrak", 2), Ai("Abstrak 150-200 kata dari penelitian: {{ringkasan}}. Akhiri dengan 'Kata kunci:' 3-5 kata."),
                    H("Abstract", 2), Ai("Terjemahkan abstrak penelitian {{ringkasan}} ke bahasa Inggris akademik 150-200 kata dengan 'Keywords:'."),
                    H("PENDAHULUAN", 2), Ai("Pendahuluan artikel jurnal: konteks, kesenjangan riset, tujuan. Penelitian: {{ringkasan}}. Tandai rujukan [sumber]."),
                    H("METODE", 2), Ai("Bagian metode berdasarkan: {{ringkasan}}."),
                    H("HASIL DAN PEMBAHASAN", 2), P("[Sajikan hasil, tabel, grafik, dan pembahasan]"),
                    H("SIMPULAN", 2), P("[Simpulan]"), H("DAFTAR PUSTAKA", 2),
                },
            },
            new DocumentTemplate
            {
                Id = "laporan_pkl", Name = "Laporan PKL/Magang", Category = "Akademik", Preset = "Laporan", Description = "Laporan praktik kerja lapangan/magang.",
                Fields = new List<TemplateField> { F("penulis", "Nama"), F("nim", "NIM/NIS"), F("institusi", "Sekolah/kampus"), F("tempat_pkl", "Tempat PKL"), F("periode", "Periode PKL"), F("kota", "Kota"), F("kegiatan", "Kegiatan selama PKL") },
                Blocks = () => new List<DocBlock>
                {
                    new DocBlock { Type = "cover", Data = new Dictionary<string, string> { ["kind"] = "LAPORAN PRAKTIK KERJA LAPANGAN", ["title"] = "{{tempat_pkl}}", ["subtitle"] = "Periode {{periode}}", ["author"] = "{{penulis}}", ["id_number"] = "{{nim}}", ["institution"] = "{{institusi}}", ["city"] = "{{kota}}", ["year"] = "{{tahun}}" } },
                    Center("LEMBAR PENGESAHAN", true), Sign2("Pembimbing Lapangan,", "[nama]", "Pembimbing Sekolah/Kampus,", "[nama]"), Break(),
                    Center("KATA PENGANTAR", true), Ai("Kata pengantar laporan PKL di {{tempat_pkl}}."), Break(), Toc(), Break(),
                    H("BAB I PENDAHULUAN"), H("1.1 Latar Belakang", 2), Ai("Latar belakang PKL di {{tempat_pkl}}."), H("1.2 Tujuan PKL", 2), Ai("Tujuan PKL."),
                    H("BAB II PROFIL INSTANSI"), H("2.1 Sejarah Singkat", 2), P("[Sejarah {{tempat_pkl}}]"), H("2.2 Struktur Organisasi", 2), P("[Bagan struktur]"),
                    H("BAB III PELAKSANAAN PKL"), H("3.1 Kegiatan", 2), Ai("Uraian kegiatan PKL berdasarkan: {{kegiatan}}."), H("3.2 Kendala dan Solusi", 2), Ai("Kendala PKL dan solusinya."),
                    H("BAB IV PENUTUP"), H("4.1 Kesimpulan", 2), Ai("Kesimpulan PKL."), H("4.2 Saran", 2), Ai("Saran untuk instansi dan sekolah/kampus."),
                },
            },
            new DocumentTemplate
            {
                Id = "laporan_praktikum", Name = "Laporan Praktikum", Category = "Akademik", Preset = "Laporan", Description = "Laporan praktikum: tujuan, dasar teori, alat bahan, prosedur, data, pembahasan.",
                Fields = new List<TemplateField> { F("judul", "Judul praktikum"), F("penulis", "Nama"), F("nim", "NIM/NIS"), F("mata_kuliah", "Mata kuliah"), F("institusi", "Kampus/sekolah"), F("kota", "Kota") },
                Blocks = () => new List<DocBlock>
                {
                    new DocBlock { Type = "cover", Data = new Dictionary<string, string> { ["kind"] = "LAPORAN PRAKTIKUM {{mata_kuliah}}", ["title"] = "{{judul}}", ["author"] = "{{penulis}}", ["id_number"] = "{{nim}}", ["institution"] = "{{institusi}}", ["city"] = "{{kota}}", ["year"] = "{{tahun}}" } },
                    H("I. TUJUAN"), Ai("Tujuan praktikum {{judul}}."), H("II. DASAR TEORI"), Ai("Dasar teori praktikum {{judul}} 3 paragraf, dengan [sumber]."),
                    H("III. ALAT DAN BAHAN"), Ai("Daftar alat dan bahan yang umum untuk praktikum {{judul}} dalam poin."), H("IV. PROSEDUR KERJA"), Ai("Langkah kerja praktikum {{judul}} bernomor."),
                    H("V. DATA PENGAMATAN"), Table("No|Perlakuan|Hasil Pengamatan", "1|[..]|[..]"), H("VI. PEMBAHASAN"), P("[Analisis data dibandingkan teori]"),
                    H("VII. KESIMPULAN"), P("[Kesimpulan]"), H("DAFTAR PUSTAKA"),
                },
            },

            // ======================= PENDIDIKAN (GURU) =======================
            new DocumentTemplate
            {
                Id = "modul_ajar", Name = "Modul Ajar Kurikulum Merdeka", Category = "Pendidikan", Preset = "Laporan", Description = "Modul ajar: informasi umum, komponen inti, asesmen, lampiran LKPD.",
                Fields = new List<TemplateField> { F("penulis", "Nama guru"), F("instansi", "Satuan pendidikan"), F("mapel", "Mata pelajaran"), F("fase", "Fase/kelas"), F("materi", "Topik/materi"), F("alokasi", "Alokasi waktu", "2 × 45 menit"), F("model", "Model pembelajaran", "Problem Based Learning"), F("kota", "Kota") },
                Blocks = () => new List<DocBlock>
                {
                    Center("MODUL AJAR", true), Center("{{mapel}} – {{materi}}", true), Space(),
                    H("A. INFORMASI UMUM"), KV("Penyusun|: {{penulis}}", "Satuan pendidikan|: {{instansi}}", "Mata pelajaran|: {{mapel}}", "Fase/Kelas|: {{fase}}", "Topik|: {{materi}}", "Alokasi waktu|: {{alokasi}}", "Model pembelajaran|: {{model}}"),
                    H("Kompetensi Awal", 2), Ai("Kompetensi awal peserta didik untuk materi {{materi}} ({{mapel}}, {{fase}})."),
                    H("Profil Pelajar Pancasila / Dimensi Profil Lulusan", 2), Ai("Dimensi profil yang dikembangkan pada pembelajaran {{materi}}, dalam poin."),
                    H("Sarana dan Prasarana", 2), Ai("Sarana prasarana dan media pembelajaran {{materi}} dalam poin."),
                    H("B. KOMPONEN INTI"), H("Tujuan Pembelajaran", 2), Ai("Tujuan pembelajaran (ABCD) untuk {{materi}} {{fase}}, bernomor."),
                    H("Pemahaman Bermakna", 2), Ai("Pemahaman bermakna materi {{materi}}."),
                    H("Pertanyaan Pemantik", 2), Ai("3 pertanyaan pemantik untuk {{materi}}."),
                    H("Kegiatan Pembelajaran", 2), Ai("Langkah kegiatan pembelajaran {{materi}} dengan sintaks {{model}}: pendahuluan (10 menit), inti (sesuai sintaks), penutup (10 menit), total {{alokasi}}. Gunakan subjudul '## ' dan poin bernomor."),
                    H("Asesmen", 2), Ai("Asesmen diagnostik, formatif, dan sumatif untuk {{materi}} beserta teknik dan instrumennya."),
                    H("Pengayaan dan Remedial", 2), Ai("Kegiatan pengayaan dan remedial untuk {{materi}}."),
                    H("Refleksi Guru dan Peserta Didik", 2), Ai("Pertanyaan refleksi guru dan peserta didik."),
                    H("C. LAMPIRAN"), H("Lembar Kerja Peserta Didik (LKPD)", 2), Ai("LKPD sederhana untuk {{materi}}: identitas, petunjuk, 3-5 tugas/soal."),
                    H("Glosarium", 2), Ai("Glosarium 5 istilah penting materi {{materi}}."),
                    H("Daftar Pustaka", 2), P("[Buku siswa/buku guru Kemendikbudristek dan sumber lain]"),
                    Sign2("Kepala Sekolah,", "[nama]", "Guru Mata Pelajaran,", "{{penulis}}"),
                },
            },
            new DocumentTemplate
            {
                Id = "rpp", Name = "RPP Satu Lembar", Category = "Pendidikan", Description = "RPP ringkas (tujuan, langkah, asesmen).",
                Fields = new List<TemplateField> { F("penulis", "Nama guru"), F("instansi", "Sekolah"), F("mapel", "Mata pelajaran"), F("kelas", "Kelas/semester"), F("materi", "Materi"), F("alokasi", "Alokasi waktu", "2 × 45 menit"), F("kota", "Kota") },
                Blocks = () => new List<DocBlock>
                {
                    Center("RENCANA PELAKSANAAN PEMBELAJARAN (RPP)", true),
                    KV("Satuan pendidikan|: {{instansi}}", "Mata pelajaran|: {{mapel}}", "Kelas/Semester|: {{kelas}}", "Materi|: {{materi}}", "Alokasi waktu|: {{alokasi}}"),
                    H("A. Tujuan Pembelajaran", 2), Ai("Tujuan pembelajaran {{materi}} bernomor."),
                    H("B. Langkah-langkah Pembelajaran", 2), Ai("Kegiatan pendahuluan, inti, penutup untuk {{materi}} ({{alokasi}}), ringkas."),
                    H("C. Penilaian", 2), Ai("Penilaian sikap, pengetahuan, keterampilan untuk {{materi}} singkat."),
                    Sign2("Mengetahui, Kepala Sekolah", "[nama]", "Guru Mata Pelajaran,", "{{penulis}}"),
                },
            },
            new DocumentTemplate
            {
                Id = "soal_ujian", Name = "Kisi-kisi & Soal Ujian", Category = "Pendidikan", Description = "Soal pilihan ganda + esai dengan kunci jawaban.",
                Fields = new List<TemplateField> { F("instansi", "Sekolah"), F("mapel", "Mata pelajaran"), F("kelas", "Kelas"), F("materi", "Materi"), F("jumlah_pg", "Jumlah soal PG", "10"), F("jumlah_esai", "Jumlah soal esai", "3") },
                Blocks = () => new List<DocBlock>
                {
                    Center("SOAL PENILAIAN", true), Center("{{instansi}}", true),
                    KV("Mata pelajaran|: {{mapel}}", "Kelas|: {{kelas}}", "Materi|: {{materi}}", "Nama|: ....................", "No. absen|: ...."),
                    H("A. Pilihan Ganda", 2), P("Berilah tanda silang (X) pada jawaban yang paling tepat!"),
                    Ai("Buat {{jumlah_pg}} soal pilihan ganda HOTS materi {{materi}} kelas {{kelas}}, bernomor, masing-masing opsi a-e di baris terpisah. Jangan tulis kuncinya di sini."),
                    H("B. Esai", 2), Ai("Buat {{jumlah_esai}} soal esai materi {{materi}} kelas {{kelas}}, bernomor."), Break(),
                    H("Kunci Jawaban dan Pedoman Penskoran", 2), Ai("Tulis kunci jawaban untuk {{jumlah_pg}} soal pilihan ganda dan rubrik singkat {{jumlah_esai}} soal esai materi {{materi}}. Catatan: ini kunci untuk soal yang baru dibuat; periksa kecocokannya."),
                },
            },
        };

        public static DocumentTemplate Find(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return null;
            var key = id.Trim().ToLowerInvariant().Replace(" ", "_").Replace("-", "_");
            return All.FirstOrDefault(t => t.Id == key)
                ?? All.FirstOrDefault(t => t.Name.ToLowerInvariant() == id.Trim().ToLowerInvariant())
                ?? All.FirstOrDefault(t => key.Contains(t.Id) || t.Id.Contains(key))
                ?? All.FirstOrDefault(t => t.Name.ToLowerInvariant().Contains(id.Trim().ToLowerInvariant()));
        }
    }
}
