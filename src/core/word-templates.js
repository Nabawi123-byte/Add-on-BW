// Template surat & dokumen Word (berjalan 100% offline).
// build(v) menghasilkan daftar "blok" yang dirender ke HTML lalu disisipkan ke Word.

import { tanggalIndonesia, formatRupiah, parseNumber } from './util.js';
import { terbilangRupiah } from './terbilang.js';

const f = (key, label, def = '', type = 'text') => ({ key, label, def, type });

// Field profil diisi otomatis dari Pengaturan > Profil.
const P = {
  nama: (lbl = 'Nama Anda') => f('nama', lbl, (pr) => pr.nama || ''),
  jabatan: (lbl = 'Jabatan') => f('jabatan', lbl, (pr) => pr.jabatan || ''),
  instansi: (lbl = 'Instansi / Perusahaan') => f('instansi', lbl, (pr) => pr.instansi || ''),
  alamat: (lbl = 'Alamat Instansi') => f('alamat_instansi', lbl, (pr) => pr.alamat || ''),
  kota: () => f('kota', 'Kota', (pr) => pr.kota || 'Jakarta'),
  tanggal: () => f('tanggal', 'Tanggal Surat', () => tanggalIndonesia()),
};

function kop(v) {
  if (!v.instansi) return [];
  const out = [{ type: 'heading', level: 1, text: String(v.instansi).toUpperCase(), align: 'center' }];
  if (v.alamat_instansi) out.push({ type: 'paragraph', text: v.alamat_instansi, align: 'center', size: 10 });
  out.push({ type: 'line' });
  return out;
}

const tempatTanggal = (v) => ({ type: 'paragraph', text: `${v.kota || ''}, ${v.tanggal || tanggalIndonesia()}`, align: 'right' });

const kepada = (v, k = 'kepada', t = 'tempat_tujuan') => ({
  type: 'paragraph',
  text: `Kepada Yth.\n${v[k] || '..........'}\ndi ${v[t] || 'Tempat'}`,
});

const ttd = (v, { title = 'Hormat saya,', name = v.nama, position = v.jabatan, align = 'right', place = false } = {}) => ({
  type: 'signature',
  title,
  name: name || '(..............................)',
  position: position || '',
  place: place ? `${v.kota || ''}, ${v.tanggal || tanggalIndonesia()}` : '',
  align,
});

const p = (text, extra = {}) => Object.assign({ type: 'paragraph', text, align: 'justify' }, extra);
const judul = (text, sub) => {
  const out = [{ type: 'heading', level: 2, text, align: 'center', underline: true }];
  if (sub) out.push({ type: 'paragraph', text: sub, align: 'center' });
  return out;
};

export const WORD_TEMPLATES = [
  {
    id: 'izin',
    title: 'Surat Izin Tidak Masuk',
    icon: '🤒',
    keywords: ['surat izin', 'izin tidak masuk', 'izin sakit', 'izin kerja', 'izin sekolah', 'surat sakit'],
    fields: [P.kota(), P.tanggal(), f('kepada', 'Kepada', 'Bapak/Ibu Pimpinan'), P.instansi('Instansi / Sekolah'), P.nama(), f('jabatan', 'Jabatan / Kelas', (pr) => pr.jabatan || ''), f('alasan', 'Alasan', 'sakit'), f('mulai', 'Tanggal Mulai', () => tanggalIndonesia()), f('lama', 'Lama (hari)', '1')],
    build: (v) => [
      tempatTanggal(v),
      kepada({ ...v, tempat_tujuan: v.instansi ? v.instansi : 'Tempat' }),
      p('Dengan hormat,'),
      p('Saya yang bertanda tangan di bawah ini:'),
      { type: 'fields', rows: [['Nama', v.nama], ['Jabatan / Kelas', v.jabatan], ['Instansi', v.instansi]] },
      p(`Dengan ini memberitahukan bahwa saya tidak dapat hadir/masuk selama ${v.lama || 1} hari terhitung mulai tanggal ${v.mulai} dikarenakan ${v.alasan}.`),
      p('Demikian surat izin ini saya sampaikan. Atas perhatian dan pengertian Bapak/Ibu, saya ucapkan terima kasih.'),
      ttd(v),
    ],
  },
  {
    id: 'undangan',
    title: 'Surat Undangan',
    icon: '✉️',
    keywords: ['undangan', 'surat undangan', 'undangan rapat', 'mengundang'],
    fields: [P.instansi(), P.alamat(), f('nomor', 'Nomor Surat', '001/UND/2026'), f('perihal', 'Perihal', 'Undangan Rapat'), P.kota(), P.tanggal(), f('kepada', 'Kepada', 'Bapak/Ibu Undangan'), f('hari_tanggal', 'Hari, Tanggal', ''), f('waktu', 'Waktu', '09.00 WIB s.d. selesai'), f('tempat', 'Tempat', ''), f('acara', 'Acara', 'Rapat koordinasi'), P.nama('Nama Pengundang'), P.jabatan('Jabatan Pengundang')],
    build: (v) => [
      ...kop(v),
      { type: 'fields', indent: false, rows: [['Nomor', v.nomor], ['Lampiran', '-'], ['Perihal', v.perihal]] },
      tempatTanggal(v),
      kepada(v),
      p('Dengan hormat,'),
      p('Sehubungan dengan kegiatan yang akan kami selenggarakan, kami mengharapkan kehadiran Bapak/Ibu pada:'),
      { type: 'fields', rows: [['Hari, Tanggal', v.hari_tanggal], ['Waktu', v.waktu], ['Tempat', v.tempat], ['Acara', v.acara]] },
      p('Mengingat pentingnya acara tersebut, kami sangat mengharapkan kehadiran Bapak/Ibu tepat pada waktunya. Atas perhatian dan kehadirannya kami ucapkan terima kasih.'),
      ttd(v, { title: 'Hormat kami,' }),
    ],
  },
  {
    id: 'lamaran',
    title: 'Surat Lamaran Kerja',
    icon: '💼',
    keywords: ['lamaran', 'lamaran kerja', 'melamar', 'surat lamaran', 'apply kerja'],
    fields: [P.kota(), P.tanggal(), f('kepada', 'Kepada', 'HRD Manager'), f('perusahaan', 'Nama Perusahaan', ''), f('posisi', 'Posisi yang Dilamar', ''), f('sumber', 'Sumber Info Lowongan', 'situs lowongan kerja'), P.nama('Nama Lengkap'), f('ttl', 'Tempat, Tanggal Lahir', ''), f('pendidikan', 'Pendidikan Terakhir', ''), f('alamat', 'Alamat', ''), f('telepon', 'No. Telepon', (pr) => pr.telepon || ''), f('email', 'Email', (pr) => pr.email || ''), f('lampiran', 'Lampiran (pisahkan koma)', 'Daftar Riwayat Hidup, Fotokopi Ijazah dan Transkrip Nilai, Fotokopi KTP, Pas Foto Terbaru')],
    build: (v) => [
      tempatTanggal(v),
      { type: 'fields', indent: false, rows: [['Perihal', `Lamaran Pekerjaan sebagai ${v.posisi}`]] },
      kepada({ ...v, tempat_tujuan: v.perusahaan || 'Tempat' }),
      p('Dengan hormat,'),
      p(`Berdasarkan informasi lowongan pekerjaan yang saya peroleh dari ${v.sumber}, saya bermaksud mengajukan lamaran untuk posisi ${v.posisi} di ${v.perusahaan}. Adapun data diri saya sebagai berikut:`),
      { type: 'fields', rows: [['Nama', v.nama], ['Tempat, Tanggal Lahir', v.ttl], ['Pendidikan Terakhir', v.pendidikan], ['Alamat', v.alamat], ['No. Telepon', v.telepon], ['Email', v.email]] },
      p('Saya memiliki motivasi tinggi, mampu bekerja secara mandiri maupun dalam tim, serta siap belajar hal-hal baru. Sebagai bahan pertimbangan, bersama surat ini saya lampirkan:'),
      { type: 'list', ordered: true, items: String(v.lampiran || '').split(/\s*,\s*/).filter(Boolean) },
      p('Besar harapan saya untuk dapat diberikan kesempatan wawancara agar dapat menjelaskan lebih lanjut mengenai potensi diri saya. Atas perhatian Bapak/Ibu, saya ucapkan terima kasih.'),
      ttd(v, { position: '' }),
    ],
  },
  {
    id: 'kuasa',
    title: 'Surat Kuasa',
    icon: '✍️',
    keywords: ['surat kuasa', 'kuasa', 'memberi kuasa', 'pemberian kuasa'],
    fields: [f('pemberi', 'Nama Pemberi Kuasa', (pr) => pr.nama || ''), f('nik_pemberi', 'NIK Pemberi', ''), f('alamat_pemberi', 'Alamat Pemberi', ''), f('penerima', 'Nama Penerima Kuasa', ''), f('nik_penerima', 'NIK Penerima', ''), f('alamat_penerima', 'Alamat Penerima', ''), f('keperluan', 'Keperluan', 'mengambil dokumen ...', 'textarea'), P.kota(), P.tanggal()],
    build: (v) => [
      ...judul('SURAT KUASA'),
      p('Yang bertanda tangan di bawah ini:'),
      { type: 'fields', rows: [['Nama', v.pemberi], ['NIK', v.nik_pemberi], ['Alamat', v.alamat_pemberi]] },
      p('Selanjutnya disebut sebagai PEMBERI KUASA.'),
      p('Dengan ini memberikan kuasa kepada:'),
      { type: 'fields', rows: [['Nama', v.penerima], ['NIK', v.nik_penerima], ['Alamat', v.alamat_penerima]] },
      p('Selanjutnya disebut sebagai PENERIMA KUASA.'),
      p(`------------------------------------- KHUSUS -------------------------------------`, { align: 'center', bold: true }),
      p(`Untuk ${v.keperluan}.`),
      p('Demikian surat kuasa ini dibuat dengan sebenar-benarnya untuk dapat dipergunakan sebagaimana mestinya.'),
      { type: 'paragraph', text: `${v.kota}, ${v.tanggal}`, align: 'right' },
      { type: 'signatures', items: [{ title: 'Penerima Kuasa,', name: v.penerima }, { title: 'Pemberi Kuasa,', name: v.pemberi, note: 'Meterai 10.000' }] },
    ],
  },
  {
    id: 'pernyataan',
    title: 'Surat Pernyataan',
    icon: '📜',
    keywords: ['surat pernyataan', 'pernyataan', 'menyatakan'],
    fields: [P.nama(), f('nik', 'NIK / NIP', ''), f('ttl', 'Tempat, Tanggal Lahir', ''), f('alamat', 'Alamat', ''), f('isi', 'Isi Pernyataan', 'tidak sedang terikat kontrak kerja dengan perusahaan lain', 'textarea'), P.kota(), P.tanggal()],
    build: (v) => [
      ...judul('SURAT PERNYATAAN'),
      p('Yang bertanda tangan di bawah ini:'),
      { type: 'fields', rows: [['Nama', v.nama], ['NIK / NIP', v.nik], ['Tempat, Tanggal Lahir', v.ttl], ['Alamat', v.alamat]] },
      p(`Dengan ini menyatakan dengan sesungguhnya bahwa saya ${v.isi}.`),
      p('Demikian surat pernyataan ini saya buat dengan sebenar-benarnya, dalam keadaan sadar dan tanpa paksaan dari pihak mana pun. Apabila di kemudian hari pernyataan ini terbukti tidak benar, saya bersedia menerima sanksi sesuai ketentuan yang berlaku.'),
      ttd(v, { title: 'Yang menyatakan,', position: '', place: true, name: v.nama }),
      { type: 'paragraph', text: '(Meterai 10.000)', align: 'right', size: 9, italic: true },
    ],
  },
  {
    id: 'keterangan',
    title: 'Surat Keterangan',
    icon: '🏷️',
    keywords: ['surat keterangan', 'keterangan kerja', 'keterangan domisili', 'keterangan aktif', 'menerangkan'],
    fields: [P.instansi(), P.alamat(), f('nomor', 'Nomor Surat', '001/SK/2026'), f('jenis', 'Jenis Keterangan', 'Keterangan Kerja'), P.nama('Nama Penandatangan'), P.jabatan('Jabatan Penandatangan'), f('yang_diterangkan', 'Nama yang Diterangkan', ''), f('identitas', 'NIK / NIP / NIS', ''), f('jabatan_ybs', 'Jabatan / Status', ''), f('isi', 'Isi Keterangan', 'benar merupakan karyawan aktif di instansi kami', 'textarea'), P.kota(), P.tanggal()],
    build: (v) => [
      ...kop(v),
      ...judul(`SURAT ${String(v.jenis).toUpperCase()}`, `Nomor: ${v.nomor}`),
      p('Yang bertanda tangan di bawah ini:'),
      { type: 'fields', rows: [['Nama', v.nama], ['Jabatan', v.jabatan], ['Instansi', v.instansi]] },
      p('Dengan ini menerangkan bahwa:'),
      { type: 'fields', rows: [['Nama', v.yang_diterangkan], ['NIK / NIP / NIS', v.identitas], ['Jabatan / Status', v.jabatan_ybs]] },
      p(`Adalah ${v.isi}.`),
      p('Demikian surat keterangan ini dibuat untuk dapat dipergunakan sebagaimana mestinya.'),
      ttd(v, { title: 'Hormat kami,', place: true }),
    ],
  },
  {
    id: 'resign',
    title: 'Surat Pengunduran Diri',
    icon: '🚪',
    keywords: ['pengunduran diri', 'resign', 'mengundurkan diri', 'surat resign'],
    fields: [P.kota(), P.tanggal(), f('kepada', 'Kepada', 'Pimpinan HRD'), P.instansi('Perusahaan'), P.nama(), P.jabatan(), f('efektif', 'Tanggal Efektif', ''), f('alasan', 'Alasan (opsional)', 'ingin mengembangkan karier di bidang lain')],
    build: (v) => [
      tempatTanggal(v),
      kepada({ ...v, tempat_tujuan: v.instansi || 'Tempat' }),
      { type: 'fields', indent: false, rows: [['Perihal', 'Pengunduran Diri']] },
      p('Dengan hormat,'),
      p('Saya yang bertanda tangan di bawah ini:'),
      { type: 'fields', rows: [['Nama', v.nama], ['Jabatan', v.jabatan]] },
      p(`Dengan ini mengajukan pengunduran diri dari ${v.instansi} terhitung mulai tanggal ${v.efektif}${v.alasan ? ', dikarenakan ' + v.alasan : ''}.`),
      p('Saya mengucapkan terima kasih atas kesempatan, kepercayaan, dan pengalaman berharga yang telah diberikan selama saya bekerja. Saya juga memohon maaf apabila selama bekerja terdapat kesalahan yang saya perbuat. Saya akan menyelesaikan seluruh tanggung jawab dan proses serah terima pekerjaan dengan baik.'),
      p('Demikian surat pengunduran diri ini saya sampaikan. Atas perhatian Bapak/Ibu, saya ucapkan terima kasih.'),
      ttd(v),
    ],
  },
  {
    id: 'permohonan',
    title: 'Surat Permohonan',
    icon: '🙏',
    keywords: ['surat permohonan', 'permohonan', 'mengajukan permohonan', 'memohon'],
    fields: [P.instansi(), P.alamat(), f('nomor', 'Nomor Surat', '001/PRM/2026'), f('perihal', 'Perihal', 'Permohonan ...'), P.kota(), P.tanggal(), f('kepada', 'Kepada', ''), f('isi', 'Isi Permohonan', '', 'textarea'), P.nama(), P.jabatan()],
    build: (v) => [
      ...kop(v),
      { type: 'fields', indent: false, rows: [['Nomor', v.nomor], ['Lampiran', '-'], ['Perihal', v.perihal]] },
      tempatTanggal(v),
      kepada(v),
      p('Dengan hormat,'),
      p(v.isi || '..........'),
      p('Demikian surat permohonan ini kami sampaikan. Besar harapan kami agar permohonan ini dapat dikabulkan. Atas perhatian dan kerja samanya, kami ucapkan terima kasih.'),
      ttd(v, { title: 'Hormat kami,' }),
    ],
  },
  {
    id: 'cuti',
    title: 'Permohonan Cuti',
    icon: '🏖️',
    keywords: ['cuti', 'permohonan cuti', 'surat cuti', 'izin cuti'],
    fields: [P.kota(), P.tanggal(), f('kepada', 'Kepada', 'Kepala Bagian SDM'), P.instansi(), P.nama(), P.jabatan(), f('jenis_cuti', 'Jenis Cuti', 'cuti tahunan'), f('mulai', 'Mulai', ''), f('selesai', 'Sampai', ''), f('alasan', 'Keperluan', 'keperluan keluarga'), f('pengganti', 'Pelaksana Tugas Sementara', '')],
    build: (v) => [
      tempatTanggal(v),
      kepada({ ...v, tempat_tujuan: v.instansi || 'Tempat' }),
      { type: 'fields', indent: false, rows: [['Perihal', `Permohonan ${v.jenis_cuti}`]] },
      p('Dengan hormat,'),
      p('Saya yang bertanda tangan di bawah ini:'),
      { type: 'fields', rows: [['Nama', v.nama], ['Jabatan', v.jabatan]] },
      p(`Bermaksud mengajukan ${v.jenis_cuti} mulai tanggal ${v.mulai} sampai dengan ${v.selesai} untuk ${v.alasan}.${v.pengganti ? ` Selama cuti, tugas saya akan dilaksanakan sementara oleh ${v.pengganti}.` : ''}`),
      p('Demikian permohonan ini saya ajukan. Atas perhatian dan persetujuan Bapak/Ibu, saya ucapkan terima kasih.'),
      ttd(v),
    ],
  },
  {
    id: 'tugas',
    title: 'Surat Tugas',
    icon: '🧭',
    keywords: ['surat tugas', 'penugasan', 'menugaskan', 'surat perintah tugas'],
    fields: [P.instansi(), P.alamat(), f('nomor', 'Nomor Surat', '001/ST/2026'), P.nama('Nama Pemberi Tugas'), P.jabatan('Jabatan Pemberi Tugas'), f('petugas', 'Nama yang Ditugaskan (pisahkan koma)', ''), f('jabatan_petugas', 'Jabatan yang Ditugaskan', ''), f('tugas', 'Tugas', '', 'textarea'), f('tempat', 'Tempat', ''), f('waktu', 'Waktu Pelaksanaan', ''), P.kota(), P.tanggal()],
    build: (v) => [
      ...kop(v),
      ...judul('SURAT TUGAS', `Nomor: ${v.nomor}`),
      p('Yang bertanda tangan di bawah ini:'),
      { type: 'fields', rows: [['Nama', v.nama], ['Jabatan', v.jabatan]] },
      p('Memberikan tugas kepada:'),
      {
        type: 'table',
        header: true,
        rows: [['No', 'Nama', 'Jabatan'], ...String(v.petugas || '').split(/\s*,\s*/).filter(Boolean).map((n, i) => [String(i + 1), n, v.jabatan_petugas || ''])],
      },
      p(`Untuk ${v.tugas} yang dilaksanakan di ${v.tempat} pada ${v.waktu}.`),
      p('Demikian surat tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab.'),
      ttd(v, { title: 'Yang memberi tugas,', place: true }),
    ],
  },
  {
    id: 'edaran',
    title: 'Surat Edaran',
    icon: '📢',
    keywords: ['surat edaran', 'edaran', 'pengumuman'],
    fields: [P.instansi(), P.alamat(), f('nomor', 'Nomor', '001/SE/2026'), f('perihal', 'Tentang', ''), f('kepada', 'Ditujukan kepada', 'Seluruh Karyawan'), f('isi', 'Isi Edaran (satu poin per baris)', '', 'textarea'), P.nama(), P.jabatan(), P.kota(), P.tanggal()],
    build: (v) => [
      ...kop(v),
      ...judul('SURAT EDARAN', `Nomor: ${v.nomor}`),
      { type: 'paragraph', text: `TENTANG\n${String(v.perihal).toUpperCase()}`, align: 'center', bold: true },
      p(`Kepada Yth. ${v.kepada}`),
      p('Sehubungan dengan hal tersebut di atas, dengan ini disampaikan hal-hal sebagai berikut:'),
      { type: 'list', ordered: true, items: String(v.isi || '').split(/\n+/).filter(Boolean) },
      p('Demikian surat edaran ini disampaikan untuk diperhatikan dan dilaksanakan sebagaimana mestinya.'),
      ttd(v, { title: 'Ditetapkan oleh,', place: true }),
    ],
  },
  {
    id: 'memo',
    title: 'Memo Internal',
    icon: '🗒️',
    keywords: ['memo', 'memorandum', 'nota dinas'],
    fields: [P.instansi(), f('kepada', 'Kepada', ''), f('dari', 'Dari', (pr) => [pr.nama, pr.jabatan].filter(Boolean).join(' - ')), f('tembusan', 'Tembusan', '-'), P.tanggal(), f('perihal', 'Perihal', ''), f('isi', 'Isi Memo', '', 'textarea')],
    build: (v) => [
      ...kop(v),
      ...judul('MEMO INTERNAL'),
      { type: 'fields', indent: false, rows: [['Kepada', v.kepada], ['Dari', v.dari], ['Tembusan', v.tembusan], ['Tanggal', v.tanggal], ['Perihal', v.perihal]] },
      { type: 'line' },
      p(v.isi || '..........'),
      p('Terima kasih atas perhatian dan kerja samanya.'),
    ],
  },
  {
    id: 'notulen',
    title: 'Notulen Rapat',
    icon: '📝',
    keywords: ['notulen', 'notula', 'risalah rapat', 'catatan rapat', 'minutes'],
    fields: [P.instansi(), f('agenda', 'Agenda Rapat', ''), f('hari_tanggal', 'Hari, Tanggal', () => tanggalIndonesia()), f('waktu', 'Waktu', ''), f('tempat', 'Tempat', ''), f('pimpinan', 'Pimpinan Rapat', ''), f('notulis', 'Notulis', (pr) => pr.nama || ''), f('peserta', 'Peserta (pisahkan koma)', ''), f('pembahasan', 'Pembahasan (satu poin per baris)', '', 'textarea'), f('keputusan', 'Keputusan (satu poin per baris)', '', 'textarea'), f('tindak_lanjut', 'Tindak Lanjut (format: tugas | PIC | tenggat)', '', 'textarea')],
    build: (v) => {
      const tl = String(v.tindak_lanjut || '').split(/\n+/).filter(Boolean).map((l, i) => {
        const [a, b, c] = l.split('|').map((s) => s.trim());
        return [String(i + 1), a || '', b || '', c || ''];
      });
      return [
        ...kop(v),
        ...judul('NOTULEN RAPAT'),
        { type: 'fields', rows: [['Agenda', v.agenda], ['Hari, Tanggal', v.hari_tanggal], ['Waktu', v.waktu], ['Tempat', v.tempat], ['Pimpinan Rapat', v.pimpinan], ['Notulis', v.notulis]] },
        { type: 'heading', level: 3, text: 'Peserta Rapat' },
        { type: 'list', ordered: true, items: String(v.peserta || '').split(/\s*,\s*/).filter(Boolean) },
        { type: 'heading', level: 3, text: 'Pembahasan' },
        { type: 'list', ordered: true, items: String(v.pembahasan || '').split(/\n+/).filter(Boolean) },
        { type: 'heading', level: 3, text: 'Keputusan' },
        { type: 'list', ordered: true, items: String(v.keputusan || '').split(/\n+/).filter(Boolean) },
        { type: 'heading', level: 3, text: 'Tindak Lanjut' },
        { type: 'table', header: true, rows: [['No', 'Tindak Lanjut', 'Penanggung Jawab', 'Tenggat'], ...(tl.length ? tl : [['1', '', '', '']])] },
        { type: 'signatures', items: [{ title: 'Pimpinan Rapat,', name: v.pimpinan }, { title: 'Notulis,', name: v.notulis }] },
      ];
    },
  },
  {
    id: 'berita-acara',
    title: 'Berita Acara',
    icon: '📋',
    keywords: ['berita acara', 'serah terima', 'bast'],
    fields: [P.instansi(), f('judul', 'Judul', 'Berita Acara Serah Terima Barang'), f('nomor', 'Nomor', '001/BA/2026'), f('hari_tanggal', 'Hari, Tanggal', () => tanggalIndonesia()), f('pihak1', 'Pihak Pertama (nama - jabatan)', ''), f('pihak2', 'Pihak Kedua (nama - jabatan)', ''), f('isi', 'Isi / Uraian (satu poin per baris)', '', 'textarea'), P.kota()],
    build: (v) => [
      ...kop(v),
      ...judul(String(v.judul).toUpperCase(), `Nomor: ${v.nomor}`),
      p(`Pada hari ini, ${v.hari_tanggal}, kami yang bertanda tangan di bawah ini:`),
      { type: 'fields', rows: [['Pihak Pertama', v.pihak1], ['Pihak Kedua', v.pihak2]] },
      p('Menyatakan hal-hal sebagai berikut:'),
      { type: 'list', ordered: true, items: String(v.isi || '').split(/\n+/).filter(Boolean) },
      p('Demikian berita acara ini dibuat dengan sebenar-benarnya dan ditandatangani oleh kedua belah pihak untuk dipergunakan sebagaimana mestinya.'),
      { type: 'signatures', items: [{ title: 'Pihak Pertama,', name: String(v.pihak1 || '').split(' - ')[0] }, { title: 'Pihak Kedua,', name: String(v.pihak2 || '').split(' - ')[0] }] },
    ],
  },
  {
    id: 'kwitansi',
    title: 'Kwitansi',
    icon: '🧾',
    keywords: ['kwitansi', 'kuitansi', 'tanda terima', 'bukti pembayaran'],
    fields: [f('nomor', 'No. Kwitansi', '001/KW/2026'), f('dari', 'Sudah Terima Dari', ''), f('jumlah', 'Jumlah (angka)', ''), f('untuk', 'Untuk Pembayaran', ''), P.kota(), P.tanggal(), f('penerima', 'Penerima', (pr) => pr.nama || '')],
    build: (v) => {
      const n = parseNumber(v.jumlah) || 0;
      return [
        ...judul('KWITANSI', `No: ${v.nomor}`),
        {
          type: 'fields',
          rows: [
            ['Sudah terima dari', v.dari],
            ['Uang sejumlah', terbilangRupiah(n)],
            ['Untuk pembayaran', v.untuk],
          ],
        },
        { type: 'paragraph', text: `Terbilang: ${formatRupiah(n)},-`, bold: true, size: 14, border: true },
        ttd(v, { title: 'Penerima,', name: v.penerima, position: '', place: true }),
      ];
    },
  },
  {
    id: 'penawaran',
    title: 'Surat Penawaran',
    icon: '🏷️',
    keywords: ['penawaran', 'surat penawaran', 'penawaran harga', 'quotation'],
    fields: [P.instansi(), P.alamat(), f('nomor', 'Nomor Surat', '001/PNW/2026'), P.kota(), P.tanggal(), f('kepada', 'Kepada', ''), f('item', 'Item (format: nama | qty | harga) satu per baris', '', 'textarea'), f('ketentuan', 'Ketentuan (satu per baris)', 'Harga belum termasuk PPN\nPenawaran berlaku 14 hari\nPembayaran 50% di muka', 'textarea'), P.nama(), P.jabatan()],
    build: (v) => {
      let total = 0;
      const rows = String(v.item || '').split(/\n+/).filter(Boolean).map((l, i) => {
        const [nama, qty, harga] = l.split('|').map((s) => s.trim());
        const q = parseNumber(qty) || 0;
        const h = parseNumber(harga) || 0;
        total += q * h;
        return [String(i + 1), nama || '', String(q), formatRupiah(h), formatRupiah(q * h)];
      });
      return [
        ...kop(v),
        { type: 'fields', indent: false, rows: [['Nomor', v.nomor], ['Perihal', 'Penawaran Harga']] },
        tempatTanggal(v),
        kepada(v),
        p('Dengan hormat,'),
        p('Bersama surat ini, kami bermaksud menawarkan produk/jasa kami dengan rincian sebagai berikut:'),
        { type: 'table', header: true, rows: [['No', 'Uraian', 'Qty', 'Harga Satuan', 'Jumlah'], ...(rows.length ? rows : [['1', '', '', '', '']]), ['', 'TOTAL', '', '', formatRupiah(total)]] },
        p('Ketentuan:'),
        { type: 'list', items: String(v.ketentuan || '').split(/\n+/).filter(Boolean) },
        p('Demikian surat penawaran ini kami sampaikan. Kami berharap dapat menjalin kerja sama yang baik dengan Bapak/Ibu. Atas perhatiannya kami ucapkan terima kasih.'),
        ttd(v, { title: 'Hormat kami,' }),
      ];
    },
  },
  {
    id: 'peringatan',
    title: 'Surat Peringatan (SP)',
    icon: '⚠️',
    keywords: ['surat peringatan', 'sp1', 'sp 1', 'sp2', 'sp 2', 'sp3', 'sp 3', 'teguran', 'peringatan'],
    fields: [P.instansi(), P.alamat(), f('tingkat', 'Tingkat (I / II / III)', 'I'), f('nomor', 'Nomor', '001/SP/2026'), f('karyawan', 'Nama Karyawan', ''), f('jabatan_karyawan', 'Jabatan Karyawan', ''), f('pelanggaran', 'Pelanggaran', '', 'textarea'), f('berlaku', 'Masa Berlaku', '6 (enam) bulan'), P.nama(), P.jabatan(), P.kota(), P.tanggal()],
    build: (v) => [
      ...kop(v),
      ...judul(`SURAT PERINGATAN ${v.tingkat}`, `Nomor: ${v.nomor}`),
      p('Surat peringatan ini ditujukan kepada:'),
      { type: 'fields', rows: [['Nama', v.karyawan], ['Jabatan', v.jabatan_karyawan]] },
      p(`Berdasarkan hasil evaluasi, yang bersangkutan telah melakukan pelanggaran berupa ${v.pelanggaran}. Oleh karena itu, perusahaan memberikan Surat Peringatan ${v.tingkat} yang berlaku selama ${v.berlaku} sejak tanggal surat ini diterbitkan.`),
      p('Kami berharap yang bersangkutan dapat memperbaiki sikap dan kinerjanya. Apabila pelanggaran serupa terulang selama masa berlaku surat ini, perusahaan akan mengambil tindakan lebih lanjut sesuai peraturan yang berlaku.'),
      { type: 'signatures', items: [{ title: 'Diterima oleh,', name: v.karyawan }, { title: 'Hormat kami,', name: v.nama, position: v.jabatan }], place: `${v.kota}, ${v.tanggal}` },
    ],
  },
  {
    id: 'rekomendasi',
    title: 'Surat Rekomendasi',
    icon: '⭐',
    keywords: ['rekomendasi', 'surat rekomendasi', 'referensi kerja'],
    fields: [P.instansi(), P.alamat(), P.nama('Nama Pemberi Rekomendasi'), P.jabatan(), f('kandidat', 'Nama yang Direkomendasikan', ''), f('hubungan', 'Hubungan / Lama Mengenal', 'bekerja di bawah supervisi saya selama 2 tahun'), f('kelebihan', 'Kelebihan (satu per baris)', 'Disiplin dan bertanggung jawab\nMampu bekerja sama dalam tim\nCepat belajar', 'textarea'), f('tujuan', 'Tujuan Rekomendasi', 'melanjutkan karier / studi'), P.kota(), P.tanggal()],
    build: (v) => [
      ...kop(v),
      ...judul('SURAT REKOMENDASI'),
      p('Yang bertanda tangan di bawah ini:'),
      { type: 'fields', rows: [['Nama', v.nama], ['Jabatan', v.jabatan], ['Instansi', v.instansi]] },
      p(`Dengan ini memberikan rekomendasi kepada ${v.kandidat}, yang ${v.hubungan}. Selama mengenal yang bersangkutan, saya menilai bahwa ia memiliki kualitas sebagai berikut:`),
      { type: 'list', items: String(v.kelebihan || '').split(/\n+/).filter(Boolean) },
      p(`Oleh karena itu, saya merekomendasikan yang bersangkutan untuk ${v.tujuan}. Saya bersedia dihubungi apabila diperlukan informasi lebih lanjut.`),
      ttd(v, { title: 'Hormat saya,', place: true }),
    ],
  },
];

export function findWordTemplate(text) {
  let best = null;
  for (const t of WORD_TEMPLATES) {
    for (const k of t.keywords) {
      if (text.includes(k) && (!best || k.length > best.len)) best = { t, len: k.length };
    }
  }
  return best ? best.t : null;
}

/** Nilai default untuk form template (profil + nilai dari perintah). */
export function templateDefaults(tpl, profile = {}, given = {}) {
  const out = {};
  const keys = Object.keys(given || {});
  const fieldKeys = new Set(tpl.fields.map((x) => x.key));
  const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
  const find = (fd) => {
    const label = norm(fd.label);
    const exact = keys.find((g) => g === fd.key || g === label);
    if (exact !== undefined) return given[exact];
    // Kecocokan longgar hanya untuk kunci yang bukan milik field lain (mis. "hari_tanggal" bukan "tanggal").
    const k = keys.find((g) => !fieldKeys.has(g) && (g.endsWith('_' + fd.key) || (label.startsWith(g) && g.length >= 4)));
    return k === undefined ? undefined : given[k];
  };
  for (const fd of tpl.fields) {
    const g = find(fd);
    out[fd.key] = g !== undefined && g !== '' ? g : typeof fd.def === 'function' ? fd.def(profile) : fd.def;
  }
  return out;
}
