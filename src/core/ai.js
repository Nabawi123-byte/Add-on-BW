// Agen AI Master: Claude + alat (tools) yang langsung mengubah Word/Excel.
// Berjalan di panel add-in (browser), memakai API key milik pengguna.

import Anthropic from '@anthropic-ai/sdk';
import { validate } from './schema.js';

export const MODELS = [
  { id: 'claude-opus-5-5', label: 'Claude Opus 5.5 (paling cerdas)' },
  { id: 'claude-sonnet-5-5', label: 'Claude Sonnet 5.5 (cepat & hemat)' },
  { id: 'claude-haiku-5-5', label: 'Claude Haiku 5.5 (paling hemat)' },
];
export const DEFAULT_MODEL = 'claude-opus-5-5';

// Fallback otomatis di server bila permintaan ditolak filter keamanan (tidak tersedia untuk Haiku).
const FALLBACK_BETA = 'server-side-fallback-2026-07-01';
const supportsFallback = (model) => model === 'claude-opus-5-5' || model === 'claude-sonnet-5-5';

/* ------------------------------ definisi alat ------------------------------ */

const OPS_EXCEL_DOC = `Daftar operasi (field "op"):
- write {address, values: 2D array}: isi sel mulai dari address (sel kiri-atas atau rentang pas). String diawali "=" adalah rumus. Rumus WAJIB nama fungsi bahasa Inggris dan pemisah koma (SUM, IF, VLOOKUP, XLOOKUP, SUMIFS, ...), apa pun bahasa Excel pengguna. Opsional numberFormat.
- format {address, bold, italic, underline, fontSize, fontName, fontColor, fill, hAlign (Left|Center|Right|Justify), vAlign (Top|Center|Bottom), wrap, numberFormat, border (all|outline|top|top-double|bottom|none), merge, columnWidth (dalam karakter), rowHeight (poin)}
- numberFormat {address, format}: kode format en-US, mis. '"Rp "#,##0' untuk rupiah, '0.0%', 'dd/mm/yyyy', '#,##0'.
- table {address, name, style, headers}: jadikan Excel Table (style mis. TableStyleMedium2).
- chart {address, type (ColumnClustered|BarClustered|Line|Pie|Doughnut|Area|XYScatter|ColumnStacked), title, position (sel kiri-atas)}
- conditional {address, rule (greater|less|equal|between|notEqual|greaterOrEqual|lessOrEqual|text|duplicate|unique|top|bottom|dataBar|colorScale|formula), value, value2, fill, fontColor, bold}
- validation {address, list: [pilihan dropdown]}
- sort {address, column (indeks 0-based atau nama header), ascending, hasHeaders}
- filter {address}; freeze {rows, cols}; unfreeze {}; autofit {address?}; select {address}
- clear {address, what: All|Contents|Formats}; removeDuplicates {address, columns?, hasHeaders}
- insertRows/deleteRows/insertColumns/deleteColumns {address}
- addSheet {name, activate}; activate {sheet}; renameSheet {sheet?, name}; deleteSheet {sheet}
- name {name, address}: nama rentang; terbilang {source, target, rupiah}: tulis terbilang dari sel source ke target.
Setiap operasi boleh punya "sheet" untuk menargetkan sheet tertentu. Alamat boleh "A1", "B2:D10", atau "'Nama Sheet'!B2:D10".`;

const BLOCKS_DOC = `Blok konten (field "type"):
- heading {text, level 1-3, align}
- paragraph {text, align (left|center|right|justify), bold, italic, size, indent, border}. Teks mendukung **tebal**, *miring*, __garis bawah__, dan baris baru.
- list {items: [...], ordered}
- table {rows: 2D array string, header: true bila baris pertama judul}
- fields {rows: [[label, nilai], ...]}: baris "Nama : nilai" rapi tanpa garis (khas surat resmi Indonesia)
- signature {place, title, name, position, align}: blok tanda tangan; signatures {items:[{title,name,position,note}], place}: beberapa tanda tangan berdampingan
- line {}: garis pemisah kop surat; spacer {}; pageBreak {}; html {html}: HTML mentah bila perlu format khusus.`;

const OPS_WORD_DOC = `Operasi edit (field "op"):
- replace {find, replace, matchCase, wholeWord}; highlight {find, color (kuning|hijau|biru|merah|ungu|abu atau #RRGGBB)}
- format {scope (selection|document), bold, italic, underline, fontName, fontSize, color, highlight, align (Left|Centered|Right|Justified), lineSpacing (kelipatan: 1, 1.15, 1.5, 2), spaceAfter (poin), firstLineIndent (poin), style (Heading1|Heading2|Heading3|Normal|Title|Quote ...), case (upper|lower|title|sentence)}
- header {text, align, bold, fontSize}; header {text, footer: true} untuk footer
- pageNumbers {align, total}; toc {location: cursor|start}; deleteEmptyParagraphs {}; collapseSpaces {}; pageBreak {}
- list {ordered}: jadikan paragraf terpilih daftar; tidy {fontName, fontSize, lineSpacing}; clearDocument {} (HANYA bila pengguna jelas meminta menghapus seluruh isi).`;

const RUN_CODE_TOOL = (host) => ({
  name: 'run_office_js',
  description: `Jalankan kode JavaScript Office.js khusus di dalam ${host === 'excel' ? 'Excel.run' : 'Word.run'} untuk hal yang tidak bisa dilakukan alat lain. Variabel tersedia: context, ${host === 'excel' ? 'Excel' : 'Word'}, Office. Kode adalah isi fungsi async; gunakan await context.sync() dan "return" untuk mengembalikan hasil (akan di-JSON-kan). Pengguna mungkin diminta menyetujui kode ini dulu, jadi pakai hanya bila perlu.`,
  input_schema: {
    type: 'object',
    properties: {
      explanation: { type: 'string', description: 'Penjelasan singkat (bahasa Indonesia) tentang apa yang dilakukan kode' },
      code: { type: 'string', description: 'Isi fungsi async' },
    },
    required: ['explanation', 'code'],
  },
});

const MASTER_TOOL = {
  name: 'master_command',
  description:
    'Jalankan perintah bawaan Master (template siap pakai yang sudah teruji). Gunakan untuk template keuangan/dokumen standar, mis. "simulasi cicilan pinjaman 50 juta bunga 10% 24 bulan", "buat anggaran bulanan", "buat laporan laba rugi", "buat daftar gaji 10 karyawan", "buat absensi bulan desember", "surat izin sakit nama: Budi, alasan: demam". Hasilnya berupa pesan status.',
  input_schema: { type: 'object', properties: { command: { type: 'string' } }, required: ['command'] },
};

export function toolsFor(host) {
  if (host === 'excel') {
    return [
      {
        name: 'excel_info',
        description: 'Info workbook: daftar sheet & area terpakai, sheet aktif, seleksi saat ini (alamat + nilai), daftar tabel. Panggil dulu sebelum mengubah data yang belum Anda lihat.',
        input_schema: { type: 'object', properties: {} },
      },
      {
        name: 'excel_read',
        description: 'Baca nilai dan rumus suatu rentang. Tanpa address = seluruh area terpakai di sheet. Rumus dikembalikan dengan alamat relatif terhadap sel kiri-atas rentang.',
        input_schema: {
          type: 'object',
          properties: {
            sheet: { type: 'string' },
            address: { type: 'string' },
            max_cells: { type: 'integer', description: 'Batas sel yang dibaca (default 3000)' },
          },
        },
      },
      {
        name: 'excel_apply',
        description: `Ubah workbook dengan daftar operasi yang dijalankan berurutan. Error per operasi dilaporkan tanpa menghentikan operasi lain.\n${OPS_EXCEL_DOC}`,
        input_schema: {
          type: 'object',
          properties: {
            sheet: { type: 'string', description: 'Sheet target (default: sheet aktif; dibuat bila belum ada)' },
            ops: { type: 'array', items: { type: 'object', properties: { op: { type: 'string' } }, required: ['op'] } },
          },
          required: ['ops'],
        },
      },
      MASTER_TOOL,
      RUN_CODE_TOOL('excel'),
    ];
  }
  return [
    {
      name: 'word_read',
      description: 'Baca isi dokumen Word: teks seleksi, teks dokumen (bisa dipotong per halaman dengan offset), dan daftar paragraf beserta gaya (style).',
      input_schema: {
        type: 'object',
        properties: { max_chars: { type: 'integer' }, offset: { type: 'integer' } },
      },
    },
    {
      name: 'word_insert',
      description: `Sisipkan konten berformat ke dokumen.\nlocation: cursor (default, di posisi kursor / setelah teks terpilih), replace (ganti teks terpilih), start, end, replace_document (ganti seluruh isi; hanya bila diminta).\n${BLOCKS_DOC}`,
      input_schema: {
        type: 'object',
        properties: {
          location: { type: 'string', enum: ['cursor', 'replace', 'start', 'end', 'replace_document'] },
          blocks: { type: 'array', items: { type: 'object', properties: { type: { type: 'string' } }, required: ['type'] } },
        },
        required: ['blocks'],
      },
    },
    {
      name: 'word_edit',
      description: `Edit dokumen dengan daftar operasi berurutan.\n${OPS_WORD_DOC}`,
      input_schema: {
        type: 'object',
        properties: { ops: { type: 'array', items: { type: 'object', properties: { op: { type: 'string' } }, required: ['op'] } } },
        required: ['ops'],
      },
    },
    MASTER_TOOL,
    RUN_CODE_TOOL('word'),
  ].map((t) => t);
}

export function systemPrompt(host, profile = {}) {
  const app = host === 'excel' ? 'Microsoft Excel' : 'Microsoft Word';
  const prof = Object.entries(profile)
    .filter(([, v]) => v)
    .map(([k, v]) => `${k}: ${v}`)
    .join('; ');
  return `Anda adalah "Master", asisten AI yang bekerja langsung di dalam ${app} pengguna lewat add-in. Pengguna mengetik perintah dalam bahasa sehari-hari (biasanya bahasa Indonesia) dan Anda MENGERJAKANNYA langsung di dokumen memakai alat yang tersedia, bukan sekadar menjelaskan caranya.

Cara kerja:
- Bila perintah menyangkut data/isi yang belum Anda lihat, baca dulu (${host === 'excel' ? 'excel_info / excel_read' : 'word_read'}). Jangan mengarang isi dokumen.
- Kerjakan tugas sampai selesai dalam sesedikit mungkin panggilan alat; gabungkan banyak operasi dalam satu panggilan.
- Untuk template standar (cicilan, anggaran, laba rugi, neraca, gaji, faktur, stok, absensi, nilai, surat-surat resmi) boleh memakai master_command, lalu sesuaikan hasilnya bila perlu.
- Pakai run_office_js hanya bila alat lain tidak mampu.
- Jangan menghapus sheet, mengosongkan dokumen, atau menimpa data yang ada kecuali pengguna jelas memintanya. Bila tidak yakin lokasi, tulis di area kosong.
${host === 'excel'
    ? `- Excel: rumus selalu memakai nama fungsi bahasa Inggris dan pemisah koma (API Office.js tidak bergantung pada bahasa Excel pengguna). Buat rumus yang hidup (bukan angka hasil hitungan sendiri) agar tetap benar saat data berubah. Format uang Indonesia: '"Rp "#,##0'. Tanggal: 'dd/mm/yyyy'. Beri judul tebal, header berwarna, lebar kolom rapi.`
    : `- Word: tulis dengan bahasa Indonesia baku sesuai EYD dan struktur surat resmi Indonesia (kop, nomor/lampiran/perihal, tempat-tanggal, tujuan, salam pembuka, isi, salam penutup, tanda tangan). Gunakan blok "fields" untuk data diri dan "signature" untuk tanda tangan. Bagian yang datanya tidak diketahui diisi titik-titik.`}
- Setelah selesai, balas SINGKAT (1-3 kalimat) apa yang sudah dikerjakan dan di mana hasilnya. Jika pengguna hanya bertanya, jawab langsung dengan ringkas.
${prof ? `\nProfil pengguna (pakai untuk mengisi surat/dokumen bila relevan): ${prof}` : ''}`;
}

/* ------------------------------ loop agen ------------------------------ */

export class OfflineError extends Error {}

/** Membuang blok sebelum batas fallback (sesuai aturan echo fallback). */
function echoContent(content) {
  let last = -1;
  content.forEach((b, i) => {
    if (b.type === 'fallback') last = i;
  });
  if (last < 0) return content;
  const keepTypes = new Set(['text']);
  return content.filter((b, i) => i > last || keepTypes.has(b.type));
}

export class MasterAgent {
  /**
   * opts: { apiKey, model, effort, host, profile, handlers: {toolName: async(input)=>any}, onText, onTool }
   */
  constructor(opts) {
    this.opts = opts;
    this.messages = [];
    this.client = new Anthropic({ apiKey: opts.apiKey, dangerouslyAllowBrowser: true, maxRetries: 2 });
    this.tools = toolsFor(opts.host).map((t) => Object.assign({}, t, { eager_input_streaming: true }));
    this.schemas = Object.fromEntries(this.tools.map((t) => [t.name, t.input_schema]));
  }

  reset() {
    this.messages = [];
  }

  async run(userText, contextNote) {
    const { model, effort, host, profile, handlers, onText, onTool, signal } = this.opts;
    const text = contextNote ? `${userText}\n\n<konteks_dokumen>\n${contextNote}\n</konteks_dokumen>` : userText;
    const snapshot = this.messages.map((m) => Object.assign({}, m, { content: Array.isArray(m.content) ? m.content.slice() : m.content }));
    const last = this.messages[this.messages.length - 1];
    if (last && last.role === 'user') last.content = [...last.content, { type: 'text', text }];
    else this.messages.push({ role: 'user', content: [{ type: 'text', text }] });
    const base = {
      model: model || DEFAULT_MODEL,
      max_tokens: 32000,
      system: systemPrompt(host, profile),
      tools: this.tools,
      output_config: { effort: effort || 'medium' },
      cache_control: { type: 'ephemeral' },
    };
    let jsonRetries = 0;
    let finalText = '';
    for (let turn = 0; turn < 25; turn++) {
      const params = Object.assign({}, base, { messages: this.messages });
      if (supportsFallback(params.model)) {
        params.betas = [FALLBACK_BETA];
        params.fallbacks = 'default';
      }
      let message;
      try {
        const stream = supportsFallback(params.model) ? this.client.beta.messages.stream(params, { signal }) : this.client.messages.stream(params, { signal });
        stream.on('text', (d) => onText && onText(d));
        message = await stream.finalMessage();
        jsonRetries = 0;
      } catch (err) {
        if (err instanceof Anthropic.APIConnectionError && !(err instanceof Anthropic.APIConnectionTimeoutError)) {
          if (turn === 0) this.messages = snapshot;
          const e = new OfflineError(turn === 0 ? 'Tidak dapat terhubung ke Claude (offline?).' : 'Koneksi terputus di tengah pekerjaan; sebagian langkah mungkin sudah dijalankan.');
          e.partial = turn > 0;
          throw e;
        }
        if (err instanceof Anthropic.APIError || jsonRetries++ >= 2 || (signal && signal.aborted)) {
          if (turn === 0) this.messages = snapshot;
          throw err;
        }
        continue; // input alat tidak bisa di-parse: ulangi giliran ini
      }

      const content = echoContent(message.content);
      const textNow = content.filter((b) => b.type === 'text').map((b) => b.text).join('');
      if (textNow) finalText = textNow;

      if (message.stop_reason === 'refusal') {
        this.messages.push({ role: 'assistant', content: content.filter((b) => b.type === 'text').length ? content.filter((b) => b.type === 'text') : [{ type: 'text', text: '(ditolak)' }] });
        return { text: finalText || 'Maaf, permintaan ini tidak dapat diproses.', refused: true };
      }
      if (!content.length) content.push({ type: 'text', text: '(kosong)' });
      if (message.stop_reason === 'pause_turn') {
        this.messages.push({ role: 'assistant', content });
        continue;
      }
      const toolUses = content.filter((b) => b.type === 'tool_use');
      if (message.stop_reason === 'max_tokens' && toolUses.length) {
        // Input alat bisa terpotong: jangan dijalankan.
        this.messages.push({ role: 'assistant', content: content.filter((b) => b.type !== 'tool_use') });
        throw new Error('Jawaban AI terpotong (max_tokens). Coba bagi perintah menjadi bagian yang lebih kecil.');
      }
      this.messages.push({ role: 'assistant', content });
      if (message.stop_reason !== 'tool_use' || !toolUses.length) {
        return { text: finalText, usage: message.usage };
      }

      const results = [];
      for (const tu of toolUses) {
        const errors = validate(this.schemas[tu.name], tu.input);
        if (!this.schemas[tu.name]) errors.push(`alat ${tu.name} tidak dikenal`);
        if (errors.length) {
          results.push({ type: 'tool_result', tool_use_id: tu.id, is_error: true, content: `Input tidak valid: ${errors.join('; ')}` });
          continue;
        }
        if (onTool) onTool(tu.name, tu.input, 'start');
        try {
          const out = await handlers[tu.name](tu.input);
          const s = typeof out === 'string' ? out : JSON.stringify(out === undefined ? { ok: true } : out);
          results.push({ type: 'tool_result', tool_use_id: tu.id, content: s.length > 60000 ? s.slice(0, 60000) + '…(dipotong)' : s });
          if (onTool) onTool(tu.name, tu.input, 'done', out);
        } catch (e) {
          results.push({ type: 'tool_result', tool_use_id: tu.id, is_error: true, content: String((e && e.message) || e) });
          if (onTool) onTool(tu.name, tu.input, 'error', e);
        }
      }
      this.messages.push({ role: 'user', content: results });
    }
    return { text: finalText || 'Batas langkah tercapai. Ketik "lanjutkan" untuk meneruskan.' };
  }
}

/** Menerjemahkan error API menjadi pesan yang mudah dipahami. */
export function friendlyError(err) {
  if (err instanceof OfflineError) return err.message;
  if (err instanceof Anthropic.AuthenticationError) return 'API key tidak valid. Periksa di tab Pengaturan.';
  if (err instanceof Anthropic.PermissionDeniedError) return 'API key tidak memiliki akses ke model ini.';
  if (err instanceof Anthropic.RateLimitError) return 'Terlalu banyak permintaan / batas pemakaian tercapai. Coba lagi sebentar.';
  if (err instanceof Anthropic.BadRequestError) return `Permintaan ditolak API: ${err.message}`;
  if (err instanceof Anthropic.InternalServerError) return 'Server Claude sedang sibuk. Coba lagi.';
  if (err instanceof Anthropic.APIUserAbortError) return 'Dibatalkan.';
  if (err instanceof Anthropic.APIError) return `Error API (${err.status}): ${err.message}`;
  return String((err && err.message) || err);
}

/** Cek cepat API key (permintaan sangat kecil). */
export async function testApiKey(apiKey, model) {
  const client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true, maxRetries: 0 });
  const msg = await client.messages.create({
    model: model || DEFAULT_MODEL,
    max_tokens: 64,
    output_config: { effort: 'low' },
    messages: [{ role: 'user', content: 'Balas: OK' }],
  });
  return msg.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
}
