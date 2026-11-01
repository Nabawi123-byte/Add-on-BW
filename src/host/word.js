/* global Word, Office */
// Executor Word: menyisipkan blok (surat, tabel, judul) dan mengedit dokumen lewat Office.js.

import { renderSegments } from '../core/word-render.js';
import { WORD_TEMPLATES } from '../core/word-templates.js';
import { tanggalIndonesia, titleCase, escapeHtml } from '../core/util.js';
import { terbilang, terbilangRupiah } from '../core/terbilang.js';

const HIGHLIGHT = { kuning: '#FFFF00', hijau: '#00FF00', biru: '#00FFFF', merah: '#FF0000', ungu: '#FF00FF', abu: '#C0C0C0', yellow: '#FFFF00', green: '#00FF00' };
const PAGE_BREAK_HTML = '<p style="page-break-before:always;margin:0;"></p>';

/* --------------------------- penyisipan blok --------------------------- */

/**
 * location: 'cursor' (default), 'replace' (ganti seleksi), 'start', 'end', 'replace_document'
 */
export async function insertBlocks(blocks, location = 'cursor') {
  const html = renderSegments(blocks).join(PAGE_BREAK_HTML);
  return Word.run(async (context) => {
    const body = context.document.body;
    let inserted;
    if (location === 'end') inserted = body.insertHtml(html, 'End');
    else if (location === 'start') inserted = body.insertHtml(html, 'Start');
    else if (location === 'replace_document') {
      body.clear();
      inserted = body.insertHtml(html, 'Start');
    } else {
      const sel = context.document.getSelection();
      sel.load('text');
      await context.sync();
      const hasText = sel.text && sel.text.trim().length > 0;
      inserted = sel.insertHtml(html, location === 'replace' || !hasText ? 'Replace' : 'After');
    }
    inserted.select('End');
    await context.sync();
    return { message: `✅ ${blocks.length} bagian disisipkan.` };
  });
}

function ooxmlPackage(bodyXml) {
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}</w:body></w:document></pkg:xmlData></pkg:part>
</pkg:package>`;
}

const fld = (instr, placeholder) => `<w:fldSimple w:instr="${instr}"><w:r><w:t xml:space="preserve">${placeholder}</w:t></w:r></w:fldSimple>`;
const run = (t) => `<w:r><w:t xml:space="preserve">${escapeHtml(t)}</w:t></w:r>`;

/* ------------------------------ util ------------------------------ */

async function targetRange(context, scope) {
  if (scope === 'document') return context.document.body.getRange('Whole');
  const sel = context.document.getSelection();
  sel.load('text');
  await context.sync();
  if (!sel.text || !sel.text.trim()) return sel.paragraphs.getFirst().getRange('Whole');
  return sel;
}

function caseFn(mode) {
  return {
    upper: (s) => s.toUpperCase(),
    lower: (s) => s.toLowerCase(),
    title: titleCase,
    sentence: (s) => s.toLowerCase().replace(/(^\s*|[.!?]\s+)([a-z])/g, (m, a, b) => a + b.toUpperCase()),
  }[mode];
}

async function replaceAll(context, find, replace, opts = {}) {
  if (!find) throw new Error('Teks yang dicari kosong');
  if (find.length > 255) throw new Error('Teks yang dicari maksimal 255 karakter');
  const results = context.document.body.search(find, { matchCase: !!opts.matchCase, matchWholeWord: !!opts.wholeWord });
  results.load('items');
  await context.sync();
  results.items.forEach((r) => r.insertText(replace, 'Replace'));
  await context.sync();
  return results.items.length;
}

/* ------------------------- operasi untuk AI ------------------------- */

const OP_HANDLERS = {
  insert: async (context, op) => {
    await insertBlocks(op.blocks || [], op.location || 'cursor');
    return 'teks disisipkan';
  },
  replace: async (context, op) => `${await replaceAll(context, op.find, op.replace === undefined ? '' : op.replace, op)} penggantian`,
  highlight: async (context, op) => {
    const results = context.document.body.search(op.find, { matchCase: !!op.matchCase, matchWholeWord: !!op.wholeWord });
    results.load('items');
    await context.sync();
    const color = HIGHLIGHT[op.color] || op.color || '#FFFF00';
    results.items.forEach((r) => (r.font.highlightColor = color));
    return `${results.items.length} ditandai`;
  },
  format: async (context, op) => {
    const range = await targetRange(context, op.scope);
    const f = range.font;
    if (op.bold !== undefined) f.bold = !!op.bold;
    if (op.italic !== undefined) f.italic = !!op.italic;
    if (op.underline !== undefined) f.underline = op.underline ? 'Single' : 'None';
    if (op.fontName) f.name = op.fontName;
    if (op.fontSize) f.size = Number(op.fontSize);
    if (op.color) f.color = op.color;
    if (op.highlight !== undefined) f.highlightColor = op.highlight ? HIGHLIGHT[op.highlight] || op.highlight : null;
    if (op.align || op.lineSpacing || op.spaceAfter !== undefined || op.firstLineIndent !== undefined || op.style) {
      const paras = range.paragraphs;
      paras.load('items');
      await context.sync();
      for (const p of paras.items) {
        if (op.align) p.alignment = op.align;
        if (op.lineSpacing) p.lineSpacing = Number(op.lineSpacing) * 12;
        if (op.spaceAfter !== undefined) p.spaceAfter = Number(op.spaceAfter);
        if (op.firstLineIndent !== undefined) p.firstLineIndent = Number(op.firstLineIndent);
        if (op.style) p.styleBuiltIn = op.style;
      }
    }
    if (op.case) {
      const fn = caseFn(op.case);
      if (!fn) throw new Error('mode huruf tidak dikenal');
      if (op.scope === 'document') {
        const paras = context.document.body.paragraphs;
        paras.load('items/text,items/tableNestingLevel');
        await context.sync();
        paras.items.forEach((p) => {
          if (p.text && !p.tableNestingLevel) p.insertText(fn(p.text), 'Replace');
        });
      } else {
        range.load('text');
        await context.sync();
        range.insertText(fn(range.text), 'Replace');
      }
    }
    return 'format diterapkan';
  },
  header: async (context, op) => {
    const sections = context.document.sections;
    sections.load('items');
    await context.sync();
    for (const s of sections.items) {
      const h = op.footer ? s.getFooter('Primary') : s.getHeader('Primary');
      h.clear();
      const p = h.insertParagraph(op.text || '', 'Start');
      p.alignment = op.align || 'Centered';
      if (op.bold) p.font.bold = true;
      if (op.fontSize) p.font.size = Number(op.fontSize);
    }
    return op.footer ? 'footer dipasang' : 'header dipasang';
  },
  pageNumbers: async (context, op) => {
    const sections = context.document.sections;
    sections.load('items');
    await context.sync();
    const xml = ooxmlPackage(`<w:p><w:pPr><w:jc w:val="${op.align || 'center'}"/></w:pPr>${run('Halaman ')}${fld(' PAGE ', '1')}${op.total === false ? '' : run(' dari ') + fld(' NUMPAGES ', '1')}</w:p>`);
    for (const s of sections.items) {
      const footer = s.getFooter('Primary');
      footer.clear();
      footer.insertOoxml(xml, 'Start');
    }
    return 'nomor halaman dipasang';
  },
  toc: async (context, op) => {
    const xml = ooxmlPackage(
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>DAFTAR ISI</w:t></w:r></w:p><w:p>${fld(' TOC \\o &quot;1-3&quot; \\h \\z \\u ', 'Klik kanan di sini lalu pilih &quot;Update Field&quot; (atau tekan F9) untuk menampilkan daftar isi.')}</w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p>`,
    );
    const where = op.location === 'cursor' ? context.document.getSelection() : context.document.body;
    where.insertOoxml(xml, op.location === 'cursor' ? 'Before' : 'Start');
    return 'daftar isi disisipkan';
  },
  deleteEmptyParagraphs: async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load('items/text,items/tableNestingLevel');
    await context.sync();
    const pics = paras.items.map((p) => {
      const c = p.inlinePictures;
      c.load('items');
      return c;
    });
    await context.sync();
    let n = 0;
    paras.items.forEach((p, i) => {
      if (i === paras.items.length - 1) return;
      if (!p.tableNestingLevel && !p.text.trim() && !pics[i].items.length) {
        p.delete();
        n++;
      }
    });
    return `${n} paragraf kosong dihapus`;
  },
  collapseSpaces: async (context) => {
    let total = 0;
    for (let i = 0; i < 10; i++) {
      const n = await replaceAll(context, '  ', ' ');
      total += n;
      if (!n) break;
    }
    return `${total} spasi ganda dirapikan`;
  },
  pageBreak: async (context) => {
    context.document.getSelection().insertBreak('Page', 'After');
    return 'halaman baru';
  },
  list: async (context, op) => {
    const sel = context.document.getSelection();
    const paras = sel.paragraphs;
    paras.load('items/text,items/isListItem');
    await context.sync();
    const items = paras.items.filter((p) => p.text.trim());
    if (!items.length) throw new Error('Pilih paragraf yang ingin dijadikan daftar.');
    const list = items[0].startNewList();
    list.load('id');
    await context.sync();
    for (const p of items.slice(1)) p.attachToList(list.id, 0);
    if (op.ordered) list.setLevelNumbering(0, 'Arabic', [0, '.']);
    else list.setLevelBullet(0, 'Solid');
    return `${items.length} paragraf dijadikan daftar`;
  },
  tidy: async (context, op) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load('items/styleBuiltIn,items/alignment,items/tableNestingLevel,items/text');
    await context.sync();
    for (const p of paras.items) {
      const isNormal = p.styleBuiltIn === 'Normal' || p.styleBuiltIn === 'Other';
      if (!isNormal || p.tableNestingLevel) continue;
      p.font.name = op.fontName || 'Times New Roman';
      p.font.size = Number(op.fontSize) || 12;
      p.lineSpacing = (Number(op.lineSpacing) || 1.5) * 12;
      p.spaceAfter = 6;
      if (p.alignment === 'Left' && p.text.length > 80) p.alignment = 'Justified';
    }
    await context.sync();
    const n = await OP_HANDLERS.collapseSpaces(context);
    return `dokumen dirapikan (${n})`;
  },
  clearDocument: async (context) => {
    context.document.body.clear();
    return 'dokumen dikosongkan';
  },
};

export const WORD_OPS = Object.keys(OP_HANDLERS);

export async function applyWordOps(ops) {
  const done = [];
  const errors = [];
  for (const op of ops || []) {
    const h = OP_HANDLERS[op && op.op];
    if (!h) {
      errors.push(`operasi "${op && op.op}" tidak dikenal`);
      continue;
    }
    try {
      const info = op.op === 'insert' ? await h(null, op) : await Word.run(async (context) => {
        const r = await h(context, op);
        await context.sync();
        return r;
      });
      done.push(info || op.op);
    } catch (e) {
      errors.push(`${op.op}: ${(e && e.message) || e}`);
    }
  }
  return { done, errors };
}

async function one(op) {
  const res = await applyWordOps([op]);
  if (res.errors.length) throw new Error(res.errors[0]);
  return res.done[0];
}

/* -------------------------- perintah offline -------------------------- */

function signatureBlock(profile) {
  return {
    type: 'signature',
    place: `${profile.kota || ''}, ${tanggalIndonesia()}`,
    title: 'Hormat saya,',
    name: profile.nama || '(..............................)',
    position: profile.jabatan || '',
  };
}

const ACTIONS = {
  async template({ id, values }) {
    const tpl = WORD_TEMPLATES.find((t) => t.id === id);
    if (!tpl) throw new Error(`Template ${id} tidak ditemukan`);
    await insertBlocks(tpl.build(values), 'cursor');
    return { message: `✅ ${tpl.title} disisipkan. Periksa kembali bagian bertitik-titik.` };
  },
  async table({ headers, cols, rows = 5, name }) {
    const hs = headers && headers.length ? headers.map((h) => titleCase(h)) : Array.from({ length: Math.min(cols || 3, 30) }, (_, i) => `Kolom ${i + 1}`);
    const body = Array.from({ length: Math.min(rows || 5, 500) }, (_, i) => hs.map((h, j) => (j === 0 && /^no/i.test(h) ? String(i + 1) : ' ')));
    const blocks = [];
    if (name) blocks.push({ type: 'heading', level: 3, text: titleCase(name) });
    blocks.push({ type: 'table', header: true, rows: [hs, ...body] });
    await insertBlocks(blocks, 'cursor');
    return { message: `✅ Tabel ${hs.length} kolom × ${body.length} baris disisipkan.` };
  },
  async replace({ find, replace }) {
    const n = await Word.run((context) => replaceAll(context, find, replace));
    return { message: n ? `✅ ${n} kemunculan "${find}" diganti menjadi "${replace}".` : `ℹ️ "${find}" tidak ditemukan.` };
  },
  async highlight({ find, color }) {
    const info = await one({ op: 'highlight', find, color });
    return { message: `✅ "${find}": ${info}.` };
  },
  async terbilang({ number, rupiah }) {
    if (number === null || number === undefined) {
      const text = await Word.run(async (context) => {
        const sel = context.document.getSelection();
        sel.load('text');
        await context.sync();
        return sel.text;
      });
      const m = String(text || '').match(/[\d.,]+/);
      if (!m) throw new Error('Ketik angkanya, misalnya "terbilang 1.500.000", atau pilih teks angka.');
      number = Number(m[0].replace(/\./g, '').replace(',', '.'));
    }
    const t = rupiah ? terbilangRupiah(number) : terbilang(number);
    await insertBlocks([{ type: 'paragraph', text: t }], 'cursor');
    return { message: `✅ ${t}` };
  },
  async date() {
    await Word.run(async (context) => {
      context.document.getSelection().insertText(tanggalIndonesia(), 'Replace');
      await context.sync();
    });
    return { message: `✅ ${tanggalIndonesia()}` };
  },
  async count() {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      const paras = body.paragraphs;
      paras.load('items/text');
      const sel = context.document.getSelection();
      sel.load('text');
      await context.sync();
      const words = (t) => (String(t).match(/\S+/g) || []).length;
      const nPara = paras.items.filter((p) => p.text.trim()).length;
      let msg = `📊 Dokumen: ${words(body.text).toLocaleString('id-ID')} kata, ${body.text.replace(/\s/g, '').length.toLocaleString('id-ID')} karakter (tanpa spasi), ${nPara} paragraf.`;
      if (sel.text && sel.text.trim()) msg += ` Seleksi: ${words(sel.text)} kata.`;
      return { message: msg };
    });
  },
  async cleanSpaces() {
    return { message: `✅ ${await one({ op: 'collapseSpaces' })}.` };
  },
  async removeEmpty() {
    return { message: `✅ ${await one({ op: 'deleteEmptyParagraphs' })}.` };
  },
  async toc() {
    await one({ op: 'toc', location: 'cursor' });
    return { message: '✅ Daftar isi disisipkan. Klik kanan → Update Field (F9) untuk memperbarui. Daftar isi membaca paragraf bergaya Judul/Heading.' };
  },
  async pageNumbers() {
    await one({ op: 'pageNumbers' });
    return { message: '✅ Nomor halaman "Halaman X dari Y" dipasang di footer.' };
  },
  async pageBreak() {
    await one({ op: 'pageBreak' });
    return { message: '✅ Halaman baru disisipkan.' };
  },
  async header({ text }) {
    await one({ op: 'header', text, bold: true });
    return { message: '✅ Header dipasang.' };
  },
  async footer({ text }) {
    await one({ op: 'header', text, footer: true, fontSize: 9 });
    return { message: '✅ Footer dipasang.' };
  },
  async signature(params, env) {
    await insertBlocks([signatureBlock(env.profile || {})], 'cursor');
    return { message: '✅ Blok tanda tangan disisipkan.' };
  },
  async heading({ text }) {
    await insertBlocks([{ type: 'heading', level: 1, text, align: 'center' }], 'cursor');
    return { message: '✅ Judul disisipkan.' };
  },
  async applyHeading({ level }) {
    await one({ op: 'format', style: `Heading${Math.min(Math.max(level || 1, 1), 3)}` });
    return { message: `✅ Paragraf dijadikan Judul ${level || 1}.` };
  },
  async list({ ordered }) {
    const info = await one({ op: 'list', ordered });
    return { message: `✅ ${info}.` };
  },
  async case({ mode, scope }) {
    await one({ op: 'format', case: mode, scope });
    return { message: '✅ Huruf diubah.' };
  },
  async spacing({ value, scope }) {
    await one({ op: 'format', lineSpacing: value, scope });
    return { message: `✅ Spasi ${value} diterapkan${scope === 'document' ? ' ke seluruh dokumen' : ''}.` };
  },
  async align({ align, scope }) {
    await one({ op: 'format', align, scope });
    return { message: '✅ Perataan diterapkan.' };
  },
  async font({ name, size, scope }) {
    await one({ op: 'format', fontName: name, fontSize: size, scope });
    return { message: `✅ Font ${[name, size ? size + ' pt' : ''].filter(Boolean).join(' ')} diterapkan${scope === 'document' ? ' ke seluruh dokumen' : ''}.` };
  },
  async style(params) {
    await one(Object.assign({ op: 'format' }, params));
    return { message: '✅ Gaya teks diterapkan.' };
  },
  async tidy() {
    const info = await one({ op: 'tidy' });
    return { message: `✅ ${info}: Times New Roman 12, spasi 1,5, paragraf panjang rata kiri-kanan.` };
  },
};

export async function runWordAction(action, params, env = {}) {
  const name = action.replace(/^word\./, '');
  const fn = ACTIONS[name];
  if (!fn) throw new Error(`Aksi Word "${action}" belum tersedia`);
  return fn(params || {}, env);
}

/* ---------------------------- alat untuk AI ---------------------------- */

export async function readDocument({ maxChars = 20000, offset = 0, includeParagraphs = true } = {}) {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load('text');
    const sel = context.document.getSelection();
    sel.load('text');
    const paras = body.paragraphs;
    paras.load('items/text,items/styleBuiltIn,items/alignment,items/tableNestingLevel,items/isListItem');
    await context.sync();
    const text = body.text || '';
    const cap = Math.min(Number(maxChars) || 20000, 100000);
    const out = {
      totalChars: text.length,
      words: (text.match(/\S+/g) || []).length,
      paragraphCount: paras.items.length,
      selection: sel.text ? sel.text.slice(0, 5000) : '',
      text: text.slice(offset, offset + cap),
      truncated: offset + cap < text.length,
    };
    if (includeParagraphs) {
      out.paragraphs = paras.items.slice(0, 300).map((p, i) => ({
        i,
        style: p.styleBuiltIn,
        align: p.alignment,
        inTable: p.tableNestingLevel > 0 || undefined,
        list: p.isListItem || undefined,
        text: p.text.length > 160 ? p.text.slice(0, 160) + '…' : p.text,
      }));
    }
    return out;
  });
}

export async function runCustomWord(code) {
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
  return Word.run(async (context) => {
    const fn = new AsyncFunction('context', 'Word', 'Office', code);
    const result = await fn(context, Word, Office);
    await context.sync();
    return result;
  });
}
