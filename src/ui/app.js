/* global Office */
// Panel Master: kotak perintah, template, makro, pengaturan.

import { parseChain, prefersAI, HELP, canonical } from '../core/parser.js';
import { WORD_TEMPLATES, templateDefaults } from '../core/word-templates.js';
import { EXCEL_TEMPLATES } from '../core/excel-templates.js';
import { BULAN, escapeHtml, parseNumber } from '../core/util.js';
import { MasterAgent, MODELS, friendlyError, OfflineError, testApiKey } from '../core/ai.js';
import * as X from '../host/excel.js';
import * as W from '../host/word.js';
import { load, save, getSettings, getProfile } from './store.js';

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

const state = {
  host: null, // 'excel' | 'word' | null (pratinjau di browser)
  agent: null,
  agentKey: '',
  busy: false,
  abort: null,
  historyIdx: -1,
  online: navigator.onLine,
};

/* ------------------------------ tampilan log ------------------------------ */

function md(text) {
  const lines = escapeHtml(text).split(/\n/);
  return lines
    .map((l) =>
      l
        .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
        .replace(/`([^`]+)`/g, '<code>$1</code>')
        .replace(/^\s*[-•]\s+(.*)$/, '<span class="li">• $1</span>'),
    )
    .join('<br>');
}

function addMsg(role, text, opts = {}) {
  const log = $('#log');
  const el = document.createElement('div');
  el.className = `msg ${role}`;
  if (opts.html) el.innerHTML = opts.html;
  else el.innerHTML = md(text || '');
  log.appendChild(el);
  log.scrollTop = log.scrollHeight;
  return el;
}

function scrollLog() {
  const log = $('#log');
  log.scrollTop = log.scrollHeight;
}

function setBusy(b) {
  state.busy = b;
  $('#sendBtn').hidden = b;
  $('#stopBtn').hidden = !b;
  $('#input').disabled = false;
  document.body.classList.toggle('busy', b);
}

/* ------------------------------ status jaringan ------------------------------ */

function updateNet() {
  const s = getSettings();
  const badge = $('#netBadge');
  const hasKey = !!s.apiKey;
  let label;
  let cls;
  if (s.mode === 'offline') {
    label = 'Mode Offline';
    cls = 'off';
  } else if (!hasKey) {
    label = 'Offline (tanpa API key)';
    cls = 'off';
  } else if (state.online) {
    label = 'AI Online';
    cls = 'on';
  } else {
    label = 'Tidak ada internet';
    cls = 'warn';
  }
  badge.textContent = label;
  badge.className = `badge ${cls}`;
  renderQueue();
}

window.addEventListener('online', () => {
  state.online = true;
  updateNet();
  const q = queueForHost();
  if (q.length) addMsg('sys', `🌐 Internet tersambung. Ada ${q.length} perintah di antrean — klik **Jalankan antrean**.`);
});
window.addEventListener('offline', () => {
  state.online = false;
  updateNet();
});

/* ------------------------------ antrean offline ------------------------------ */

function queueForHost() {
  return load('queue', []).filter((q) => q.host === state.host);
}

function renderQueue() {
  const q = queueForHost();
  $('#queueBar').hidden = !q.length;
  $('#queueCount').textContent = String(q.length);
}

function enqueue(text) {
  const q = load('queue', []);
  q.push({ text, host: state.host, time: Date.now() });
  save('queue', q.slice(-50));
  renderQueue();
}

async function runQueue() {
  const all = load('queue', []);
  const mine = all.filter((q) => q.host === state.host);
  save('queue', all.filter((q) => q.host !== state.host));
  renderQueue();
  for (const item of mine) {
    addMsg('user', item.text);
    const ok = await runAI(item.text);
    if (!ok) break;
  }
}

/* ------------------------------ eksekusi perintah ------------------------------ */

function env() {
  return { profile: getProfile(), settings: getSettings() };
}

function reportResult(res) {
  if (!res) return;
  if (res.message) addMsg('bot', res.message);
  if (res.errors && res.errors.length) {
    addMsg('err', `⚠️ ${res.errors.length} langkah gagal:\n- ${res.errors.slice(0, 6).join('\n- ')}`);
  }
}

async function executeOffline(cmd, { interactive = true } = {}) {
  const { action, params } = cmd;
  if (action === 'help') return showHelp();
  if (action === 'templates') return switchTab('templates');
  if (action === 'undoHint') {
    return addMsg('sys', state.host === 'word' ? 'Tekan **Ctrl+Z** di Word untuk membatalkan perubahan terakhir.' : 'Coba **Ctrl+Z**. Sebagian perubahan dari add-in tidak bisa di-undo di Excel; bila perlu tutup tanpa menyimpan. Simpan dokumen sebelum perintah besar.');
  }
  if (action === 'macro.run') return runMacro(params.name);
  if (action === 'word.template') {
    const tpl = WORD_TEMPLATES.find((t) => t.id === params.id);
    const values = templateDefaults(tpl, getProfile(), params.given || {});
    if (interactive) return showWordTemplateForm(tpl, values, $('#log'));
    return reportResult(await W.runWordAction('word.template', { id: tpl.id, values }, env()));
  }
  if (!state.host) throw new Error('Buka panel ini dari Word atau Excel untuk menjalankan perintah.');
  if (action.startsWith('excel.')) {
    if (state.host !== 'excel') throw new Error('Perintah ini khusus Excel.');
    return reportResult(await X.runExcelAction(action, params, env()));
  }
  if (action.startsWith('word.')) {
    if (state.host !== 'word') throw new Error('Perintah ini khusus Word.');
    return reportResult(await W.runWordAction(action, params, env()));
  }
  throw new Error(`Aksi ${action} tidak dikenal`);
}

async function runOffline(chain) {
  for (const cmd of chain) {
    if (chain.length > 1) addMsg('sys', `▶ ${cmd.label}`);
    try {
      await executeOffline(cmd, { interactive: chain.length === 1 });
    } catch (e) {
      addMsg('err', `❌ ${friendlyOfficeError(e)}`);
      return false;
    }
  }
  return true;
}

function friendlyOfficeError(e) {
  const m = String((e && e.message) || e);
  if (/InvalidArgument|invalid argument/i.test(m)) return `Argumen tidak valid untuk Office (${m}). Coba pilih sel/teks yang sesuai.`;
  if (/ApiNotFound|is not a function|not supported/i.test(m)) return `Fitur ini belum didukung versi Office Anda (${m}). Perbarui Office ke versi terbaru.`;
  if (/Excel is not defined|Word is not defined/.test(m)) return 'Panel dibuka di luar Word/Excel. Buka dari Office.';
  return m;
}

/* ------------------------------ AI ------------------------------ */

function getAgent() {
  const s = getSettings();
  const key = [s.apiKey, s.model, s.effort, state.host, JSON.stringify(getProfile())].join('|');
  if (!state.agent || state.agentKey !== key) {
    const prev = state.agent;
    state.agent = new MasterAgent({
      apiKey: s.apiKey,
      model: s.model,
      effort: s.effort,
      host: state.host,
      profile: getProfile(),
      handlers: toolHandlers(),
    });
    if (prev && prev.opts.host === state.host && prev.opts.model === s.model) state.agent.messages = prev.messages;
    state.agentKey = key;
  }
  return state.agent;
}

function confirmModal(title, bodyHtml, okLabel = 'Jalankan') {
  return new Promise((resolve) => {
    $('#modalTitle').textContent = title;
    $('#modalBody').innerHTML = bodyHtml;
    $('#modalOk').textContent = okLabel;
    $('#modal').hidden = false;
    const done = (v) => {
      $('#modal').hidden = true;
      $('#modalOk').onclick = null;
      $('#modalCancel').onclick = null;
      resolve(v);
    };
    $('#modalOk').onclick = () => done(true);
    $('#modalCancel').onclick = () => done(false);
  });
}

async function masterCommandForAI(command) {
  const chain = parseChain(command, state.host);
  if (!chain) return { ok: false, message: `Perintah "${command}" tidak dikenali mesin bawaan. Kerjakan dengan alat lain.` };
  const out = [];
  for (const cmd of chain) {
    if (cmd.action === 'word.template') {
      const tpl = WORD_TEMPLATES.find((t) => t.id === cmd.params.id);
      const values = templateDefaults(tpl, getProfile(), cmd.params.given || {});
      out.push(await W.runWordAction('word.template', { id: tpl.id, values }, env()));
    } else if (cmd.action.startsWith('excel.')) out.push(await X.runExcelAction(cmd.action, cmd.params, env()));
    else if (cmd.action.startsWith('word.')) out.push(await W.runWordAction(cmd.action, cmd.params, env()));
    else out.push({ message: `Aksi ${cmd.action} tidak tersedia untuk AI.` });
  }
  return out;
}

function toolHandlers() {
  return {
    excel_info: () => X.getWorkbookInfo(),
    excel_read: (i) => X.readRange({ sheet: i.sheet, address: i.address, maxCells: i.max_cells }),
    excel_apply: (i) => X.applyOps({ sheet: i.sheet, ops: i.ops }),
    word_read: (i) => W.readDocument({ maxChars: i.max_chars, offset: i.offset }),
    word_insert: (i) => W.insertBlocks(i.blocks, i.location),
    word_edit: (i) => W.applyWordOps(i.ops),
    master_command: (i) => masterCommandForAI(i.command),
    run_office_js: async (i) => {
      const policy = getSettings().codePolicy;
      if (policy === 'never') throw new Error('Pengguna menonaktifkan kode khusus. Gunakan alat lain.');
      if (policy === 'ask') {
        const ok = await confirmModal(
          'Izinkan kode khusus?',
          `<p>${escapeHtml(i.explanation)}</p><details><summary>Lihat kode</summary><pre>${escapeHtml(i.code)}</pre></details>`,
        );
        if (!ok) throw new Error('Pengguna menolak menjalankan kode ini.');
      }
      const res = state.host === 'excel' ? await X.runCustomExcel(i.code) : await W.runCustomWord(i.code);
      return res === undefined ? { ok: true } : res;
    },
  };
}

const TOOL_LABEL = {
  excel_info: '🔎 Melihat workbook',
  excel_read: '📖 Membaca data',
  excel_apply: '🛠️ Mengubah sheet',
  word_read: '📖 Membaca dokumen',
  word_insert: '✍️ Menulis ke dokumen',
  word_edit: '🛠️ Mengedit dokumen',
  master_command: '⚡ Perintah bawaan',
  run_office_js: '🧩 Kode khusus',
};

async function contextNote() {
  try {
    if (state.host === 'excel') {
      const info = await X.getWorkbookInfo();
      const sel = info.selection;
      const preview = (sel.values || []).slice(0, 8).map((r) => r.slice(0, 8));
      return `Sheet aktif: ${info.activeSheet}. Seleksi: ${sel.address} (${sel.rows}×${sel.columns}). Pratinjau seleksi: ${JSON.stringify(preview)}. Sheet: ${info.sheets.map((s) => `${s.name}${s.usedRange ? ' [' + s.usedRange.split('!').pop() + ']' : ' [kosong]'}`).join(', ')}.`;
    }
    if (state.host === 'word') {
      const d = await W.readDocument({ maxChars: 0, includeParagraphs: false });
      return `Dokumen: ${d.words} kata, ${d.paragraphCount} paragraf. Teks terpilih: ${d.selection ? JSON.stringify(d.selection.slice(0, 1500)) : '(tidak ada, kursor saja)'}.`;
    }
  } catch (e) {
    return '';
  }
  return '';
}

async function runAI(text) {
  const s = getSettings();
  if (!s.apiKey) {
    addMsg('err', 'Perintah ini butuh AI. Isi **API key** di tab Pengaturan, atau gunakan perintah offline (ketik "bantuan").');
    return false;
  }
  if (!state.host) {
    addMsg('err', 'Buka panel ini dari Word atau Excel.');
    return false;
  }
  if (!state.online) {
    enqueue(text);
    addMsg('sys', '📥 Sedang offline. Perintah disimpan di antrean dan bisa dijalankan saat internet kembali.');
    return false;
  }
  const agent = getAgent();
  const controller = new AbortController();
  state.abort = controller;
  agent.opts.signal = controller.signal;
  const bubble = addMsg('bot', '');
  bubble.classList.add('streaming');
  const textEl = document.createElement('div');
  const toolsEl = document.createElement('div');
  toolsEl.className = 'tools';
  bubble.appendChild(toolsEl);
  bubble.appendChild(textEl);
  let streamed = '';
  agent.opts.onText = (d) => {
    streamed += d;
    textEl.innerHTML = md(streamed);
    scrollLog();
  };
  agent.opts.onTool = (name, input, phase, out) => {
    if (phase === 'start') {
      const t = document.createElement('div');
      t.className = 'tool';
      t.textContent = `${TOOL_LABEL[name] || name}…`;
      t.dataset.name = name;
      toolsEl.appendChild(t);
      streamed = streamed ? streamed + '\n' : '';
    } else {
      const t = $$('.tool', toolsEl).filter((x) => x.dataset.name === name && !x.dataset.done).pop();
      if (t) {
        t.dataset.done = '1';
        const errs = out && out.errors && out.errors.length ? ` (⚠️ ${out.errors.length} gagal)` : '';
        t.textContent = `${TOOL_LABEL[name] || name} ${phase === 'error' ? '❌ ' + friendlyOfficeError(out) : '✓' + errs}`;
        if (phase === 'error' || errs) t.classList.add('bad');
      }
    }
    scrollLog();
  };
  setBusy(true);
  try {
    const note = await contextNote();
    const res = await agent.run(text, note);
    textEl.innerHTML = md(res.text || streamed || 'Selesai.');
    return true;
  } catch (e) {
    if (e instanceof OfflineError) {
      state.online = false;
      updateNet();
      if (!e.partial) {
        enqueue(text);
        textEl.innerHTML = md('📥 Tidak bisa terhubung ke Claude. Perintah disimpan di antrean.');
      } else textEl.innerHTML = md(`⚠️ ${e.message}`);
      setTimeout(() => {
        state.online = navigator.onLine;
        updateNet();
      }, 15000);
    } else {
      bubble.classList.add('err');
      textEl.innerHTML = md(`❌ ${friendlyError(e)}`);
    }
    return false;
  } finally {
    bubble.classList.remove('streaming');
    setBusy(false);
    state.abort = null;
  }
}

/* ------------------------------ alur utama ------------------------------ */

function pushHistory(text) {
  const h = load('history', []);
  if (h[h.length - 1] !== text) h.push(text);
  save('history', h.slice(-100));
  state.historyIdx = -1;
}

async function submit(textArg) {
  const text = (textArg !== undefined ? textArg : $('#input').value).trim();
  if (!text || state.busy) return;
  $('#input').value = '';
  autoGrow();
  renderSuggestions('');
  pushHistory(text);
  addMsg('user', text);
  await route(text);
}

async function route(text) {
  const s = getSettings();
  const chain = parseChain(text, state.host || 'excel');
  const canAI = !!s.apiKey && state.host;
  if (s.mode === 'offline' || (!canAI && chain)) {
    if (chain) return runOffline(chain);
    addMsg('err', 'Perintah tidak dikenali mesin offline.' + (s.apiKey ? ' Ubah mode ke Otomatis/AI untuk memakai Claude.' : ' Isi API key agar perintah bebas bisa dikerjakan AI.'));
    return showClosest(text);
  }
  if (s.mode === 'ai') return runAI(text);
  // Otomatis: perintah yang dikenali & sederhana -> offline (cepat, gratis); selain itu -> AI
  if (chain && !prefersAI(text)) return runOffline(chain);
  if (canAI) return runAI(text);
  if (chain) return runOffline(chain);
  addMsg('err', 'Perintah tidak dikenali. Isi API key agar Claude bisa mengerjakan perintah bebas.');
  return showClosest(text);
}

function examples() {
  if (state.host === 'word') return HELP.word.flatMap(([, ex]) => ex);
  if (state.host === 'excel') return HELP.excel.flatMap(([, ex]) => ex);
  return [...HELP.excel, ...HELP.word].flatMap(([, ex]) => ex);
}

function closest(text, n = 5) {
  const words = canonical(text).split(' ').filter((w) => w.length > 2);
  if (!words.length) return [];
  return examples()
    .map((e) => ({ e, score: words.filter((w) => canonical(e).includes(w)).length }))
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, n)
    .map((x) => x.e);
}

function showClosest(text) {
  const c = closest(text);
  if (!c.length) return addMsg('sys', 'Ketik **bantuan** untuk melihat semua perintah.');
  const el = addMsg('sys', 'Mungkin maksud Anda:');
  const box = document.createElement('div');
  box.className = 'chips';
  c.forEach((e) => box.appendChild(chip(e)));
  el.appendChild(box);
}

function chip(text) {
  const b = document.createElement('button');
  b.className = 'chip';
  b.type = 'button';
  b.textContent = text;
  b.onclick = () => {
    $('#input').value = text.replace(/\s*\(rantai perintah\)$/, '');
    $('#input').focus();
    autoGrow();
  };
  return b;
}

function showHelp() {
  const groups = state.host === 'word' ? HELP.word : state.host === 'excel' ? HELP.excel : [...HELP.excel, ...HELP.word];
  const el = addMsg('bot', '**Perintah offline** (jalan tanpa internet). Klik contoh untuk memakainya:');
  for (const [title, ex] of groups) {
    const h = document.createElement('div');
    h.className = 'help-title';
    h.textContent = title;
    const box = document.createElement('div');
    box.className = 'chips';
    ex.forEach((e) => box.appendChild(chip(e)));
    el.appendChild(h);
    el.appendChild(box);
  }
  const ai = document.createElement('div');
  ai.className = 'help-ai';
  ai.innerHTML = md(`**Dengan AI (online):**\n${HELP.ai.join('\n')}\n\nGabungkan perintah dengan "lalu" atau ";" — misalnya: buat tabel kolom Nama, Gaji lalu format rupiah`);
  el.appendChild(ai);
  scrollLog();
}

/* ------------------------------ saran ketik ------------------------------ */

function renderSuggestions(text) {
  const box = $('#suggestions');
  box.innerHTML = '';
  const list = text && text.length >= 2 ? closest(text, 4) : examples().filter((_, i) => i % 7 === 0).slice(0, 5);
  list.forEach((e) => box.appendChild(chip(e)));
}

function autoGrow() {
  const t = $('#input');
  t.style.height = 'auto';
  t.style.height = Math.min(t.scrollHeight, 160) + 'px';
}

/* ------------------------------ template ------------------------------ */

const EXCEL_FIELDS = {
  cicilan: [
    ['pokok', 'Jumlah pinjaman (Rp)', '50.000.000'],
    ['bunga', 'Bunga per tahun (%)', '10'],
    ['tenor', 'Tenor (bulan)', '12'],
    ['metode', 'Metode', 'anuitas', ['anuitas', 'flat']],
  ],
  absensi: [
    ['bulan', 'Bulan', String(new Date().getMonth()), BULAN.map((b, i) => [String(i), b])],
    ['tahun', 'Tahun', String(new Date().getFullYear())],
    ['baris', 'Jumlah orang', '15'],
  ],
  tabungan: [
    ['target', 'Target (Rp)', '12.000.000'],
    ['tenor', 'Lama (bulan)', '12'],
  ],
  anggaran: [['pemasukan', 'Total pemasukan (Rp)', '10.000.000']],
  'arus-kas': [['saldo', 'Saldo awal (Rp)', '0'], ['baris', 'Jumlah baris', '30']],
  'kas-kecil': [['saldo', 'Dana kas kecil (Rp)', '1.000.000'], ['baris', 'Jumlah baris', '30']],
  gaji: [['baris', 'Jumlah karyawan', '10']],
  faktur: [['kepada', 'Kepada (pelanggan)', ''], ['baris', 'Jumlah baris item', '10']],
  stok: [['baris', 'Jumlah barang', '20']],
  nilai: [['kelas', 'Kelas', ''], ['kkm', 'KKM', '75'], ['baris', 'Jumlah siswa', '30']],
  keuangan: [['baris', 'Jumlah baris', '50']],
  hutang: [['baris', 'Jumlah baris', '30']],
  penjualan: [['baris', 'Jumlah baris', '50']],
  tugas: [['baris', 'Jumlah tugas', '20']],
};

function renderTemplates() {
  const list = $('#templateList');
  list.innerHTML = '';
  const tpls = state.host === 'word' ? WORD_TEMPLATES : state.host === 'excel' ? EXCEL_TEMPLATES : [];
  if (!tpls.length) {
    list.innerHTML = '<p class="muted">Buka dari Word atau Excel untuk melihat template.</p>';
    return;
  }
  for (const t of tpls) {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'card';
    card.innerHTML = `<span class="ic">${t.icon}</span><span><b>${escapeHtml(t.title)}</b><small>${escapeHtml(t.desc || (t.fields ? t.fields.length + ' isian' : ''))}</small></span>`;
    card.onclick = () => openTemplate(t);
    list.appendChild(card);
  }
}

function openTemplate(t) {
  const holder = $('#templateForm');
  holder.innerHTML = '';
  holder.hidden = false;
  if (state.host === 'word') showWordTemplateForm(t, templateDefaults(t, getProfile(), {}), holder);
  else showExcelTemplateForm(t, holder);
  holder.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function field(key, label, value, options, type) {
  const id = `f_${key}_${Math.random().toString(36).slice(2, 7)}`;
  let input;
  if (options) {
    input = `<select id="${id}" data-key="${key}">${options
      .map((o) => {
        const [v, l] = Array.isArray(o) ? o : [o, o];
        return `<option value="${escapeHtml(v)}" ${String(v) === String(value) ? 'selected' : ''}>${escapeHtml(l)}</option>`;
      })
      .join('')}</select>`;
  } else if (type === 'textarea') {
    input = `<textarea id="${id}" data-key="${key}" rows="3">${escapeHtml(value)}</textarea>`;
  } else {
    input = `<input id="${id}" data-key="${key}" value="${escapeHtml(value)}">`;
  }
  return `<label for="${id}">${escapeHtml(label)}</label>${input}`;
}

function formValues(form) {
  const v = {};
  $$('[data-key]', form).forEach((el) => (v[el.dataset.key] = el.value));
  return v;
}

function showWordTemplateForm(tpl, values, container) {
  const form = document.createElement('form');
  form.className = 'tform';
  form.innerHTML = `<h3>${tpl.icon} ${escapeHtml(tpl.title)}</h3>${tpl.fields.map((f) => field(f.key, f.label, values[f.key] || '', null, f.type)).join('')}<div class="row"><button type="submit" class="primary">Sisipkan ke dokumen</button><button type="button" class="ghost" data-act="cancel">Batal</button></div>`;
  form.onsubmit = async (ev) => {
    ev.preventDefault();
    const v = formValues(form);
    try {
      reportResult(await W.runWordAction('word.template', { id: tpl.id, values: v }, env()));
      form.remove();
      if (container.id === 'templateForm') {
        container.hidden = true;
        switchTab('chat');
      }
    } catch (e) {
      addMsg('err', `❌ ${friendlyOfficeError(e)}`);
      if (container.id === 'templateForm') switchTab('chat');
    }
  };
  $('[data-act=cancel]', form).onclick = () => {
    form.remove();
    if (container.id === 'templateForm') container.hidden = true;
  };
  if (container.id === 'log') {
    const wrap = addMsg('bot', `Lengkapi isian **${tpl.title}** lalu klik Sisipkan (nilai awal diambil dari Profil):`);
    wrap.appendChild(form);
    scrollLog();
  } else container.appendChild(form);
}

function showExcelTemplateForm(tpl, container) {
  const fields = EXCEL_FIELDS[tpl.id] || [];
  const form = document.createElement('form');
  form.className = 'tform';
  form.innerHTML = `<h3>${tpl.icon} ${escapeHtml(tpl.title)}</h3><p class="muted">${escapeHtml(tpl.desc)}</p>${fields.map(([k, l, d, opts]) => field(k, l, d, opts)).join('')}<div class="row"><button type="submit" class="primary">Buat di sheet baru</button><button type="button" class="ghost" data-act="cancel">Batal</button></div>`;
  form.onsubmit = async (ev) => {
    ev.preventDefault();
    const raw = formValues(form);
    const params = { id: tpl.id };
    for (const [k, v] of Object.entries(raw)) {
      if (v === '') continue;
      if (['pokok', 'target', 'pemasukan', 'saldo'].includes(k)) params[k] = parseNumber(v);
      else if (k === 'bunga') params[k] = (parseNumber(v) || 0) / 100;
      else if (['tenor', 'baris', 'tahun', 'kkm', 'bulan'].includes(k)) params[k] = parseInt(v, 10);
      else params[k] = v;
    }
    switchTab('chat');
    addMsg('user', `Template: ${tpl.title}`);
    try {
      reportResult(await X.runExcelAction('excel.template', params, env()));
    } catch (e) {
      addMsg('err', `❌ ${friendlyOfficeError(e)}`);
    }
    container.hidden = true;
    form.remove();
  };
  $('[data-act=cancel]', form).onclick = () => {
    container.hidden = true;
    form.remove();
  };
  container.appendChild(form);
}

/* ------------------------------ makro ------------------------------ */

function getMacros() {
  return load('macros', {});
}

async function runMacro(name) {
  const macros = getMacros();
  const key = Object.keys(macros).find((k) => k.toLowerCase() === String(name).toLowerCase());
  if (!key) return addMsg('err', `Makro "${name}" tidak ditemukan.`);
  addMsg('sys', `▶ Menjalankan makro **${key}** (${macros[key].length} perintah)`);
  for (const cmd of macros[key]) {
    addMsg('user', cmd);
    const ok = await route(cmd);
    if (ok === false) break;
  }
}

function renderMacros() {
  const macros = getMacros();
  const list = $('#macroList');
  list.innerHTML = '';
  const names = Object.keys(macros);
  if (!names.length) list.innerHTML = '<p class="muted">Belum ada makro. Makro = kumpulan perintah yang dijalankan sekaligus.</p>';
  for (const n of names) {
    const row = document.createElement('div');
    row.className = 'macro';
    row.innerHTML = `<div><b>${escapeHtml(n)}</b><small>${macros[n].map(escapeHtml).join(' → ')}</small></div>`;
    const run = document.createElement('button');
    run.className = 'primary small';
    run.textContent = 'Jalankan';
    run.onclick = () => {
      switchTab('chat');
      runMacro(n);
    };
    const del = document.createElement('button');
    del.className = 'ghost small';
    del.textContent = 'Hapus';
    del.onclick = () => {
      const m = getMacros();
      delete m[n];
      save('macros', m);
      renderMacros();
    };
    const btns = document.createElement('div');
    btns.className = 'row';
    btns.append(run, del);
    row.appendChild(btns);
    list.appendChild(row);
  }
}

/* ------------------------------ pengaturan ------------------------------ */

const PROFILE_FIELDS = [
  ['nama', 'Nama lengkap'],
  ['jabatan', 'Jabatan'],
  ['instansi', 'Instansi / perusahaan / sekolah'],
  ['alamat', 'Alamat instansi'],
  ['kota', 'Kota'],
  ['telepon', 'Telepon'],
  ['email', 'Email'],
];

function renderSettings() {
  const s = getSettings();
  $('#apiKey').value = s.apiKey;
  $('#model').innerHTML = MODELS.map((m) => `<option value="${m.id}" ${m.id === s.model ? 'selected' : ''}>${escapeHtml(m.label)}</option>`).join('');
  $('#effort').value = s.effort;
  $('#codePolicy').value = s.codePolicy;
  $('#modeSelect').value = s.mode;
  const p = getProfile();
  $('#profileFields').innerHTML = PROFILE_FIELDS.map(([k, l]) => field(k, l, p[k] || '')).join('');
}

function saveSettings() {
  const s = getSettings();
  s.apiKey = $('#apiKey').value.trim();
  s.model = $('#model').value;
  s.effort = $('#effort').value;
  s.codePolicy = $('#codePolicy').value;
  save('settings', s);
  save('profile', formValues($('#profileFields')));
  updateNet();
  toast('Pengaturan disimpan.');
}

function toast(text) {
  const t = $('#toast');
  t.textContent = text;
  t.hidden = false;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => (t.hidden = true), 2500);
}

/* ------------------------------ tab & init ------------------------------ */

function switchTab(name) {
  $$('.tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === name));
  $$('.panel').forEach((p) => (p.hidden = p.id !== `panel-${name}`));
  if (name === 'templates') renderTemplates();
  if (name === 'macros') renderMacros();
  if (name === 'settings') renderSettings();
  if (name === 'chat') $('#input').focus();
}

function bind() {
  $$('.tab').forEach((t) => (t.onclick = () => switchTab(t.dataset.tab)));
  $('#sendBtn').onclick = () => submit();
  $('#stopBtn').onclick = () => state.abort && state.abort.abort();
  $('#input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      submit();
    } else if ((e.key === 'ArrowUp' || e.key === 'ArrowDown') && !$('#input').value.includes('\n')) {
      const h = load('history', []);
      if (!h.length) return;
      if (e.key === 'ArrowUp') state.historyIdx = state.historyIdx < 0 ? h.length - 1 : Math.max(0, state.historyIdx - 1);
      else state.historyIdx = state.historyIdx < 0 ? -1 : state.historyIdx + 1;
      if (state.historyIdx >= h.length) state.historyIdx = -1;
      $('#input').value = state.historyIdx < 0 ? '' : h[state.historyIdx];
      e.preventDefault();
    }
  });
  $('#input').addEventListener('input', () => {
    autoGrow();
    renderSuggestions($('#input').value);
  });
  $('#newChat').onclick = () => {
    if (state.agent) state.agent.reset();
    $('#log').innerHTML = '';
    welcome();
  };
  $('#helpBtn').onclick = showHelp;
  $('#runQueue').onclick = () => runQueue();
  $('#clearQueue').onclick = () => {
    save('queue', load('queue', []).filter((q) => q.host !== state.host));
    renderQueue();
  };
  $('#modeSelect').onchange = () => {
    const s = getSettings();
    s.mode = $('#modeSelect').value;
    save('settings', s);
    updateNet();
    toast({ auto: 'Mode Otomatis: perintah sederhana offline, sisanya AI.', ai: 'Mode AI: semua perintah dikerjakan Claude.', offline: 'Mode Offline: hanya mesin perintah lokal.' }[s.mode]);
  };
  $('#saveSettings').onclick = saveSettings;
  $('#toggleKey').onclick = () => {
    const i = $('#apiKey');
    i.type = i.type === 'password' ? 'text' : 'password';
  };
  $('#testKey').onclick = async () => {
    const key = $('#apiKey').value.trim();
    if (!key) return toast('Isi API key dulu.');
    $('#testKey').disabled = true;
    $('#keyStatus').textContent = 'Menguji…';
    try {
      const r = await testApiKey(key, $('#model').value);
      $('#keyStatus').textContent = `✅ Terhubung (${r.trim().slice(0, 20)})`;
    } catch (e) {
      $('#keyStatus').textContent = `❌ ${friendlyError(e)}`;
    } finally {
      $('#testKey').disabled = false;
    }
  };
  $('#saveMacro').onclick = () => {
    const name = $('#macroName').value.trim();
    const cmds = $('#macroCommands').value.split('\n').map((s) => s.trim()).filter(Boolean);
    if (!name || !cmds.length) return toast('Isi nama dan minimal satu perintah.');
    const m = getMacros();
    m[name] = cmds;
    save('macros', m);
    $('#macroName').value = '';
    $('#macroCommands').value = '';
    renderMacros();
    toast(`Makro "${name}" disimpan. Jalankan dengan: jalankan makro ${name}`);
  };
  $('#fromHistory').onclick = () => {
    $('#macroCommands').value = load('history', []).slice(-5).join('\n');
  };
  $('#clearHistory').onclick = () => {
    save('history', []);
    toast('Riwayat perintah dihapus.');
  };
}

function welcome() {
  const where = state.host === 'excel' ? 'Excel' : state.host === 'word' ? 'Word' : 'pratinjau browser';
  addMsg(
    'bot',
    `Halo! Saya **Master** untuk ${where}. Ketik perintah biasa, contoh:\n` +
      (state.host === 'word'
        ? '- buat surat izin sakit\n- buat notulen rapat\n- font Times New Roman ukuran 12 semua'
        : '- simulasi cicilan pinjaman 50 juta bunga 10% 24 bulan\n- buat tabel karyawan dengan kolom Nama, Jabatan, Gaji\n- jumlahkan / urutkan / grafik') +
      '\n\nKetik **bantuan** untuk semua perintah.',
  );
  if (!getSettings().apiKey) addMsg('sys', '💡 Tanpa API key, Master tetap jalan offline dengan perintah bawaan. Isi API key di Pengaturan agar Claude bisa mengerjakan perintah bebas.');
}

function start(host) {
  state.host = host;
  document.body.dataset.host = host || 'preview';
  $('#hostBadge').textContent = host === 'excel' ? 'Excel' : host === 'word' ? 'Word' : 'Pratinjau';
  bind();
  renderSettings();
  updateNet();
  renderSuggestions('');
  welcome();
  switchTab('chat');
}

function boot() {
  if (typeof Office === 'undefined' || !Office.onReady) {
    start(null);
    return;
  }
  Office.onReady((info) => {
    const h = info && info.host;
    const host = h === Office.HostType.Excel ? 'excel' : h === Office.HostType.Word ? 'word' : null;
    const forced = new URLSearchParams(location.search).get('host');
    start(host || (forced === 'excel' || forced === 'word' ? forced : null));
  });
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
