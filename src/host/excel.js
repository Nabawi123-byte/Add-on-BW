/* global Excel, Office */
// Executor Excel: menjalankan operasi rencana (plan), perintah offline, dan alat AI lewat Office.js.

import {
  cellAddress,
  columnLetter,
  parseAddress,
  normalize,
  titleCase,
  looksLikeMoneyColumn,
  looksLikeDateColumn,
  FORMAT_RUPIAH,
  FORMAT_ANGKA,
  FORMAT_PERSEN,
  FORMAT_TANGGAL,
  clamp,
} from '../core/util.js';
import { terbilang, terbilangRupiah } from '../core/terbilang.js';
import { EXCEL_TEMPLATES } from '../core/excel-templates.js';

const CHAR_WIDTH = 7; // perkiraan poin per karakter untuk lebar kolom
const LIST_SHEET = '_MasterList';
const MAX_WRITE_CELLS = 200000;

/* ------------------------------ util ------------------------------ */

function splitSheet(address) {
  const a = String(address || '');
  const bang = a.lastIndexOf('!');
  if (bang < 0) return { sheet: null, addr: a };
  return { sheet: a.slice(0, bang).replace(/^'|'$/g, ''), addr: a.slice(bang + 1) };
}

function rangeOf(context, sheet, address) {
  const { sheet: s, addr } = splitSheet(address);
  const ws = s ? context.workbook.worksheets.getItem(s) : sheet;
  return ws.getRange(addr);
}

function fill2D(rows, cols, v) {
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => v));
}

function cleanSheetName(name) {
  return String(name || 'Sheet').replace(/[[\]:*?/\\]/g, ' ').replace(/^'+|'+$/g, '').trim().slice(0, 31) || 'Sheet';
}

async function uniqueSheetName(context, wanted) {
  const sheets = context.workbook.worksheets;
  sheets.load('items/name');
  await context.sync();
  const names = new Set(sheets.items.map((s) => s.name.toLowerCase()));
  const base = cleanSheetName(wanted);
  let name = base;
  let i = 2;
  while (names.has(name.toLowerCase())) name = `${base.slice(0, 26)} (${i++})`;
  return name;
}

async function uniqueTableName(context, wanted) {
  const tables = context.workbook.tables;
  tables.load('items/name');
  const names = context.workbook.names;
  names.load('items/name');
  await context.sync();
  const used = new Set([...tables.items.map((t) => t.name.toLowerCase()), ...names.items.map((n) => n.name.toLowerCase())]);
  let base = String(wanted || 'Tabel').replace(/[^A-Za-z0-9_]/g, '');
  if (!/^[A-Za-z_]/.test(base)) base = 'T' + base;
  base = base.slice(0, 200) || 'Tabel';
  let name = base;
  let i = 2;
  while (used.has(name.toLowerCase())) name = `${base}${i++}`;
  return name;
}

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const isEmpty = (v) => v === '' || v === null || v === undefined;

/* ---------------------- operasi rencana (ops) ---------------------- */

const BORDER_EDGES = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];

async function opWrite(context, sheet, op) {
  const values = (op.values || []).map((r) => (Array.isArray(r) ? r : [r]));
  if (!values.length) return;
  const cols = Math.max(...values.map((r) => r.length));
  if (values.length * cols > MAX_WRITE_CELLS) throw new Error('data terlalu besar');
  const norm = values.map((r) => Array.from({ length: cols }, (_, j) => (r[j] === undefined || r[j] === null ? '' : r[j])));
  const start = rangeOf(context, sheet, op.address).getCell(0, 0);
  const target = start.getResizedRange(norm.length - 1, cols - 1);
  target.formulas = norm;
  if (op.numberFormat) target.numberFormat = fill2D(norm.length, cols, op.numberFormat);
}

async function opFormat(context, sheet, op) {
  const range = rangeOf(context, sheet, op.address);
  const fmt = range.format;
  if (op.bold !== undefined) fmt.font.bold = !!op.bold;
  if (op.italic !== undefined) fmt.font.italic = !!op.italic;
  if (op.underline !== undefined) fmt.font.underline = op.underline ? 'Single' : 'None';
  if (op.fontSize) fmt.font.size = Number(op.fontSize);
  if (op.fontName) fmt.font.name = op.fontName;
  if (op.fontColor) fmt.font.color = op.fontColor;
  if (op.fill) fmt.fill.color = op.fill;
  if (op.fill === null || op.fill === 'none') fmt.fill.clear();
  if (op.hAlign) fmt.horizontalAlignment = op.hAlign;
  if (op.vAlign) fmt.verticalAlignment = op.vAlign;
  if (op.wrap !== undefined) fmt.wrapText = !!op.wrap;
  if (op.columnWidth) fmt.columnWidth = Number(op.columnWidth) * CHAR_WIDTH;
  if (op.rowHeight) fmt.rowHeight = Number(op.rowHeight);
  if (op.numberFormat || op.border) {
    range.load('rowCount,columnCount');
    await context.sync();
  }
  if (op.numberFormat) range.numberFormat = fill2D(range.rowCount, range.columnCount, op.numberFormat);
  if (op.border) {
    const set = (edge, style = 'Continuous') => {
      const b = fmt.borders.getItem(edge);
      b.style = style;
      if (style !== 'None') b.color = '#000000';
    };
    if (op.border === 'all' || op.border === 'outline') BORDER_EDGES.forEach((e) => set(e));
    if (op.border === 'all') {
      if (range.rowCount > 1) set('InsideHorizontal');
      if (range.columnCount > 1) set('InsideVertical');
    }
    if (op.border === 'top') set('EdgeTop');
    if (op.border === 'top-double') set('EdgeTop', 'Double');
    if (op.border === 'bottom') set('EdgeBottom');
    if (op.border === 'none') {
      BORDER_EDGES.forEach((e) => set(e, 'None'));
      if (range.rowCount > 1) set('InsideHorizontal', 'None');
      if (range.columnCount > 1) set('InsideVertical', 'None');
    }
  }
  if (op.merge) range.merge(false);
}

async function opTable(context, sheet, op) {
  const { sheet: s, addr } = splitSheet(op.address);
  const ws = s ? context.workbook.worksheets.getItem(s) : sheet;
  const table = ws.tables.add(addr, op.headers !== false);
  table.name = await uniqueTableName(context, op.name || 'Tabel');
  table.style = op.style || 'TableStyleMedium2';
  if (op.showTotals) table.showTotals = true;
  return `tabel ${table.name}`;
}

async function opChart(context, sheet, op) {
  const range = rangeOf(context, sheet, op.address);
  const { sheet: s } = splitSheet(op.address);
  const ws = s ? context.workbook.worksheets.getItem(s) : sheet;
  const chart = ws.charts.add(op.type || 'ColumnClustered', range, 'Auto');
  if (op.title) {
    chart.title.text = op.title;
    chart.title.visible = true;
  }
  if (op.position) {
    chart.setPosition(op.position);
  } else {
    range.load('rowIndex,columnIndex,columnCount');
    await context.sync();
    chart.setPosition(cellAddress(range.rowIndex + 1, range.columnIndex + range.columnCount + 2));
  }
  chart.width = Number(op.width) || 460;
  chart.height = Number(op.height) || 280;
  if (/Pie|Doughnut/.test(op.type || '')) {
    chart.dataLabels.showPercentage = true;
    chart.dataLabels.showValue = false;
  }
}

function cfValue(v) {
  if (typeof v === 'number') return String(v);
  const s = String(v === undefined ? '' : v);
  if (/^[=\-\d.]/.test(s) && !isNaN(Number(s))) return s;
  if (s.startsWith('=') || s.startsWith('"')) return s;
  return `"${s.replace(/"/g, '""')}"`;
}

async function opConditional(context, sheet, op) {
  const range = rangeOf(context, sheet, op.address);
  const CF = range.conditionalFormats;
  const style = (f) => {
    if (op.fill) f.fill.color = op.fill;
    if (op.fontColor) f.font.color = op.fontColor;
    if (op.bold) f.font.bold = true;
  };
  const ops = { greater: 'GreaterThan', less: 'LessThan', equal: 'EqualTo', between: 'Between', notEqual: 'NotEqualTo', greaterOrEqual: 'GreaterThanOrEqual', lessOrEqual: 'LessThanOrEqual' };
  if (ops[op.rule]) {
    const cf = CF.add('CellValue');
    const rule = { formula1: cfValue(op.value), operator: ops[op.rule] };
    if (op.rule === 'between') rule.formula2 = cfValue(op.value2);
    cf.cellValue.rule = rule;
    style(cf.cellValue.format);
  } else if (op.rule === 'text') {
    const cf = CF.add('ContainsText');
    cf.textComparison.rule = { operator: 'Contains', text: String(op.value) };
    style(cf.textComparison.format);
  } else if (op.rule === 'duplicate' || op.rule === 'unique') {
    const cf = CF.add('PresetCriteria');
    cf.preset.rule = { criterion: op.rule === 'duplicate' ? 'DuplicateValues' : 'UniqueValues' };
    style(cf.preset.format);
  } else if (op.rule === 'top' || op.rule === 'bottom') {
    const cf = CF.add('TopBottom');
    cf.topBottom.rule = { rank: Number(op.value) || 10, type: op.rule === 'top' ? 'TopItems' : 'BottomItems' };
    style(cf.topBottom.format);
  } else if (op.rule === 'dataBar') {
    CF.add('DataBar');
  } else if (op.rule === 'colorScale') {
    const cf = CF.add('ColorScale');
    cf.colorScale.criteria = {
      minimum: { formula: null, type: 'LowestValue', color: '#F8696B' },
      midpoint: { formula: '50', type: 'Percentile', color: '#FFEB84' },
      maximum: { formula: null, type: 'HighestValue', color: '#63BE7B' },
    };
  } else if (op.rule === 'formula') {
    const cf = CF.add('Custom');
    cf.custom.rule.formula = String(op.value).startsWith('=') ? op.value : '=' + op.value;
    style(cf.custom.format);
  } else {
    throw new Error(`aturan "${op.rule}" tidak dikenal`);
  }
}

/** Daftar pilihan (dropdown) disimpan di sheet tersembunyi agar aman di semua bahasa/locale Excel. */
async function opValidation(context, sheet, op) {
  const items = (op.list || []).map(String).filter(Boolean);
  if (!items.length) throw new Error('daftar pilihan kosong');
  let ls = context.workbook.worksheets.getItemOrNullObject(LIST_SHEET);
  await context.sync();
  if (ls.isNullObject) {
    ls = context.workbook.worksheets.add(LIST_SHEET);
    ls.visibility = 'Hidden';
  }
  const used = ls.getUsedRangeOrNullObject(true);
  used.load('columnCount,columnIndex');
  await context.sync();
  const col = used.isNullObject ? 1 : used.columnIndex + used.columnCount + 1;
  const letter = columnLetter(col);
  ls.getRange(`${letter}1:${letter}${items.length}`).values = items.map((v) => [v]);
  const range = rangeOf(context, sheet, op.address);
  range.dataValidation.clear();
  range.dataValidation.rule = { list: { inCellDropDown: true, source: `='${LIST_SHEET}'!$${letter}$1:$${letter}$${items.length}` } };
}

async function opFreeze(context, sheet, op) {
  const rows = Number(op.rows) || 0;
  const cols = Number(op.cols) || 0;
  sheet.freezePanes.unfreeze();
  if (rows && cols) sheet.freezePanes.freezeAt(sheet.getRange(`A1:${columnLetter(cols)}${rows}`));
  else if (rows) sheet.freezePanes.freezeRows(rows);
  else if (cols) sheet.freezePanes.freezeColumns(cols);
}

async function opAutofit(context, sheet, op) {
  const range = op.address ? rangeOf(context, sheet, op.address) : sheet.getUsedRange();
  range.format.autofitColumns();
  await context.sync();
  // Batasi kolom yang terlalu lebar (mis. karena judul panjang)
  const cols = range.columns;
  cols.load('items/format/columnWidth');
  await context.sync();
  for (const c of cols.items) if (c.format.columnWidth > 60 * CHAR_WIDTH) c.format.columnWidth = 45 * CHAR_WIDTH;
}

async function opTerbilang(context, sheet, op) {
  const src = rangeOf(context, sheet, op.source);
  src.load('values');
  await context.sync();
  const v = Number(src.values[0][0]) || 0;
  rangeOf(context, sheet, op.target).values = [[op.rupiah === false ? terbilang(v) : terbilangRupiah(v)]];
  sheet.load('name');
  await context.sync();
  rememberTerbilang({ sheet: sheet.name, source: op.source, target: op.target, rupiah: op.rupiah !== false });
}

async function opSort(context, sheet, op) {
  const range = rangeOf(context, sheet, op.address);
  let key = op.column;
  if (typeof key !== 'number') {
    range.load('values');
    await context.sync();
    const headers = range.values[0].map((h) => normalize(h));
    const idx = headers.findIndex((h) => h === normalize(key)) >= 0 ? headers.findIndex((h) => h === normalize(key)) : headers.findIndex((h) => key && h.includes(normalize(key)));
    key = idx >= 0 ? idx : 0;
  }
  range.sort.apply([{ key, ascending: op.ascending !== false }], false, op.hasHeaders !== false);
}

const OP_HANDLERS = {
  write: opWrite,
  format: opFormat,
  numberFormat: (c, s, op) => opFormat(c, s, { address: op.address, numberFormat: op.format || op.numberFormat }),
  table: opTable,
  chart: opChart,
  conditional: opConditional,
  validation: opValidation,
  freeze: opFreeze,
  unfreeze: async (c, s) => s.freezePanes.unfreeze(),
  autofit: opAutofit,
  terbilang: opTerbilang,
  sort: opSort,
  select: async (c, s, op) => rangeOf(c, s, op.address).select(),
  clear: async (c, s, op) => rangeOf(c, s, op.address).clear(op.what || 'Contents'),
  filter: async (c, s, op) => s.autoFilter.apply(rangeOf(c, s, op.address)),
  removeDuplicates: async (c, s, op) => {
    const range = rangeOf(c, s, op.address);
    range.load('columnCount');
    await c.sync();
    const cols = op.columns || Array.from({ length: range.columnCount }, (_, i) => i);
    const res = range.removeDuplicates(cols, op.hasHeaders !== false);
    res.load('removed');
    await c.sync();
    return `${res.removed} baris duplikat dihapus`;
  },
  insertRows: async (c, s, op) => rangeOf(c, s, op.address).insert('Down'),
  deleteRows: async (c, s, op) => rangeOf(c, s, op.address).delete('Up'),
  insertColumns: async (c, s, op) => rangeOf(c, s, op.address).insert('Right'),
  deleteColumns: async (c, s, op) => rangeOf(c, s, op.address).delete('Left'),
  addSheet: async (c, s, op) => {
    const name = await uniqueSheetName(c, op.name || 'Sheet Baru');
    const ws = c.workbook.worksheets.add(name);
    if (op.activate !== false) ws.activate();
    return `sheet "${name}"`;
  },
  activate: async (c, s, op) => c.workbook.worksheets.getItem(op.sheet).activate(),
  renameSheet: async (c, s, op) => {
    const ws = op.sheet ? c.workbook.worksheets.getItem(op.sheet) : s;
    ws.name = cleanSheetName(op.name);
  },
  deleteSheet: async (c, s, op) => c.workbook.worksheets.getItem(op.sheet).delete(),
  name: async (c, s, op) => c.workbook.names.add(op.name, rangeOf(c, s, op.address)),
};

export const EXCEL_OPS = Object.keys(OP_HANDLERS);

/** Menjalankan daftar operasi pada sheet tertentu; error per operasi dikumpulkan, tidak menghentikan semuanya. */
async function runOps(context, sheet, ops) {
  const done = [];
  const errors = [];
  for (const op of ops || []) {
    const h = OP_HANDLERS[op && op.op];
    if (!h) {
      errors.push(`operasi "${op && op.op}" tidak dikenal`);
      continue;
    }
    try {
      const ws = op.sheet ? context.workbook.worksheets.getItem(op.sheet) : sheet;
      const info = await h(context, ws, op);
      await context.sync();
      done.push(info || op.op);
    } catch (e) {
      errors.push(`${op.op}${op.address ? ' ' + op.address : ''}: ${(e && e.message) || e}`);
    }
  }
  return { done, errors };
}

/** Menjalankan SheetPlan (dari template, perintah offline, atau AI). */
export async function applyPlan(plan) {
  return Excel.run(async (context) => {
    let sheet;
    let sheetName;
    if (plan.newSheet) {
      sheetName = await uniqueSheetName(context, plan.sheetName || 'Master');
      sheet = context.workbook.worksheets.add(sheetName);
      sheet.activate();
      await context.sync();
    } else if (plan.sheetName) {
      sheet = context.workbook.worksheets.getItemOrNullObject(plan.sheetName);
      await context.sync();
      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add(cleanSheetName(plan.sheetName));
        await context.sync();
      }
      sheetName = plan.sheetName;
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      await context.sync();
      sheetName = sheet.name;
    }
    const res = await runOps(context, sheet, plan.ops);
    return Object.assign({ sheetName }, res);
  });
}

/* ------------------- penyimpanan pengaturan dokumen ------------------- */

function rememberTerbilang(entry) {
  try {
    const s = Office.context.document.settings;
    const list = (s.get('master.terbilang') || []).filter((e) => !(e.sheet === entry.sheet && e.target === entry.target));
    list.push(entry);
    s.set('master.terbilang', list.slice(-200));
    s.saveAsync(() => {});
  } catch (e) {
    /* pengaturan dokumen tidak tersedia */
  }
}

/* ------------------------- konteks seleksi ------------------------- */

/**
 * Area kerja untuk perintah offline: seleksi bila lebih dari satu sel,
 * selain itu tabel tempat sel aktif berada, atau area terpakai di sheet.
 */
async function workArea(context) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const sel = context.workbook.getSelectedRange();
  sel.load('address,rowCount,columnCount,rowIndex,columnIndex');
  sheet.load('name');
  await context.sync();
  let range = sel;
  let source = 'seleksi';
  if (sel.rowCount * sel.columnCount === 1) {
    let tables = null;
    try {
      tables = sel.getTables(false); // ExcelApi 1.9
      tables.load('items/name');
      await context.sync();
    } catch (e) {
      tables = null;
    }
    if (tables && tables.items && tables.items.length) {
      range = tables.items[0].getRange();
      source = `tabel ${tables.items[0].name}`;
    } else {
      const used = sheet.getUsedRangeOrNullObject(true);
      await context.sync();
      if (used.isNullObject) throw new Error('Sheet masih kosong. Isi data atau pilih sel terlebih dahulu.');
      range = used;
      source = 'seluruh data di sheet';
    }
  }
  range.load('address,values,formulas,numberFormat,rowCount,columnCount,rowIndex,columnIndex');
  await context.sync();
  // Abaikan baris ringkasan (TOTAL / RATA-RATA ...) buatan Master di bawah data.
  if (range.rowCount > 2) {
    const lastF = range.formulas[range.rowCount - 1];
    const labels = Object.values(FN_LABEL);
    const isSummary = lastF.some((f) => typeof f === 'string' && labels.includes(f)) && lastF.some((f) => typeof f === 'string' && /^=(SUM|AVERAGE|MAX|MIN|COUNTA)\(/.test(f));
    if (isSummary) {
      range = range.getResizedRange(-1, 0);
      range.load('address,values,formulas,numberFormat,rowCount,columnCount,rowIndex,columnIndex');
      await context.sync();
      source += ' (tanpa baris total)';
    }
  }
  const values = range.values;
  const hasHeader = values.length > 1 && values[0].every((v) => typeof v === 'string' || isEmpty(v)) && values.slice(1).some((r) => r.some(isNum));
  return { sheet, range, source, values, formulas: range.formulas, numberFormat: range.numberFormat, hasHeader, row: range.rowIndex + 1, col: range.columnIndex + 1 };
}

function numericColumns(area) {
  const start = area.hasHeader ? 1 : 0;
  const cols = [];
  for (let j = 0; j < area.values[0].length; j++) {
    const data = area.values.slice(start).map((r) => r[j]);
    const nums = data.filter(isNum).length;
    const texts = data.filter((v) => typeof v === 'string' && v !== '').length;
    if (nums > 0 && nums >= texts) cols.push(j);
  }
  return cols;
}

function findColumn(area, name) {
  if (!name || !area.hasHeader) return -1;
  const headers = area.values[0].map((h) => normalize(h));
  const n = normalize(name);
  let idx = headers.findIndex((h) => h === n);
  if (idx < 0) idx = headers.findIndex((h) => h.includes(n) || (h && n.includes(h)));
  if (idx < 0 && /^[a-z]{1,2}$/.test(n)) {
    const c = parseAddress(n.toUpperCase() + '1');
    if (c) idx = c.col - area.col;
  }
  return idx;
}

const FN_LABEL = { SUM: 'TOTAL', AVERAGE: 'RATA-RATA', MAX: 'MAKSIMUM', MIN: 'MINIMUM', COUNTA: 'JUMLAH DATA' };

/* ------------------------- perintah offline ------------------------- */

async function withArea(fn) {
  return Excel.run(async (context) => {
    const area = await workArea(context);
    const out = await fn(context, area);
    await context.sync();
    return out;
  });
}

const ACTIONS = {
  async template({ id, ...params }, env) {
    const tpl = EXCEL_TEMPLATES.find((t) => t.id === id);
    if (!tpl) throw new Error(`Template ${id} tidak ditemukan`);
    const plan = tpl.build(params, env.profile || {});
    const res = await applyPlan(plan.toJSON());
    return { message: `✅ ${tpl.title} dibuat di sheet "${res.sheetName}". ${plan.summary || ''}`, errors: res.errors };
  },

  async table({ headers, cols, rows = 10, name }) {
    rows = clamp(rows || 10, 1, 10000);
    let hs = headers && headers.length ? headers.map((h) => titleCase(h)) : Array.from({ length: clamp(cols || 3, 1, 200) }, (_, i) => `Kolom ${i + 1}`);
    const addNo = headers && headers.length && !hs.some((h) => /^(no|no\.|nomor)$/i.test(h));
    if (addNo) hs = ['No', ...hs];
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const sel = context.workbook.getSelectedRange();
      sel.load('values,rowIndex,columnIndex,rowCount,columnCount');
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load('rowIndex,rowCount');
      await context.sync();
      let r = 1;
      let c = 1;
      const selEmpty = sel.values.every((row) => row.every(isEmpty));
      if (selEmpty) {
        r = sel.rowIndex + 1;
        c = sel.columnIndex + 1;
      } else if (!used.isNullObject) {
        r = used.rowIndex + used.rowCount + 2;
      }
      const ops = [];
      if (name) {
        ops.push({ op: 'write', address: cellAddress(r, c), values: [[titleCase(name)]] });
        ops.push({ op: 'format', address: cellAddress(r, c), bold: true, fontSize: 14 });
        r += 1;
      }
      const values = [hs];
      for (let i = 0; i < rows; i++) values.push(hs.map((h, j) => (addNo && j === 0 ? i + 1 : '')));
      ops.push({ op: 'write', address: cellAddress(r, c), values });
      hs.forEach((h, j) => {
        const col = columnLetter(c + j);
        const rg = `${col}${r + 1}:${col}${r + rows}`;
        if (looksLikeMoneyColumn(h)) ops.push({ op: 'numberFormat', address: rg, format: FORMAT_RUPIAH });
        else if (looksLikeDateColumn(h)) ops.push({ op: 'numberFormat', address: rg, format: FORMAT_TANGGAL });
      });
      ops.push({ op: 'table', address: `${cellAddress(r, c)}:${cellAddress(r + rows, c + hs.length - 1)}`, name: name || 'Tabel' });
      ops.push({ op: 'autofit', address: `${cellAddress(r, c)}:${cellAddress(r + rows, c + hs.length - 1)}` });
      ops.push({ op: 'select', address: cellAddress(r + 1, c + (addNo ? 1 : 0)) });
      const res = await runOps(context, sheet, ops);
      return { message: `✅ Tabel ${hs.length} kolom × ${rows} baris dibuat di ${cellAddress(r, c)}.`, errors: res.errors };
    });
  },

  async toTable() {
    return withArea(async (context, area) => {
      const t = area.sheet.tables.add(area.range, area.hasHeader || area.values[0].every((v) => typeof v === 'string'));
      t.name = await uniqueTableName(context, 'Tabel');
      t.style = 'TableStyleMedium2';
      return { message: `✅ ${area.range.address} dijadikan tabel.` };
    });
  },

  async aggregate({ fn }) {
    return withArea(async (context, area) => {
      const cols = numericColumns(area);
      if (!cols.length) throw new Error('Tidak ada kolom angka di area ini.');
      const start = area.hasHeader ? area.row + 1 : area.row;
      const end = area.row + area.values.length - 1;
      const outRow = end + 1;
      const row = area.values[0].map(() => '');
      cols.forEach((j) => {
        const L = columnLetter(area.col + j);
        row[j] = `=${fn}(${L}${start}:${L}${end})`;
      });
      const labelCol = row.findIndex((v, j) => !cols.includes(j));
      if (labelCol >= 0) row[labelCol] = FN_LABEL[fn] || fn;
      const out = area.sheet.getRange(`${cellAddress(outRow, area.col)}:${cellAddress(outRow, area.col + row.length - 1)}`);
      out.formulas = [row];
      out.format.font.bold = true;
      out.format.fill.color = '#D9E1F2';
      out.format.borders.getItem('EdgeTop').style = 'Double';
      const fmts = area.numberFormat[area.numberFormat.length - 1];
      out.numberFormat = [fmts.map((f, j) => (cols.includes(j) && fn !== 'COUNTA' ? f : 'General'))];
      return { message: `✅ Baris ${FN_LABEL[fn] || fn} ditambahkan di baris ${outRow} (${area.source}).` };
    });
  },

  async sort({ ascending, column }) {
    return withArea(async (context, area) => {
      let key = findColumn(area, column);
      if (key < 0) key = numericColumns(area)[0] !== undefined && column ? numericColumns(area)[0] : 0;
      const data = area.hasHeader ? area.range.getOffsetRange(1, 0).getResizedRange(-1, 0) : area.range;
      data.sort.apply([{ key, ascending }], false, false);
      const header = area.hasHeader ? area.values[0][key] : columnLetter(area.col + key);
      return { message: `✅ Data diurutkan berdasarkan "${header}" ${ascending ? 'menaik (A→Z / kecil→besar)' : 'menurun (Z→A / besar→kecil)'}.` };
    });
  },

  async filter() {
    return withArea(async (context, area) => {
      if (area.source.startsWith('tabel')) return { message: 'ℹ️ Tabel sudah memiliki tombol filter di baris judul.' };
      area.sheet.autoFilter.apply(area.range);
      return { message: `✅ Filter dipasang pada ${area.range.address}.` };
    });
  },

  async numberFormat({ kind }) {
    const formats = { rupiah: FORMAT_RUPIAH, angka: FORMAT_ANGKA, persen: FORMAT_PERSEN, tanggal: FORMAT_TANGGAL, teks: '@' };
    return Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.load('rowCount,columnCount,address');
      await context.sync();
      if (sel.rowCount * sel.columnCount > 1) {
        sel.numberFormat = fill2D(sel.rowCount, sel.columnCount, formats[kind]);
        sel.format.autofitColumns();
        return { message: `✅ Format ${kind} diterapkan ke ${sel.address}.` };
      }
      const area = await workArea(context);
      const cols = kind === 'teks' ? area.values[0].map((_, j) => j) : numericColumns(area);
      if (!cols.length) throw new Error('Pilih sel yang ingin diformat terlebih dahulu.');
      const start = area.hasHeader ? area.row + 1 : area.row;
      const end = area.row + area.values.length - 1;
      cols.forEach((j) => {
        const L = columnLetter(area.col + j);
        const rg = area.sheet.getRange(`${L}${start}:${L}${end}`);
        rg.numberFormat = fill2D(end - start + 1, 1, formats[kind]);
      });
      area.range.format.autofitColumns();
      return { message: `✅ Format ${kind} diterapkan ke ${cols.length} kolom angka (${area.source}).` };
    });
  },

  async chart({ type, title }) {
    return withArea(async (context, area) => {
      const chart = area.sheet.charts.add(type || 'ColumnClustered', area.range, 'Auto');
      chart.title.text = title || (area.hasHeader ? `Grafik ${area.values[0].filter(Boolean).slice(-1)[0]}` : 'Grafik');
      chart.setPosition(cellAddress(area.row, area.col + area.values[0].length + 1));
      chart.width = 460;
      chart.height = 280;
      return { message: `✅ Grafik dibuat dari ${area.range.address}.` };
    });
  },

  async highlight({ rule, value, fill, fontColor }) {
    return withArea(async (context, area) => {
      let target = area.range;
      if (area.hasHeader) target = area.range.getOffsetRange(1, 0).getResizedRange(-1, 0);
      if (rule !== 'duplicate' && rule !== 'text') {
        const cols = numericColumns(area);
        if (cols.length && cols.length < area.values[0].length && area.source !== 'seleksi') {
          // Hanya kolom angka bila memakai seluruh area
          for (const j of cols) {
            const L = columnLetter(area.col + j);
            const rg = `${L}${area.hasHeader ? area.row + 1 : area.row}:${L}${area.row + area.values.length - 1}`;
            await opConditional(context, area.sheet, { address: rg, rule, value, fill, fontColor });
          }
          return { message: `✅ Sel di kolom angka ditandai (${rule} ${value}).` };
        }
      }
      target.load('address');
      await context.sync();
      await opConditional(context, area.sheet, { address: target.address, rule, value, fill, fontColor });
      return { message: `✅ Penandaan otomatis dipasang pada ${target.address}.` };
    });
  },

  async dedupe() {
    return withArea(async (context, area) => {
      const cols = area.values[0].map((_, i) => i);
      const res = area.range.removeDuplicates(cols, area.hasHeader || area.values[0].every((v) => typeof v === 'string'));
      res.load('removed,uniqueRemaining');
      await context.sync();
      return { message: `✅ ${res.removed} baris duplikat dihapus, ${res.uniqueRemaining} baris tersisa.` };
    });
  },

  async percentOfTotal() {
    return withArea(async (context, area) => {
      const cols = numericColumns(area);
      if (!cols.length) throw new Error('Tidak ada kolom angka.');
      const j = cols[cols.length - 1];
      const L = columnLetter(area.col + j);
      const start = area.hasHeader ? area.row + 1 : area.row;
      const end = area.row + area.values.length - 1;
      const outCol = columnLetter(area.col + area.values[0].length);
      const vals = [];
      if (area.hasHeader) vals.push(['% dari Total']);
      for (let r = start; r <= end; r++) vals.push([`=IF(SUM($${L}$${start}:$${L}$${end})=0,0,${L}${r}/SUM($${L}$${start}:$${L}$${end}))`]);
      const out = area.sheet.getRange(`${outCol}${area.row}:${outCol}${end}`);
      out.formulas = vals;
      out.numberFormat = vals.map((_, i) => [area.hasHeader && i === 0 ? 'General' : FORMAT_PERSEN]);
      if (area.hasHeader) area.sheet.getRange(`${outCol}${area.row}`).format.font.bold = true;
      return { message: `✅ Kolom persentase dari total ditambahkan di kolom ${outCol}.` };
    });
  },

  async rank({ ascending }) {
    return withArea(async (context, area) => {
      const cols = numericColumns(area);
      if (!cols.length) throw new Error('Tidak ada kolom angka.');
      const L = columnLetter(area.col + cols[cols.length - 1]);
      const start = area.hasHeader ? area.row + 1 : area.row;
      const end = area.row + area.values.length - 1;
      const outCol = columnLetter(area.col + area.values[0].length);
      const vals = area.hasHeader ? [['Peringkat']] : [];
      for (let r = start; r <= end; r++) vals.push([`=IF(${L}${r}="","",RANK.EQ(${L}${r},$${L}$${start}:$${L}$${end},${ascending ? 1 : 0}))`]);
      area.sheet.getRange(`${outCol}${area.row}:${outCol}${end}`).formulas = vals;
      if (area.hasHeader) area.sheet.getRange(`${outCol}${area.row}`).format.font.bold = true;
      return { message: `✅ Kolom peringkat ditambahkan di kolom ${outCol}.` };
    });
  },

  async numbering() {
    return withArea(async (context, area) => {
      const start = area.hasHeader ? 1 : 0;
      const n = area.values.length - start;
      const firstHeader = String(area.values[0][0] || '');
      const fillExisting = area.source === 'seleksi' || /^(no\.?|nomor|#)$/i.test(firstHeader.trim());
      if (fillExisting) {
        area.range.getCell(start, 0).getResizedRange(n - 1, 0).values = Array.from({ length: n }, (_, i) => [i + 1]);
        return { message: `✅ Nomor urut 1–${n} diisi di kolom ${columnLetter(area.col)}.` };
      }
      // Jangan menimpa data: sisipkan kolom "No" baru di kiri data.
      const col = area.range.getCell(0, 0).getResizedRange(area.values.length - 1, 0);
      col.insert('Right');
      const fresh = area.sheet.getRange(cellAddress(area.row, area.col)).getResizedRange(area.values.length - 1, 0);
      fresh.values = (area.hasHeader ? [['No']] : []).concat(Array.from({ length: n }, (_, i) => [i + 1]));
      fresh.format.horizontalAlignment = 'Center';
      if (area.hasHeader) fresh.getCell(0, 0).format.font.bold = true;
      return { message: `✅ Kolom "No" (1–${n}) disisipkan di kolom ${columnLetter(area.col)}; data lama bergeser ke kanan.` };
    });
  },

  async groupSum({ column, value }) {
    return withArea(async (context, area) => {
      if (!area.hasHeader) throw new Error('Data perlu baris judul (header) untuk direkap.');
      let g = findColumn(area, column);
      if (g < 0) g = area.values[0].findIndex((_, j) => !numericColumns(area).includes(j));
      if (g < 0) throw new Error('Kolom pengelompokan tidak ditemukan. Contoh: "rekap total per Produk".');
      const nums = numericColumns(area).filter((j) => j !== g);
      let v = findColumn(area, value);
      if (v < 0 || !nums.includes(v)) v = nums[nums.length - 1];
      if (v === undefined) throw new Error('Kolom angka untuk dijumlahkan tidak ditemukan.');
      const groups = [];
      for (const r of area.values.slice(1)) {
        const k = r[g];
        if (!isEmpty(k) && !groups.includes(k)) groups.push(k);
      }
      if (!groups.length) throw new Error('Tidak ada data untuk direkap.');
      const start = area.row + 1;
      const end = area.row + area.values.length - 1;
      const GL = columnLetter(area.col + g);
      const VL = columnLetter(area.col + v);
      const oc = area.col + area.values[0].length + 2;
      const OL = columnLetter(oc);
      const vals = [[`Rekap per ${area.values[0][g]}`, `Total ${area.values[0][v]}`]];
      groups.forEach((k, i) => vals.push([k, `=SUMIF($${GL}$${start}:$${GL}$${end},${OL}${area.row + 1 + i},$${VL}$${start}:$${VL}$${end})`]));
      vals.push(['TOTAL', `=SUM(${columnLetter(oc + 1)}${area.row + 1}:${columnLetter(oc + 1)}${area.row + groups.length})`]);
      const ops = [
        { op: 'write', address: cellAddress(area.row, oc), values: vals },
        { op: 'format', address: `${cellAddress(area.row, oc)}:${cellAddress(area.row, oc + 1)}`, bold: true, fill: '#1F3864', fontColor: '#FFFFFF' },
        { op: 'format', address: `${cellAddress(area.row + groups.length + 1, oc)}:${cellAddress(area.row + groups.length + 1, oc + 1)}`, bold: true, fill: '#D9E1F2', border: 'top-double' },
        { op: 'format', address: `${cellAddress(area.row, oc)}:${cellAddress(area.row + groups.length + 1, oc + 1)}`, border: 'all' },
        { op: 'numberFormat', address: `${cellAddress(area.row + 1, oc + 1)}:${cellAddress(area.row + groups.length + 1, oc + 1)}`, format: area.numberFormat[1] ? area.numberFormat[1][v] : FORMAT_ANGKA },
        { op: 'autofit', address: `${cellAddress(area.row, oc)}:${cellAddress(area.row + groups.length + 1, oc + 1)}` },
        { op: 'chart', address: `${cellAddress(area.row, oc)}:${cellAddress(area.row + groups.length, oc + 1)}`, type: groups.length <= 6 ? 'Pie' : 'BarClustered', title: vals[0][1] + ' per ' + area.values[0][g], position: cellAddress(area.row, oc + 3) },
      ];
      const res = await runOps(context, area.sheet, ops);
      return { message: `✅ Rekap ${groups.length} kelompok "${area.values[0][g]}" dibuat di kolom ${OL} beserta grafik.`, errors: res.errors };
    });
  },

  async terbilang({ number }) {
    return Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.load('values,address,rowCount,columnCount,rowIndex,columnIndex');
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      await context.sync();
      if (number !== null && number !== undefined) {
        const cell = sel.getCell(0, 0);
        cell.values = [[terbilangRupiah(number)]];
        return { message: `✅ ${terbilangRupiah(number)}` };
      }
      const nums = [];
      sel.values.forEach((row, i) => row.forEach((v, j) => isNum(v) && nums.push({ i, j, v })));
      if (!nums.length) throw new Error('Pilih sel berisi angka, atau ketik "terbilang 1.500.000".');
      if (nums.length > 500) throw new Error('Maksimal 500 sel sekaligus.');
      const outCol = sel.columnIndex + sel.columnCount;
      for (const { i, j, v } of nums) {
        const src = cellAddress(sel.rowIndex + 1 + i, sel.columnIndex + 1 + j);
        const tgt = cellAddress(sel.rowIndex + 1 + i, outCol + 1 + (sel.columnCount > 1 ? j : 0));
        sheet.getRange(tgt).values = [[terbilangRupiah(v)]];
        rememberTerbilang({ sheet: sheet.name, source: src, target: tgt, rupiah: true });
      }
      return { message: `✅ Terbilang ditulis untuk ${nums.length} angka di kolom ${columnLetter(outCol + 1)}. Ketik "perbarui terbilang" bila angkanya berubah.` };
    });
  },

  async refreshTerbilang() {
    let list = [];
    try {
      list = Office.context.document.settings.get('master.terbilang') || [];
    } catch (e) {
      list = [];
    }
    if (!list.length) return { message: 'ℹ️ Belum ada sel terbilang yang tercatat di dokumen ini.' };
    return Excel.run(async (context) => {
      let n = 0;
      for (const e of list) {
        const ws = context.workbook.worksheets.getItemOrNullObject(e.sheet);
        await context.sync();
        if (ws.isNullObject) continue;
        const src = ws.getRange(e.source);
        src.load('values');
        await context.sync();
        const v = Number(src.values[0][0]) || 0;
        ws.getRange(e.target).values = [[e.rupiah ? terbilangRupiah(v) : terbilang(v)]];
        n++;
      }
      return { message: `✅ ${n} sel terbilang diperbarui.` };
    });
  },

  async newSheet({ name }) {
    return Excel.run(async (context) => {
      const nm = await uniqueSheetName(context, name ? titleCase(name) : 'Sheet Baru');
      context.workbook.worksheets.add(nm).activate();
      return { message: `✅ Sheet "${nm}" dibuat.` };
    });
  },

  async transpose() {
    return withArea(async (context, area) => {
      const t = area.values[0].map((_, j) => area.values.map((r) => r[j]));
      const r = area.row + area.values.length + 2;
      area.sheet.getRange(cellAddress(r, area.col)).getResizedRange(t.length - 1, t[0].length - 1).values = t;
      return { message: `✅ Hasil transpos ditulis mulai ${cellAddress(r, area.col)}.` };
    });
  },

  async removeEmptyRows() {
    return withArea(async (context, area) => {
      let n = 0;
      for (let i = area.values.length - 1; i >= 0; i--) {
        if (area.values[i].every(isEmpty)) {
          area.range.getRow(i).delete('Up');
          n++;
        }
      }
      return { message: n ? `✅ ${n} baris kosong dihapus.` : 'ℹ️ Tidak ada baris kosong.' };
    });
  },

  async freeze() {
    return withArea(async (context, area) => {
      const rows = area.hasHeader ? area.row : Math.max(1, area.row);
      area.sheet.freezePanes.freezeRows(rows);
      return { message: `✅ Baris 1–${rows} dibekukan.` };
    });
  },

  async unfreeze() {
    return Excel.run(async (context) => {
      context.workbook.worksheets.getActiveWorksheet().freezePanes.unfreeze();
      return { message: '✅ Bekuan baris/kolom dilepas.' };
    });
  },

  async trim() {
    return transformText((s) => s.replace(/\s+/g, ' ').trim(), 'Spasi berlebih dirapikan');
  },

  async case({ mode }) {
    const fn = { upper: (s) => s.toUpperCase(), lower: (s) => s.toLowerCase(), title: titleCase, sentence: (s) => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() }[mode];
    return transformText(fn, 'Huruf diubah');
  },

  async border() {
    return withArea(async (context, area) => {
      area.range.load('address');
      await context.sync();
      await opFormat(context, area.sheet, { address: area.range.address.split('!').pop(), border: 'all' });
      return { message: `✅ Garis tabel dipasang pada ${area.range.address}.` };
    });
  },

  async bold({ headerOnly }) {
    return withArea(async (context, area) => {
      const target = headerOnly ? area.range.getRow(0) : area.range;
      target.format.font.bold = true;
      if (headerOnly) target.format.fill.color = '#D9E1F2';
      return { message: '✅ Teks ditebalkan.' };
    });
  },

  async clearFormat() {
    return Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.clear('Formats');
      return { message: '✅ Format sel terpilih dihapus.' };
    });
  },

  async autofit() {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      await opAutofit(context, sheet, {});
      return { message: '✅ Lebar kolom dirapikan.' };
    });
  },

  async today() {
    return Excel.run(async (context) => {
      const d = new Date();
      const cell = context.workbook.getSelectedRange().getCell(0, 0);
      cell.formulas = [[`=DATE(${d.getFullYear()},${d.getMonth() + 1},${d.getDate()})`]];
      cell.numberFormat = [[FORMAT_TANGGAL]];
      return { message: '✅ Tanggal hari ini diisi.' };
    });
  },
};

async function transformText(fn, label) {
  return withArea(async (context, area) => {
    let n = 0;
    const out = area.formulas.map((row) =>
      row.map((f) => {
        if (typeof f === 'string' && f && !f.startsWith('=')) {
          const t = fn(f);
          if (t !== f) n++;
          return t;
        }
        return f;
      }),
    );
    area.range.formulas = out;
    return { message: `✅ ${label} pada ${n} sel.` };
  });
}

export async function runExcelAction(action, params, env = {}) {
  const name = action.replace(/^excel\./, '');
  const fn = ACTIONS[name];
  if (!fn) throw new Error(`Aksi Excel "${action}" belum tersedia`);
  return fn(params || {}, env);
}

/* ---------------------------- alat untuk AI ---------------------------- */

export async function getWorkbookInfo() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name,items/visibility');
    const active = sheets.getActiveWorksheet();
    active.load('name');
    const sel = context.workbook.getSelectedRange();
    sel.load('address,values,rowCount,columnCount');
    const tables = context.workbook.tables;
    tables.load('items/name');
    await context.sync();
    const used = sheets.items.map((s) => {
      const u = s.getUsedRangeOrNullObject(true);
      u.load('address,rowCount,columnCount');
      return { s, u };
    });
    const tranges = tables.items.map((t) => {
      const r = t.getRange();
      r.load('address');
      return { t, r };
    });
    await context.sync();
    const preview = sel.rowCount * sel.columnCount <= 400 ? sel.values : sel.values.slice(0, 20).map((r) => r.slice(0, 20));
    return {
      activeSheet: active.name,
      selection: { address: sel.address, rows: sel.rowCount, columns: sel.columnCount, values: preview },
      sheets: used
        .filter(({ s }) => s.name !== LIST_SHEET)
        .map(({ s, u }) => ({ name: s.name, hidden: s.visibility !== 'Visible', usedRange: u.isNullObject ? null : u.address, rows: u.isNullObject ? 0 : u.rowCount, columns: u.isNullObject ? 0 : u.columnCount })),
      tables: tranges.map(({ t, r }) => ({ name: t.name, address: r.address })),
    };
  });
}

export async function readRange({ sheet, address, maxCells = 3000, includeFormulas = true }) {
  return Excel.run(async (context) => {
    const ws = sheet ? context.workbook.worksheets.getItem(sheet) : context.workbook.worksheets.getActiveWorksheet();
    let range;
    if (address) range = rangeOf(context, ws, address);
    else {
      range = ws.getUsedRangeOrNullObject(true);
      await context.sync();
      if (range.isNullObject) return { address: null, values: [], note: 'Sheet kosong' };
    }
    range.load('address,rowCount,columnCount');
    await context.sync();
    const cap = clamp(Number(maxCells) || 3000, 1, 20000);
    let target = range;
    let truncated = false;
    if (range.rowCount * range.columnCount > cap) {
      const rows = Math.max(1, Math.floor(cap / range.columnCount));
      target = range.getCell(0, 0).getResizedRange(Math.min(rows, range.rowCount) - 1, Math.min(range.columnCount, cap) - 1);
      truncated = true;
    }
    target.load('address,values,formulas,numberFormat');
    await context.sync();
    const out = { address: target.address, fullAddress: range.address, rows: range.rowCount, columns: range.columnCount, values: target.values, truncated };
    if (includeFormulas) {
      const f = {};
      target.formulas.forEach((row, i) =>
        row.forEach((v, j) => {
          if (typeof v === 'string' && v.startsWith('=')) f[cellAddress(i + 1, j + 1)] = v;
        }),
      );
      const keys = Object.keys(f);
      if (keys.length) out.formulasRelative = keys.length > 300 ? Object.fromEntries(keys.slice(0, 300).map((k) => [k, f[k]])) : f;
    }
    return out;
  });
}

export async function applyOps({ sheet, ops }) {
  return applyPlan({ sheetName: sheet || null, newSheet: false, ops });
}

export async function runCustomExcel(code) {
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
  return Excel.run(async (context) => {
    const fn = new AsyncFunction('context', 'Excel', 'Office', code);
    const result = await fn(context, Excel, Office);
    await context.sync();
    return result;
  });
}
