// Pembuat "rencana" (plan) untuk Excel: daftar operasi yang nanti dijalankan oleh executor Office.js.
// Template, perintah offline, dan AI memakai kosakata operasi yang sama.

import { columnLetter, FORMAT_RUPIAH } from './util.js';

export const TABLE_STYLE = 'TableStyleMedium2';

export class SheetPlan {
  constructor(sheetName, { newSheet = true } = {}) {
    this.sheetName = sheetName;
    this.newSheet = newSheet;
    this.ops = [];
    this.summary = '';
  }

  push(op) {
    this.ops.push(op);
    return this;
  }

  write(address, values) {
    return this.push({ op: 'write', address, values });
  }

  format(address, opts) {
    return this.push(Object.assign({ op: 'format', address }, opts));
  }

  /** Judul besar di baris 1 dan subjudul di baris 2. */
  title(text, subtitle, width = 6) {
    const last = columnLetter(width);
    this.write('A1', [[text]]);
    this.format(`A1:${last}1`, { bold: true, fontSize: 16, fontColor: '#1F3864', merge: true, hAlign: 'Center' });
    if (subtitle) {
      this.write('A2', [[subtitle]]);
      this.format(`A2:${last}2`, { italic: true, fontColor: '#595959', merge: true, hAlign: 'Center' });
    }
    return this;
  }

  /** Blok label-nilai (misal parameter input) mulai dari baris tertentu di kolom A-B. */
  inputs(startRow, items) {
    const refs = {};
    items.forEach((it, i) => {
      const r = startRow + i;
      this.write(`A${r}:B${r}`, [[it.label, it.value]]);
      this.format(`A${r}`, { bold: true });
      const style = { numberFormat: it.format || 'General' };
      if (it.input !== false) Object.assign(style, { fill: '#FFF2CC', border: 'outline' });
      else Object.assign(style, { bold: true, fill: '#E2EFDA', border: 'outline' });
      this.format(`B${r}`, style);
      refs[it.key] = `$B$${r}`;
    });
    return refs;
  }

  /**
   * Tabel data.
   * headers: string[]
   * count: jumlah baris data
   * row(r, i, C): mengembalikan isi baris ke-i (0-based) di baris Excel r; C(namaHeader) -> huruf kolom
   * formats: { namaHeader: numberFormat }
   * totals: { namaHeader: 'SUM'|'AVERAGE'|'COUNTA'|'MAX'|'MIN' }
   */
  table({ startRow, startCol = 1, headers, count, row, formats = {}, totals = null, totalLabel = 'TOTAL', name, style = TABLE_STYLE, widths = {} }) {
    const C = (h) => {
      const idx = headers.indexOf(h);
      if (idx < 0) throw new Error(`Kolom "${h}" tidak ada`);
      return columnLetter(startCol + idx);
    };
    const firstCol = columnLetter(startCol);
    const lastCol = columnLetter(startCol + headers.length - 1);
    const headerRow = startRow;
    const first = startRow + 1;
    const last = startRow + count;
    const values = [headers.slice()];
    for (let i = 0; i < count; i++) {
      const r = first + i;
      const cells = row ? row(r, i, C) : headers.map(() => '');
      values.push(headers.map((_, j) => (cells[j] === undefined || cells[j] === null ? '' : cells[j])));
    }
    this.write(`${firstCol}${headerRow}`, values);
    for (const [h, fmt] of Object.entries(formats)) {
      this.format(`${C(h)}${first}:${C(h)}${last + (totals ? 1 : 0)}`, { numberFormat: fmt });
    }
    for (const [h, w] of Object.entries(widths)) {
      this.format(`${C(h)}${headerRow}`, { columnWidth: w });
    }
    let totalRow = null;
    if (totals) {
      totalRow = last + 1;
      const t = headers.map((h, j) => {
        if (totals[h]) return `=${totals[h]}(${C(h)}${first}:${C(h)}${last})`;
        return j === 0 ? totalLabel : '';
      });
      this.write(`${firstCol}${totalRow}`, [t]);
      this.format(`${firstCol}${totalRow}:${lastCol}${totalRow}`, { bold: true, fill: '#D9E1F2', border: 'top-double' });
    }
    // Tabel dibuat terakhir supaya baris total di bawahnya tidak ikut masuk ke tabel.
    this.push({ op: 'table', address: `${firstCol}${headerRow}:${lastCol}${last}`, name, style, headers: true });
    return {
      C,
      headerRow,
      first,
      last,
      totalRow,
      firstCol,
      lastCol,
      address: `${firstCol}${headerRow}:${lastCol}${last}`,
      col: (h) => `${C(h)}${first}:${C(h)}${last}`,
      total: (h) => `${C(h)}${totalRow}`,
    };
  }

  finish({ freezeRows = 0, autofit = true, select = 'A1' } = {}) {
    if (autofit) this.push({ op: 'autofit' });
    if (freezeRows) this.push({ op: 'freeze', rows: freezeRows });
    if (select) this.push({ op: 'select', address: select });
    return this;
  }

  toJSON() {
    return { sheetName: this.sheetName, newSheet: this.newSheet, ops: this.ops, summary: this.summary };
  }
}

export const RP = FORMAT_RUPIAH;
