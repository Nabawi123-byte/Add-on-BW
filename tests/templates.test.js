import { test } from 'node:test';
import assert from 'node:assert/strict';
import { EXCEL_TEMPLATES } from '../src/core/excel-templates.js';
import { WORD_TEMPLATES, templateDefaults } from '../src/core/word-templates.js';
import { renderSegments, blocksToText } from '../src/core/word-render.js';
import { parseAddress } from '../src/core/util.js';

const OPS = new Set(['write', 'format', 'table', 'chart', 'conditional', 'freeze', 'autofit', 'select', 'validation', 'terbilang']);

test('semua template Excel menghasilkan operasi yang valid', () => {
  for (const tpl of EXCEL_TEMPLATES) {
    for (const params of [{}, { baris: 3, tenor: 6, metode: 'flat', bulan: 1, tahun: 2028 }]) {
      const plan = tpl.build(params, { instansi: 'PT Uji', nama: 'Andi' }).toJSON();
      assert.ok(plan.sheetName, tpl.id);
      assert.ok(plan.ops.length > 3, tpl.id);
      const json = JSON.stringify(plan);
      assert.ok(!/undefined|NaN/.test(json), `${tpl.id}: ada undefined/NaN`);
      const tables = plan.ops.filter((o) => o.op === 'table');
      assert.ok(tables.length >= 0);
      for (const op of plan.ops) {
        assert.ok(OPS.has(op.op), `${tpl.id}: op ${op.op}`);
        if (op.address) assert.ok(parseAddress(op.address), `${tpl.id}: alamat ${op.address}`);
        if (op.op === 'write') {
          assert.ok(Array.isArray(op.values) && op.values.every((r) => Array.isArray(r)), `${tpl.id}: values 2D`);
          const w = op.values[0].length;
          assert.ok(op.values.every((r) => r.length === w), `${tpl.id}: baris tidak rata di ${op.address}`);
          const a = parseAddress(op.address);
          if (a.lastRow !== a.row || a.lastCol !== a.col) {
            assert.equal(a.lastRow - a.row + 1, op.values.length, `${tpl.id}: tinggi ${op.address}`);
            assert.equal(a.lastCol - a.col + 1, w, `${tpl.id}: lebar ${op.address}`);
          }
          for (const row of op.values) for (const v of row) {
            if (typeof v === 'string' && v.startsWith('=')) {
              const open = (v.match(/\(/g) || []).length;
              const close = (v.match(/\)/g) || []).length;
              assert.equal(open, close, `${tpl.id}: kurung tidak seimbang ${v}`);
            }
          }
        }
      }
    }
  }
});

test('cicilan anuitas: rumus PMT dan baris sesuai tenor', () => {
  const tpl = EXCEL_TEMPLATES.find((t) => t.id === 'cicilan');
  const plan = tpl.build({ pokok: 12000000, bunga: 0.12, tenor: 12 }).toJSON();
  const json = JSON.stringify(plan);
  assert.ok(json.includes('PMT($B$4/12,$B$5,-$B$3)'));
  const table = plan.ops.find((o) => o.op === 'table');
  assert.equal(table.address, 'A10:E22');
});

test('semua template Word ter-render tanpa nilai kosong yang rusak', () => {
  const profile = { nama: 'Budi', jabatan: 'Manajer', instansi: 'PT Maju', alamat: 'Jl. Merdeka 1', kota: 'Bandung' };
  for (const tpl of WORD_TEMPLATES) {
    const v = templateDefaults(tpl, profile, { jumlah: '1.500.000', item: 'Laptop | 2 | 7.000.000', petugas: 'Andi, Sari' });
    const blocks = tpl.build(v);
    const html = renderSegments(blocks).join('');
    assert.ok(html.length > 200, tpl.id);
    assert.ok(!/undefined|NaN|\[object/.test(html), `${tpl.id}: ${html.match(/.{40}(undefined|NaN|\[object).{20}/)}`);
    assert.ok(blocksToText(blocks).length > 20, tpl.id);
  }
  const kw = WORD_TEMPLATES.find((t) => t.id === 'kwitansi');
  const html = renderSegments(kw.build(templateDefaults(kw, profile, { jumlah: '1.500.000' }))).join('');
  assert.ok(html.includes('Satu juta lima ratus ribu rupiah'));
  assert.ok(html.includes('Rp 1.500.000'));
});

test('template Word mengambil nilai dari perintah', () => {
  const tpl = WORD_TEMPLATES.find((t) => t.id === 'izin');
  const v = templateDefaults(tpl, { nama: 'Profil' }, { surat_izin_nama: 'Budi', alasan: 'demam' });
  assert.equal(v.nama, 'Budi');
  assert.equal(v.alasan, 'demam');
});

test('kunci mirip tidak menimpa field lain', () => {
  const tpl = WORD_TEMPLATES.find((t) => t.id === 'undangan');
  const v = templateDefaults(tpl, {}, { hari_tanggal: 'Senin, 3 November 2026' });
  assert.equal(v.hari_tanggal, 'Senin, 3 November 2026');
  assert.notEqual(v.tanggal, 'Senin, 3 November 2026');
});
