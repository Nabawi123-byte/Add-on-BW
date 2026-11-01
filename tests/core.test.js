import { test } from 'node:test';
import assert from 'node:assert/strict';
import { terbilang, terbilangRupiah } from '../src/core/terbilang.js';
import { parseNumber, parsePercent, columnLetter, parseAddress, parseKeyValues, splitList } from '../src/core/util.js';

test('terbilang', () => {
  assert.equal(terbilang(0), 'nol');
  assert.equal(terbilang(11), 'sebelas');
  assert.equal(terbilang(15), 'lima belas');
  assert.equal(terbilang(100), 'seratus');
  assert.equal(terbilang(1000), 'seribu');
  assert.equal(terbilang(1500000), 'satu juta lima ratus ribu');
  assert.equal(terbilang(2025), 'dua ribu dua puluh lima');
  assert.equal(terbilang(111111), 'seratus sebelas ribu seratus sebelas');
  assert.equal(terbilang(1001000), 'satu juta seribu');
  assert.equal(terbilang(7500000000), 'tujuh miliar lima ratus juta');
  assert.equal(terbilang(12.5), 'dua belas koma lima');
  assert.equal(terbilang(-3), 'minus tiga');
  assert.equal(terbilangRupiah(250000), 'Dua ratus lima puluh ribu rupiah');
});

test('parseNumber gaya Indonesia', () => {
  assert.equal(parseNumber('Rp 1.500.000'), 1500000);
  assert.equal(parseNumber('1,5 juta'), 1500000);
  assert.equal(parseNumber('50jt'), 50000000);
  assert.equal(parseNumber('2.5'), 2.5);
  assert.equal(parseNumber('1.500'), 1500);
  assert.equal(parseNumber('1.234.567,89'), 1234567.89);
  assert.equal(parseNumber('300 ribu'), 300000);
  assert.equal(parseNumber('2 miliar'), 2e9);
  assert.equal(parseNumber('abc'), null);
  assert.equal(parsePercent('10%'), 0.1);
  assert.equal(parsePercent('7,5%'), 0.075);
});

test('alamat sel', () => {
  assert.equal(columnLetter(1), 'A');
  assert.equal(columnLetter(26), 'Z');
  assert.equal(columnLetter(27), 'AA');
  assert.deepEqual(parseAddress("'Data 1'!B3:D10"), { sheet: 'Data 1', row: 3, col: 2, lastRow: 10, lastCol: 4 });
  assert.deepEqual(parseAddress('C5'), { sheet: null, row: 5, col: 3, lastRow: 5, lastCol: 3 });
});

test('key values & list', () => {
  assert.deepEqual(parseKeyValues('surat izin nama: Budi Santoso, jabatan: Staf IT; alasan = sakit demam'), { surat_izin_nama: 'Budi Santoso', jabatan: 'Staf IT', alasan: 'sakit demam' });
  assert.deepEqual(splitList('nama, jabatan dan gaji'), ['nama', 'jabatan', 'gaji']);
});
