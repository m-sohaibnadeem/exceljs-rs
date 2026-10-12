'use strict';

// The ExcelJS bugs excel-rs fixes (COMPATIBILITY.md, "Bugs fixed from
// ExcelJS", B1-B14). The fixture snapshots hold the fixed behaviour;
// these show that it is right, mostly by round trips through both writers and
// both readers. Also D6, a deliberate difference of the read pipeline.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ExcelJS = require('..');

const FIXTURES = path.join(__dirname, '..', '..', '..', 'fixtures');

function tmpFile(name) {
  return path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'excel-rs-')), name);
}

// workbook.xlsx.load: {address: value} of the first worksheet.
async function readValues(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const out = {};
  wb.worksheets[0].eachRow(row => {
    row.eachCell(cell => {
      out[cell.address] = cell.value;
    });
  });
  return {wb, values: out};
}

// stream.xlsx.WorkbookReader: {address: value} of the first worksheet.
async function streamValues(file, options = {}) {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, {sharedStrings: 'cache', ...options});
  const out = {};
  for await (const ws of reader) {
    for await (const row of ws) {
      row.eachCell(cell => {
        out[cell.address] = cell.value;
      });
    }
    break;
  }
  return out;
}

// Writes `rows` (arrays of values) with the streaming writer.
async function streamWrite(rows, options) {
  const filename = tmpFile('stream.xlsx');
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename, ...options});
  const ws = wb.addWorksheet('S');
  for (const r of rows) ws.addRow(r).commit();
  await wb.commit();
  return filename;
}

const TRICKY = ['_x0041_ literal', 'a_x005F_b', '_x0041__x0042_', 'é_xabcd_', 'x &amp; &lt;'];

test('B1/B2: text shaped like escapes or entities round-trips (xlsx writer, both string modes)', async () => {
  for (const useSharedStrings of [true, false]) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('a &amp; b');
    TRICKY.forEach((s, i) => {
      ws.getCell(1, i + 1).value = s;
    });
    ws.getCell('A2').value = {richText: [{text: '_x0042_ run'}]};
    ws.getCell('B2').value = {formula: '"x"', result: '_x0043_ result'};
    const buffer = await wb.xlsx.writeBuffer({useSharedStrings});
    const {wb: back, values} = await readValues(buffer);
    assert.equal(back.worksheets[0].name, 'a &amp; b');
    TRICKY.forEach((s, i) => assert.equal(values[`${String.fromCharCode(65 + i)}1`], s));
    assert.deepEqual(values.A2, {richText: [{text: '_x0042_ run'}]});
    assert.deepEqual(values.B2, {formula: '"x"', result: '_x0043_ result'});
  }
});

test('B1/B2/B3: the streaming writer round-trips the same text, and the empty string', async () => {
  for (const useSharedStrings of [true, false]) {
    const file = await streamWrite([[...TRICKY, '']], {useSharedStrings});
    const expected = {};
    [...TRICKY, ''].forEach((s, i) => {
      expected[`${String.fromCharCode(65 + i)}1`] = s;
    });
    const {values} = await readValues(fs.readFileSync(file));
    assert.deepEqual(values, expected, `xlsx.load, useSharedStrings ${useSharedStrings}`);
    assert.deepEqual(await streamValues(file), expected, `stream reader, useSharedStrings ${useSharedStrings}`);
  }
});

test('B3: xlsx.writeBuffer honours useSharedStrings: false', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').addRow(['a', '', {richText: [{text: 'r'}]}]);
  const buffer = await wb.xlsx.writeBuffer({useSharedStrings: false});
  // ZIP entry names are stored uncompressed.
  assert.ok(!buffer.includes('xl/sharedStrings.xml'), 'no shared strings part');
  assert.ok(buffer.includes('xl/worksheets/sheet1.xml'));
  const {values} = await readValues(buffer);
  assert.deepEqual(values, {A1: 'a', B1: '', C1: {richText: [{text: 'r'}]}});
});

test('B4/B12: phonetic runs are not part of the text (both readers)', async () => {
  const file = path.join(FIXTURES, 'edge-inline-phonetic.xlsx');
  const {values} = await readValues(fs.readFileSync(file));
  assert.deepEqual(values, {A1: 1, B1: 'base', C1: 3, A2: 4});
  assert.deepEqual(await streamValues(file), {A1: 1, B1: 'base', C1: 3, A2: 4});
  // <si> with <rPh>, read by the stream reader's shared strings
  const sst = await streamValues(path.join(FIXTURES, 'edge-cells.xlsx'));
  assert.equal(sst.C1, 'base');
});

test('B5: <si/> is the empty string, and writes back as one', async () => {
  const {wb, values} = await readValues(fs.readFileSync(path.join(FIXTURES, 'edge-sst-empty-si.xlsx')));
  assert.equal(values.A1, '');
  assert.equal((await readValues(await wb.xlsx.writeBuffer())).values.A1, '');
});

test('B6: a shared-string cell without shared strings is empty, and the file loads', async () => {
  const {wb} = await readValues(fs.readFileSync(path.join(FIXTURES, 'edge-no-shared-strings.xlsx')));
  assert.equal(wb.worksheets[0].getCell('A1').type, ExcelJS.ValueType.Null);
});

test('B7: a formula result of 0, false or "" is kept', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('F');
  ws.getCell('A1').value = {formula: '1-1', result: 0};
  ws.getCell('A2').value = {formula: '1=2', result: false};
  ws.getCell('A3').value = {formula: '""', result: ''};
  assert.equal(ws.getCell('A1').value.result, 0);
  assert.equal(ws.getCell('A2').value.result, false);
  assert.equal(ws.getCell('A3').value.result, '');
  const {values} = await readValues(await wb.xlsx.writeBuffer());
  assert.deepEqual([values.A1, values.A2, values.A3], [
    {formula: '1-1', result: 0},
    {formula: '1=2', result: false},
    {formula: '""', result: ''},
  ]);
  const read = await readValues(fs.readFileSync(path.join(FIXTURES, 'formulas.xlsx')));
  assert.deepEqual(read.values.C3, {sharedFormula: 'C1', result: 0});
});

test('B8/B10: CDATA text and boolean words (both readers)', async () => {
  const file = path.join(FIXTURES, 'edge-cells.xlsx');
  const {values} = await readValues(fs.readFileSync(file));
  const stream = await streamValues(file);
  for (const v of [values, stream]) {
    assert.equal(v.E1, 'cdata text');
    assert.deepEqual([v.A2, v.B2, v.C2, v.D2, v.E2, v.F2], [true, false, true, true, false, false]);
  }
});

test('B9: a cell is addressed by the row it is in', async () => {
  const {wb} = await readValues(fs.readFileSync(path.join(FIXTURES, 'edge-cells.xlsx')));
  const ws = wb.worksheets[0];
  const cell = ws.getRow(8).getCell(1);
  assert.equal(cell.address, 'A8');
  assert.equal(ws.getCell('A8'), cell);
});

test('B11: the streaming writer keeps distinct rich texts apart', async () => {
  const a = {richText: [{text: 'first', font: {bold: true}}]};
  const b = {richText: [{text: 'second'}]};
  const file = await streamWrite([[a, b, a]], {useSharedStrings: true});
  const {values} = await readValues(fs.readFileSync(file));
  assert.deepEqual(
    [values.A1, values.B1, values.C1].map(v => v.richText.map(r => r.text).join('')),
    ['first', 'second', 'first']
  );
});

test('B12: the stream reader reads shared-string fonts like the xlsx reader', async () => {
  const file = path.join(FIXTURES, 'edge-sst-rich-fonts.xlsx');
  const v = await streamValues(file);
  assert.deepEqual(v.A1, {
    richText: [
      {font: {name: 'Arial', color: {argb: 'FFFF0000'}, size: 11}, text: 'red'},
      {font: {color: {theme: 1}}, text: ' themed'},
    ],
  });
  assert.equal(v.B1, 'spaced');
  const sst = await streamValues(path.join(FIXTURES, 'edge-cells.xlsx'));
  assert.equal(sst.B1, 'Ab upper and j lower hex escapes');
});

test('B13: the stream reader with styles "ignore" reads a row with s="0"', async () => {
  const file = path.join(FIXTURES, 'edge-row-style-zero.xlsx');
  assert.deepEqual(await streamValues(file, {styles: 'ignore'}), {A1: 1});
});

test('D6: a load that fails at a cell has replaced the model already', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Mine');
  await assert.rejects(
    wb.xlsx.readFile(path.join(FIXTURES, 'edge-sst-index-out-of-range.xlsx')),
    /Cannot read properties of undefined \(reading 'richText'\)/
  );
  assert.deepEqual(
    wb.worksheets.map(ws => ws.name),
    ['Edge']
  );
});
